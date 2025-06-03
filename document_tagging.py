#!/usr/bin/env python
# document_tagging.py
# ──────────────────────────────────────────────────────────────────────────────
"""
Batch-extract parameters from PDFs with both “simple” (fast) and “analytical”
(reasoning / large-context) LLM pipelines.

Main changes vs. the previous version
──────────────────────────────────────
1.  **Hard context limit** for o-series models (`CONTEXT_O1 – SAFETY_MARGIN`).
2.  **Exponential back-off** on any error that looks transient
    (`RateLimit`, `APIConnection`, etc.).
3.  **Lower default concurrency** (`MAX_WORKERS = 3`).
4.  **Error rows are kept** in the final Excel: `Answer` is `"ERROR: …"` so
    blanks now mean *“LLM said the value is empty”*, not *“the call died”*.
5.  **Verbose on-screen logging** with the exception class.

Copy-paste and run; the CLI flags are unchanged.
"""
# ──────────────────────────────────────────────────────────────────────────────

from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import tiktoken
from langchain.schema import HumanMessage, SystemMessage
from langchain_community.document_loaders import PyPDFLoader
from langchain_anthropic import ChatAnthropic
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_openai import ChatOpenAI
from openai import OpenAI

# ---------------------------------------------------------------------------
# Configuration & constants
# ---------------------------------------------------------------------------
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

INPUT_FOLDER_ANALYTICAL = "input_taggedstudies"
INPUT_FOLDER_SIMPLE = "input_taggedstudies"

QUESTIONS_FILE_ANALYTICAL = "questions_v3v2_hugo.xlsx"
QUESTIONS_FILE_SIMPLE = "questions_v3v2_June03.xlsx"

MAX_WORKERS = 3  # ↓ lower to avoid rate-limit/concurrency errors
MAX_ATTEMPTS = 6
BACKOFF_BASE = 2  # seconds; wait is BACKOFF_BASE ** attempt

# o1 / o3 / o4-mini context handling
CONTEXT_O1 = 128_000        # hard window for o-series
SAFETY_MARGIN = 4_096       # keep room for system/user messages

MAX_TOKENS_SIMPLE = 200_000  # legacy fallback for simple pipeline

# ---------------------------------------------------------------------------
# Model registry
# ---------------------------------------------------------------------------
AVAILABLE_MODELS: Dict[str, Dict[str, str]] = {
    # OpenAI
    "gpt-4o": {"provider": "openai", "name": "gpt-4o"},
    "gpt-4.1": {"provider": "openai", "name": "gpt-4.1"},
    "gpt-3.5-turbo": {"provider": "openai", "name": "gpt-3.5-turbo"},
    "o1": {"provider": "openai", "name": "o1"},
    "o3": {"provider": "openai", "name": "o3"},
    "o4-mini": {"provider": "openai", "name": "o4-mini"},
    # Anthropic
    "claude-3-opus": {"provider": "anthropic", "name": "claude-3-opus"},
    "claude-3-sonnet": {"provider": "anthropic", "name": "claude-3-sonnet"},
    # Google
    "gemini-pro": {"provider": "google", "name": "gemini-pro"},
}

DEFAULT_SIMPLE_MODELS = ["gpt-4.1"]
DEFAULT_ANALYTICAL_MODELS = ["o1"]

openai_client = OpenAI()

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------


def setup_llm(model_config: Dict[str, str]):
    """Return a *fresh* Chat* object; create one per call → thread-safe."""
    provider = model_config["provider"]
    model_name = model_config["name"]

    if provider == "openai":
        if model_name == "o1":
            return ChatOpenAI(model_name=model_name)  # o1 rejects temperature
        if model_name in {"o3", "o4-mini"}:
            return ChatOpenAI(model_name=model_name, temperature=1)
        return ChatOpenAI(model_name=model_name, temperature=0)

    if provider == "anthropic":
        return ChatAnthropic(model_name=model_name, temperature=0)

    if provider == "google":
        return ChatGoogleGenerativeAI(model=model_name, temperature=0)

    raise ValueError(f"Unsupported provider '{provider}' for model '{model_name}'.")


# ──────────────────────────────────────────────────────────────────────────
# Question loading
# ──────────────────────────────────────────────────────────────────────────


def _load_questions(xlsx_path: str) -> tuple[list[str], dict[str, str]]:
    df = pd.read_excel(xlsx_path, header=None)
    first_row = df.iloc[0]

    prefix = (
        "After having gone through the PDF/document above, extract the specific "
        "parameter outlined below based on the contextual description/information "
        "accompanying it, and always replying strictly as outlined in the Output "
        "Format below."
    )

    parameters = first_row.dropna().tolist()
    prefixed = [prefix + param for param in parameters]
    return prefixed, dict(zip(prefixed, parameters))


def load_questions_analytical():  # noqa: D401
    return _load_questions(QUESTIONS_FILE_ANALYTICAL)


def load_questions_simple():  # noqa: D401
    return _load_questions(QUESTIONS_FILE_SIMPLE)


# ──────────────────────────────────────────────────────────────────────────
# Token utils
# ──────────────────────────────────────────────────────────────────────────


def count_tokens(text: str) -> int:
    enc = tiktoken.get_encoding("cl100k_base")
    return len(enc.encode(text))


def trim_to_context(text: str, max_tokens: int) -> str:
    enc = tiktoken.get_encoding("cl100k_base")
    tokens = enc.encode(text)
    if len(tokens) <= max_tokens:
        return text
    return enc.decode(tokens[:max_tokens])


# Strip ``` fences if the model wrapped JSON in a code block
def _strip_code_fences(text: str) -> str:
    match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.S)
    return match.group(1) if match else text


def extract_value_from_json(response_text: str) -> str:
    cleaned = _strip_code_fences(response_text)
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict) and data:
            return data[list(data.keys())[0]]
    except json.JSONDecodeError:
        pass
    return response_text


# ---------------------------------------------------------------------------
# Analytical pipeline
# ---------------------------------------------------------------------------


def _ask_question_analytical(
    model_config: Dict[str, str],
    text_content: str,
    question: str,
    pdf_file: str,
    question_mapping: Dict[str, str],
) -> Dict[str, Any]:
    """Single LLM call with retries & back-off; never returns None."""
    model_name = model_config["name"]
    llm = setup_llm(model_config)

    # Hard-trim once; subsequent attempts will *also* trim by 10 %
    current_txt = trim_to_context(
        text_content, CONTEXT_O1 - SAFETY_MARGIN if model_name.startswith("o") else MAX_TOKENS_SIMPLE
    )

    for attempt in range(MAX_ATTEMPTS):
        try:
            msg = [
                SystemMessage(
                    content="You are an assistant that extracts parameters and information from documents."
                ),
                HumanMessage(content=f"Document:\n{current_txt}\n\nQuestion:\n{question}\n"),
            ]
            response = llm.invoke(msg)
            answer = extract_value_from_json(response.content)
            return {
                "Document": pdf_file,
                "Question": question_mapping.get(question, question),
                "Answer": answer,
                "Model": model_name,
                "Approach": "analytical",
            }
        except Exception as exc:
            # decide if we can retry
            retryable = any(
                key in str(exc).lower()
                for key in (
                    "rate limit",
                    "overloaded",
                    "maximum context length",
                    "context length exceeded",
                    "connection",
                    "timeout",
                )
            )

            print(
                f"[{type(exc).__name__}] {model_name} – {pdf_file} – "
                f"{question_mapping.get(question, question)[:40]}… :: {exc}"
            )

            if not retryable or attempt == MAX_ATTEMPTS - 1:
                # final answer is an explicit error
                return {
                    "Document": pdf_file,
                    "Question": question_mapping.get(question, question),
                    "Answer": f"ERROR: {type(exc).__name__}: {exc}",
                    "Model": model_name,
                    "Approach": "analytical",
                }

            # back-off then shrink context by 10 % just in case
            time.sleep(BACKOFF_BASE ** attempt)
            current_txt = trim_to_context(current_txt, int(count_tokens(current_txt) * 0.9))

    # should never reach
    return {
        "Document": pdf_file,
        "Question": question_mapping.get(question, question),
        "Answer": "ERROR: Unknown failure",
        "Model": model_name,
        "Approach": "analytical",
    }


def process_documents_analytical(
    model_name: str,
    model_config: Dict[str, str],
    questions: List[str],
    question_mapping: Dict[str, str],
) -> List[Dict[str, Any]]:
    results: list[dict[str, Any]] = []

    pdf_files = [f for f in os.listdir(INPUT_FOLDER_ANALYTICAL) if f.endswith(".pdf")]
    for pdf_file in pdf_files:
        pdf_path = os.path.join(INPUT_FOLDER_ANALYTICAL, pdf_file)
        text_content = "\n".join(page.page_content for page in PyPDFLoader(pdf_path).load())

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = [
                pool.submit(_ask_question_analytical, model_config, text_content, q, pdf_file, question_mapping)
                for q in questions
            ]
            for fut in as_completed(futures):
                results.append(fut.result())

    return results


# ---------------------------------------------------------------------------
# Simple pipeline
# ---------------------------------------------------------------------------


def _ask_question_simple(
    model_name: str,
    text_content: str,
    question: str,
    pdf_file: str,
    question_mapping: Dict[str, str],
) -> Dict[str, Any]:
    role_msg = "You are an assistant that extracts parameters and information from documents."
    user_msg = f"Document:\n{text_content}\n\nQuestion:\n{question}\n"

    try:
        response = openai_client.chat.completions.create(
            model=model_name,
            messages=[{"role": "system", "content": role_msg}, {"role": "user", "content": user_msg}],
            temperature=0,
            max_tokens=2048,
        )
        raw = response.choices[0].message.content.strip()
        answer = extract_value_from_json(raw)
        return {
            "Document": pdf_file,
            "Question": question_mapping.get(question, question),
            "Answer": answer,
            "Model": model_name,
            "Approach": "simple",
        }
    except Exception as exc:
        print(
            f"[{type(exc).__name__}] {model_name} – {pdf_file} – "
            f"{question_mapping.get(question, question)[:40]}… :: {exc}"
        )
        return {
            "Document": pdf_file,
            "Question": question_mapping.get(question, question),
            "Answer": f"ERROR: {type(exc).__name__}: {exc}",
            "Model": model_name,
            "Approach": "simple",
        }


def process_documents_simple(
    simple_models: List[str],
    questions: List[str],
    question_mapping: Dict[str, str],
) -> List[Dict[str, Any]]:
    results: list[dict[str, Any]] = []

    pdf_files = [f for f in os.listdir(INPUT_FOLDER_SIMPLE) if f.endswith(".pdf")]
    for pdf_file in pdf_files:
        pdf_path = os.path.join(INPUT_FOLDER_SIMPLE, pdf_file)
        text_content = "\n".join(page.page_content for page in PyPDFLoader(pdf_path).load())

        for model_name in simple_models:
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
                futures = [
                    pool.submit(
                        _ask_question_simple, model_name, text_content, q, pdf_file, question_mapping
                    )
                    for q in questions
                ]
                for fut in as_completed(futures):
                    results.append(fut.result())
    return results


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------


def main(
    analytical_models: List[str] | None = None,
    simple_models: List[str] | None = None,
    run_simple: bool = True,
    run_analytical: bool = True,
):
    analytical_models = analytical_models or DEFAULT_ANALYTICAL_MODELS
    simple_models = simple_models or DEFAULT_SIMPLE_MODELS

    all_results: list[dict[str, Any]] = []

    if run_simple and simple_models:
        print("\n▶ Running *simple* pipeline with:", ", ".join(simple_models))
        questions, mapping = load_questions_simple()
        all_results.extend(process_documents_simple(simple_models, questions, mapping))

    if run_analytical and analytical_models:
        print("\n▶ Running *analytical* pipeline with:", ", ".join(analytical_models))
        questions, mapping = load_questions_analytical()
        for m in analytical_models:
            if m not in AVAILABLE_MODELS:
                print(f"  – Skipping unknown model '{m}'.")
                continue
            cfg = AVAILABLE_MODELS[m]
            all_results.extend(process_documents_analytical(m, cfg, questions, mapping))

    if not all_results:
        print("\nNo results generated – check configuration and input directory.")
        return

    # ------------------------------------------------------------------
    # Persist results to Excel
    # ------------------------------------------------------------------
    df = pd.DataFrame(all_results)

    simple_df = df[df["Approach"] == "simple"]
    analytical_df = df[df["Approach"] == "analytical"]

    simple_wide = (
        simple_df.pivot_table(index=["Document"], columns="Question", values="Answer", aggfunc="first")
        .reset_index()
        .sort_index(axis=1)
    )
    analytical_wide = (
        analytical_df.pivot_table(index=["Document"], columns="Question", values="Answer", aggfunc="first")
        .reset_index()
        .sort_index(axis=1)
    )

    # combined view (simple first, analytical suffixed)
    wide = simple_wide.copy()
    if not analytical_df.empty:
        for question in analytical_df["Question"].unique():
            col_name = f"{question}_analytical"
            analytical_answers = (
                analytical_df[analytical_df["Question"] == question]
                .set_index("Document")["Answer"]
                .astype(str)
            )
            wide[col_name] = wide["Document"].map(analytical_answers)

    out_long = f"output_combined_{TIMESTAMP}.xlsx"
    out_pivot = f"output_combined_pivoted_{TIMESTAMP}.xlsx"
    out_simple = f"output_simple_{TIMESTAMP}.xlsx"
    out_analytical = f"output_analytical_{TIMESTAMP}.xlsx"

    df.to_excel(out_long, index=False)
    wide.to_excel(out_pivot, index=False)
    simple_wide.to_excel(out_simple, index=False)
    analytical_wide.to_excel(out_analytical, index=False)

    print(
        f"\n✓ Analysis complete – results written to:\n"
        f"  • {out_long}\n  • {out_pivot}\n  • {out_simple}\n  • {out_analytical}"
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run document tagging across multiple LLMs.")
    parser.add_argument(
        "--simple",
        nargs="*",
        help="Space-separated list of *simple* models to use (OpenAI chat models).",
    )
    parser.add_argument(
        "--analytical",
        nargs="*",
        help="Space-separated list of *analytical* models to use.",
    )
    parser.add_argument("--skip-simple", action="store_true", help="Skip simple pipeline.")
    parser.add_argument("--skip-analytical", action="store_true", help="Skip analytical pipeline.")
    args = parser.parse_args()

    main(
        analytical_models=args.analytical,
        simple_models=args.simple,
        run_simple=not args.skip_simple,
        run_analytical=not args.skip_analytical,
    )