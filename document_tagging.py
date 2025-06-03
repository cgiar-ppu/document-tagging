import os
import json
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any
from concurrent.futures import ThreadPoolExecutor, as_completed
import re
import tiktoken

from langchain_community.document_loaders import PyPDFLoader
from langchain_openai import ChatOpenAI
from langchain_anthropic import ChatAnthropic
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.schema import HumanMessage, SystemMessage
from openai import OpenAI

# -----------------------------------------------------------------------------
# Configuration & constants
# -----------------------------------------------------------------------------
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")

INPUT_FOLDER_ANALYTICAL = "input_taggedstudies"  # reasoning‑LLM pipeline
INPUT_FOLDER_SIMPLE = "input_taggedstudies"      # normal‑LLM pipeline

QUESTIONS_FILE_ANALYTICAL = "questions_v3v2_hugo.xlsx"
QUESTIONS_FILE_SIMPLE = "questions_v3v2_June03.xlsx"

MAX_WORKERS = 10
MAX_TOKENS = 200_000
REDUCTION_FACTOR = 0.9

# -----------------------------------------------------------------------------
# Model registry
#   – Add every model you might want to call ONCE here, tagged by provider.
# -----------------------------------------------------------------------------
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

# Normal‑LLMs vs reasoning‑LLMs ────────────────────────────────────────────────
DEFAULT_SIMPLE_MODELS = ["gpt-4.1"]          # "normal" LLMs
DEFAULT_ANALYTICAL_MODELS = ["o1"]    # "reasoning" LLMs

# Single OpenAI client for the lightweight/simple pipeline
openai_client = OpenAI()

# -----------------------------------------------------------------------------
# Helper functions
# -----------------------------------------------------------------------------

def setup_llm(model_config: Dict[str, str]):
    """Instantiate an LLM wrapper (LangChain) respecting every provider's quirks.

    **OpenAI specifics**
    ───────────────────
    • **o1**         → rejects *any* `temperature` parameter → omit entirely.
    • **o3, o4‑mini**→ accept the field but only at their default (1).
    • Everything else (gpt‑4o, gpt‑4.1, gpt‑3.5‑turbo, …) → explicit `temperature=0` for maximum determinism.
    """

    provider = model_config["provider"]
    model_name = model_config["name"]

    if provider == "openai":
        # ---- Strict rules per model -----------------------------------------
        if model_name == "o1":
            # No temperature field at all.
            return ChatOpenAI(model_name=model_name)
        if model_name in {"o3", "o4-mini"}:
            # Must keep default value (server interprets as 1).
            return ChatOpenAI(model_name=model_name, temperature=1)
        # All other chat models – allow deterministic 0 temperature.
        return ChatOpenAI(model_name=model_name, temperature=0)

    if provider == "anthropic":
        return ChatAnthropic(model_name=model_name, temperature=0)

    if provider == "google":
        return ChatGoogleGenerativeAI(model=model_name, temperature=0)

    raise ValueError(f"Unsupported provider '{provider}' for model '{model_name}'.")

    if provider == "openai":
        # o‑series prefer temperature≈0 by default like other chat models.
        return ChatOpenAI(model_name=model_name, temperature=0)
    if provider == "anthropic":
        return ChatAnthropic(model_name=model_name, temperature=0)
    if provider == "google":
        return ChatGoogleGenerativeAI(model=model_name, temperature=0)

    raise ValueError(f"Unsupported provider '{provider}' for model '{model_name}'.")


# ──────────────────────────────────────────────────────────────────────────────
# Question loading utilities
# ──────────────────────────────────────────────────────────────────────────────

def _load_questions(xlsx_path: str) -> tuple[list[str], dict[str, str]]:
    """Read the first row of the sheet and prepend the autonomous‑prompt."""
    questions_df = pd.read_excel(xlsx_path, header=None)
    row_1 = questions_df.iloc[0]

    prefix = (
        "After having gone through the PDF/document above, extract the specific "
        "parameter outlined below based on the contextual description/information "
        "accompanying it, and always replying strictly as outlined in the Output "
        "Format below."
    )

    parameters = row_1.dropna().tolist()
    prefixed = [prefix + param for param in parameters]
    return prefixed, dict(zip(prefixed, parameters))


def load_questions_analytical():
    return _load_questions(QUESTIONS_FILE_ANALYTICAL)


def load_questions_simple():
    return _load_questions(QUESTIONS_FILE_SIMPLE)


# ──────────────────────────────────────────────────────────────────────────────
# House‑keeping utils
# ──────────────────────────────────────────────────────────────────────────────

def _strip_code_fences(text: str) -> str:
    pattern = r"```(?:json)?\s*(.*?)\s*```"
    match = re.search(pattern, text, re.DOTALL)
    return match.group(1) if match else text


def extract_value_from_json(response_text: str) -> str:
    cleaned = _strip_code_fences(response_text)
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict) and data:
            return data[list(data.keys())[0]]
    except json.JSONDecodeError:
        pass
    return response_text  # fall‑back: return raw


def count_tokens(text: str) -> int:
    encoding = tiktoken.get_encoding("cl100k_base")
    return len(encoding.encode(text))


def trim_text_to_fit(text: str, max_tokens: int = MAX_TOKENS) -> str:
    if count_tokens(text) <= max_tokens:
        return text
    encoding = tiktoken.get_encoding("cl100k_base")
    tokens = encoding.encode(text)[:max_tokens]
    return encoding.decode(tokens)


# -----------------------------------------------------------------------------
# 1️⃣ Analytical pipeline     (reasoning‑heavy models / LangChain wrappers)
# -----------------------------------------------------------------------------

def _ask_question_analytical(
    llm, text_content: str, question: str, pdf_file: str, model_name: str
) -> Dict[str, Any] | None:
    current_text = text_content
    for attempt in range(3):
        try:
            msg = [
                SystemMessage(content="You are an assistant that extracts parameters and information from documents."),
                HumanMessage(content=f"Document:\n{current_text}\n\nQuestion:\n{question}\n")
            ]
            response = llm.invoke(msg)
            answer = extract_value_from_json(response.content)
            return {
                "Document": pdf_file,
                "Question": question,
                "Answer": answer,
                "Model": model_name,
                "Approach": "analytical",
            }
        except Exception as exc:
            if attempt == 2:
                print(f"[ERROR] {model_name} – {pdf_file} – '{question[:30]}…': {exc}")
            current_text = trim_text_to_fit(current_text, int(count_tokens(current_text) * REDUCTION_FACTOR))
    return None


def process_documents_analytical(
    model_name: str,
    model_config: Dict[str, str],
    questions: List[str],
    question_mapping: Dict[str, str],
) -> List[Dict[str, Any]]:
    llm = setup_llm(model_config)
    results: list[dict[str, Any]] = []

    # Iterate documents
    for pdf_file in filter(lambda f: f.endswith(".pdf"), os.listdir(INPUT_FOLDER_ANALYTICAL)):
        pdf_path = os.path.join(INPUT_FOLDER_ANALYTICAL, pdf_file)
        text_content = "\n".join(page.page_content for page in PyPDFLoader(pdf_path).load())

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = {
                pool.submit(_ask_question_analytical, llm, text_content, q, pdf_file, model_name): q
                for q in questions
            }
            for fut in as_completed(futures):
                if (res := fut.result()):
                    res["Question"] = question_mapping[res["Question"]]
                    results.append(res)
    return results


# -----------------------------------------------------------------------------
# 2️⃣ Simple pipeline         (lightweight models via direct OpenAI client)
# -----------------------------------------------------------------------------

def _ask_question_simple(
    model_name: str,
    text_content: str,
    question: str,
    pdf_file: str,
) -> Dict[str, Any] | None:
    role_msg = "You are an assistant that extracts parameters and information from documents."
    user_msg = f"Document:\n{text_content}\n\nQuestion:\n{question}\n"

    try:
        response = openai_client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": role_msg},
                {"role": "user", "content": user_msg},
            ],
            temperature=0,
            max_tokens=2048,
        )
        raw = response.choices[0].message.content.strip()
        answer = extract_value_from_json(raw)
        return {
            "Document": pdf_file,
            "Question": question,
            "Answer": answer,
            "Model": model_name,
            "Approach": "simple",
        }
    except Exception as exc:
        print(f"[ERROR] {model_name} – {pdf_file} – '{question[:30]}…': {exc}")
    return None


def process_documents_simple(
    simple_models: List[str],
    questions: List[str],
    question_mapping: Dict[str, str],
) -> List[Dict[str, Any]]:
    results: list[dict[str, Any]] = []

    for pdf_file in filter(lambda f: f.endswith(".pdf"), os.listdir(INPUT_FOLDER_SIMPLE)):
        pdf_path = os.path.join(INPUT_FOLDER_SIMPLE, pdf_file)
        text_content = "\n".join(page.page_content for page in PyPDFLoader(pdf_path).load())

        for model_name in simple_models:
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
                futures = {
                    pool.submit(_ask_question_simple, model_name, text_content, q, pdf_file): q
                    for q in questions
                }
                for fut in as_completed(futures):
                    if (res := fut.result()):
                        res["Question"] = question_mapping[res["Question"]]
                        results.append(res)
    return results


# -----------------------------------------------------------------------------
# Main orchestrator
# -----------------------------------------------------------------------------

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
    # Persist results to Excel in both long and pivoted shape
    # ------------------------------------------------------------------
    df = pd.DataFrame(all_results)
    wide = df.pivot_table(
        index=["Document", "Model", "Approach"],
        columns="Question",
        values="Answer",
        aggfunc="first",
    ).reset_index()

    out_long = f"output_combined_{TIMESTAMP}.xlsx"
    out_pivot = f"output_combined_pivoted_{TIMESTAMP}.xlsx"

    df.to_excel(out_long, index=False)
    wide.to_excel(out_pivot, index=False)

    print(
        f"\n✓ Analysis complete – results written to:\n  • {out_long}\n  • {out_pivot}"
    )


# -----------------------------------------------------------------------------
# Sample CLI usage
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    """Run with the default model selections.

    Examples:
        python document_tagging_updated.py                         # defaults
        python document_tagging_updated.py --simple gpt-4o gpt-4.1 \
                                          --analytical o1 o3      # explicit
    """

    # A tiny CLI (optional) – keeps the script self‑contained.
    import argparse

    parser = argparse.ArgumentParser(description="Run document tagging across multiple LLMs.")
    parser.add_argument(
        "--simple",
        nargs="*",
        help="Space‑separated list of *simple* models to use (OpenAI chat models).",
    )
    parser.add_argument(
        "--analytical",
        nargs="*",
        help="Space‑separated list of *analytical* models to use.",
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
