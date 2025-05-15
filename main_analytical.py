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

# Configuration
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
INPUT_FOLDER = 'input_taggedstudies'
QUESTIONS_FILE = 'questions_v3v2_hugo.xlsx'
MAX_WORKERS = 10
MAX_TOKENS = 200000  # Maximum token limit
REDUCTION_FACTOR = 0.9  # Reduce text size by 10% on failure

# Model configurations
AVAILABLE_MODELS = {
    'gpt-4o': {'provider': 'openai', 'name': 'gpt-4o'},
    'o1': {'provider': 'openai', 'name': 'o1'},
    'o3': {'provider': 'openai', 'name': 'o3'},
    'gpt-4.1': {'provider': 'openai', 'name': 'gpt-4.1'},
    'o4-mini': {'provider': 'openai', 'name': 'o4-mini'},
    'gpt-3.5-turbo': {'provider': 'openai', 'name': 'gpt-3.5-turbo'},
    'claude-3-opus': {'provider': 'anthropic', 'name': 'claude-3-opus'},
    'claude-3-sonnet': {'provider': 'anthropic', 'name': 'claude-3-sonnet'},
    'gemini-pro': {'provider': 'google', 'name': 'gemini-pro'},
}

def setup_llm(model_config: Dict[str, str]) -> Any:
    """Initialize the appropriate LLM based on the provider."""
    provider = model_config['provider']
    model_name = model_config['name']
    
    if provider == 'openai':
        if model_name == 'o1':
            return ChatOpenAI(model_name=model_name)  # No temperature for o1 model
        return ChatOpenAI(model_name=model_name, temperature=0)
    elif provider == 'anthropic':
        return ChatAnthropic(model_name=model_name, temperature=0)
    elif provider == 'google':
        return ChatGoogleGenerativeAI(model=model_name, temperature=0)
    else:
        raise ValueError(f"Unsupported provider: {provider}")

def load_questions() -> tuple:
    """Load and prepare questions from Excel file."""
    questions_df = pd.read_excel(QUESTIONS_FILE, header=None)
    row_2 = questions_df.iloc[1]
    
    prefix = "After having gone through the PDF/document above, extract the specific parameter outlined below based on the contextual description/information accompanying it, and always replying strictly as outlined in the Output Format below."
    parameters = row_2.iloc[0:16].dropna().tolist()
    original_questions = parameters.copy()
    prefixed_parameters = [prefix + param for param in parameters]
    
    return prefixed_parameters, dict(zip(prefixed_parameters, original_questions))

def extract_value_from_json(response_text: str) -> str:
    """Extract value from JSON response."""
    def remove_markdown_code_fences(text: str) -> str:
        pattern = r"```(?:json)?\s*(.*?)\s*```"
        match = re.search(pattern, text, re.DOTALL)
        return match.group(1) if match else text

    cleaned = remove_markdown_code_fences(response_text)
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict) and len(data) > 0:
            first_key = list(data.keys())[0]
            return data[first_key]
        return response_text
    except json.JSONDecodeError:
        return response_text

def count_tokens(text: str) -> int:
    """Count the number of tokens in a text using tiktoken."""
    encoding = tiktoken.get_encoding("cl100k_base")
    return len(encoding.encode(text))

def trim_text_to_fit(text: str, max_tokens: int = MAX_TOKENS) -> str:
    """Trim text to fit within token limit while preserving as much content as possible."""
    if count_tokens(text) <= max_tokens:
        return text
        
    encoding = tiktoken.get_encoding("cl100k_base")
    tokens = encoding.encode(text)
    return encoding.decode(tokens[:max_tokens])

def process_single_question(
    llm: Any,
    text_content: str,
    question: str,
    pdf_file: str,
    model_name: str
) -> Dict[str, Any]:
    """Process a single question for a document using specified LLM."""
    current_text = text_content
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            messages = [
                SystemMessage(content="You are an assistant that extracts parameters and information from documents."),
                HumanMessage(content=f"Document:\n{current_text}\n\nQuestion:\n{question}\n\n")
            ]
            
            response = llm.invoke(messages)
            answer = extract_value_from_json(response.content)
            
            return {
                'Document': pdf_file,
                'Question': question,
                'Answer': answer,
                'Model': model_name
            }
        except Exception as e:
            print(f"Attempt {attempt + 1} failed for {model_name} on {pdf_file}: {str(e)}")
            if attempt < max_retries - 1:
                # Reduce text size for next attempt
                current_text = trim_text_to_fit(current_text, int(count_tokens(current_text) * REDUCTION_FACTOR))
            else:
                print(f"Failed to process after {max_retries} attempts")
                return None

def process_documents_with_model(
    model_name: str,
    model_config: Dict[str, str],
    questions: List[str],
    question_mapping: Dict[str, str]
) -> List[Dict[str, Any]]:
    """Process all documents with a specific model."""
    results = []
    llm = setup_llm(model_config)
    
    for pdf_file in os.listdir(INPUT_FOLDER):
        if not pdf_file.endswith('.pdf'):
            continue
            
        pdf_path = os.path.join(INPUT_FOLDER, pdf_file)
        loader = PyPDFLoader(pdf_path)
        document = loader.load()
        text_content = "\n".join([page.page_content for page in document])
        
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(
                    process_single_question,
                    llm,
                    text_content,
                    question,
                    pdf_file,
                    model_name
                ): question 
                for question in questions
            }
            
            for future in as_completed(futures):
                result = future.result()
                if result:
                    result['Question'] = question_mapping[result['Question']]
                    results.append(result)
    
    return results

def main(selected_models: List[str] = None):
    """Main function to run the analysis with selected models."""
    if not selected_models:
        selected_models = list(AVAILABLE_MODELS.keys())
    
    # Validate selected models
    for model in selected_models:
        if model not in AVAILABLE_MODELS:
            raise ValueError(f"Invalid model selection: {model}")
    
    # Load questions
    questions, question_mapping = load_questions()
    all_results = []
    
    # Process documents with each selected model
    for model_name in selected_models:
        print(f"Processing documents with {model_name}...")
        model_config = AVAILABLE_MODELS[model_name]
        results = process_documents_with_model(
            model_name,
            model_config,
            questions,
            question_mapping
        )
        all_results.extend(results)
    
    # Save results
    df = pd.DataFrame(all_results)
    df.to_excel(f'output_comparison_{TIMESTAMP}.xlsx', index=False)
    
    # Create pivoted view
    pivot_df = df.pivot_table(
        index=['Document', 'Model'],
        columns='Question',
        values='Answer',
        aggfunc='first'
    ).reset_index()
    pivot_df.to_excel(f'output_comparison_pivoted_{TIMESTAMP}.xlsx')
    
    print("Analysis completed. Results saved to Excel files.")

if __name__ == "__main__":
    # Example usage: select specific models to compare
    selected_models = [
        #'gpt-4o',
        'o1'
    ]
    main(selected_models) 