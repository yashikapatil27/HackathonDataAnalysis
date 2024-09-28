import os
import docx
import openpyxl
import PyPDF2
from openai import AzureOpenAI

# Azure OpenAI model and client configuration
gpt4o_endpoint = "https://fuchsopenai.openai.azure.com/openai/deployments/gpt-4o-2/chat/completions?api-version=2023-03-15-preview"
gpt4o_key = "377635a0dfdd4f01a1352ea785ea4537"
model_name = "gpt-4o-2024-08-06"

client = AzureOpenAI(
    azure_endpoint=gpt4o_endpoint,
    api_key=gpt4o_key,
    api_version="2024-02-01"
)

def openai_prompt_response(prompt):
    """
    Function to generate a response from the Azure OpenAI API.

    Parameters:
    - prompt (str): The input text prompt for the model.

    Returns:
    - response_text (str): The generated text from the Azure OpenAI API.
    """
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error generating text: {e}")
        return None

def read_docx(file_path):
    """
    Function to read text from a .docx file.
    
    Parameters:
    - file_path (str): Path to the .docx file.

    Returns:
    - text (str): Extracted text from the .docx file.
    """
    doc = docx.Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

def read_xlsx(file_path):
    """
    Function to read text from a .xlsx file.
    
    Parameters:
    - file_path (str): Path to the .xlsx file.

    Returns:
    - text (str): Extracted data from the .xlsx file (as plain text).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    rows = list(sheet.values)
    return "\n".join([", ".join([str(cell) for cell in row]) for row in rows])

def read_pdf(file_path):
    """
    Function to read text from a .pdf file.
    
    Parameters:
    - file_path (str): Path to the .pdf file.

    Returns:
    - text (str): Extracted text from the .pdf file.
    """
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text()
    return text

def process_file(file_path):
    """
    Function to process a file based on its extension.
    
    Parameters:
    - file_path (str): Path to the file.

    Returns:
    - text (str): Extracted text from the file.
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.docx':
        return read_docx(file_path)
    elif ext == '.xlsx':
        return read_xlsx(file_path)
    elif ext == '.pdf':
        return read_pdf(file_path)
    else:
        print(f"Unsupported file format: {ext}")
        return None

# Example: Iterate through files and process each
def process_reports(report_dir):
    for file_name in os.listdir(report_dir):
        file_path = os.path.join(report_dir, file_name)
        print(f"Processing: {file_name}")
        extracted_text = process_file(file_path)
        if extracted_text:
            # Send the extracted text to OpenAI for further analysis
            response = openai_prompt_response(extracted_text)
            print(f"Response for {file_name}:\n{response}\n")

# Directory containing the reports
report_directory = "/Report_1.docx"

# Process all reports in the directory
process_reports(report_directory)
