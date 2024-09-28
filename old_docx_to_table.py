import os
import docx
import csv
from openai import AzureOpenAI

# Azure OpenAI configuration
gpt4o_endpoint = "https://fuchsopenai.openai.azure.com/openai/deployments/gpt-4o-2/chat/completions?api-version=2023-03-15-preview"
gpt4o_key = "YOUR_API_KEY_HERE"

# Azure OpenAI model and client configuration
model_name = "gpt-4o-2024-08-06"
client = AzureOpenAI(
    azure_endpoint=gpt4o_endpoint,
    api_key=gpt4o_key,
    api_version="2024-02-01"
)


def openai_prompt_response(prompt):
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ]
        )
        response_text = response.choices[0].message.content
        return response_text
    except Exception as e:
        print(f"Error generating text: {e}")
        return None


def table_from_docx(docx_path):
    doc = docx.Document(docx_path)
    matched_tables = []

    for table in doc.tables:
        # Extract headers
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        row_data = []

        # Extract data rows
        for row in table.rows[1:]:
            row_data.append([cell.text.strip() for cell in row.cells])
        
        # Add extracted headers and data to the matched tables
        matched_tables.append((headers, row_data))

    return matched_tables


def to_csv(table_data, csv_file_path):
    if not table_data:
        print("No table data found to save.")
        return

    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for headers, rows in table_data:
            writer.writerow(headers)  # Write headers
            writer.writerows(rows)      # Write data

    print(f"Table saved to {csv_file_path}")


# Loop through all files in the specified directory
reports_directory = "HackathonDataAnalysis-master/HackathonDataAnalysis-master/reports/docx_files"
output_directory = "HackathonDataAnalysis-master/HackathonDataAnalysis-master/old_extracted_tables_from_docx"

if not os.path.exists(output_directory):
    os.makedirs(output_directory)

for filename in os.listdir(reports_directory):
    if filename.endswith('.docx'):
        docx_file_path = os.path.join(reports_directory, filename)
        
        # Extract the table data
        extracted_table_data = table_from_docx(docx_file_path)

        # Create a corresponding CSV file path in the output folder
        csv_file_path = os.path.join(output_directory, f"{os.path.splitext(filename)[0]}.csv")

        # Save the extracted table to a CSV file
        to_csv(extracted_table_data, csv_file_path)
