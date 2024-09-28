import os
import docx
import csv
from openai import AzureOpenAI

# Azure OpenAI configuration
gpt4o_endpoint = "https://fuchsopenai.openai.azure.com/openai/deployments/gpt-4o-2/chat/completions?api-version=2023-03-15-preview"
gpt4o_key = "377635a0dfdd4f01a1352ea785ea4537"
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

def extract_machine_and_product_ids_from_table(table):
    """
    Extracts machine IDs and product IDs from a docx table using OpenAI's model.
    """
    # Convert table content to string format for the model to analyze
    table_content = "\n".join(
        ["\t".join([cell.text.strip() for cell in row.cells]) for row in table.rows]
    )

    prompt = (
        "You are a data extraction assistant. Please find all relevant tables with instrument data. "
        "The structure may vary, but typically includes columns related to instruments, primary substances, and conditions. "
        "Extract the names of machines (e.g., Centrifuge X100, Spectrometer Alpha-300, Thermocycler TC-5000, Microplate Reader MRX, "
        "Gas Chromatograph GC-2010, Liquid Chromatograph LC-400, Mass Spectrometer MS-20, pH Meter PH-700, "
        "Conductivity Meter CM-215, UV-Vis Spectrophotometer UV-2600, HPLC System HPLC-9000, FTIR Spectrometer FTIR-8400, "
        "NMR Spectrometer NMR-500, X-Ray Diffractometer XRD-6000, Rheometer R-4500, Titrator T-905, "
        "PCR Machine PCR-96, Ion Chromatograph IC-2100, Four Ball FB-1000, Viscometer VS-300) "
        "and the primary substances (e.g., Jojoba Oil, Beeswax, Gum, Cetyl Alcohol, Vitamin E, Glycerin, Coconut Oil, Almond Oil) "
        "from the following table content:\n"
        f"{table_content}\n"
        "Return the identified machine names under the header 'machine_id' and the primary substances under the header 'product_id'."
    )

    response_text = openai_prompt_response(prompt)
    
    if response_text:
        # Parse response to separate machine_ids and product_ids
        lines = response_text.splitlines()
        machine_ids = []
        product_ids = []
        
        for line in lines:
            if line.startswith("machine_id"):
                continue
            elif line.startswith("product_id"):
                continue
            else:
                parts = line.split('\t')
                if len(parts) > 0:
                    machine_ids.append(parts[0].strip())  # Assume first part is machine
                if len(parts) > 1:
                    product_ids.append(parts[1].strip())  # Assume second part is product
                
        return machine_ids, product_ids
    return [], []

def table_from_docx(docx_path):
    doc = docx.Document(docx_path)
    extracted_machine_ids = []
    extracted_product_ids = []

    for table in doc.tables:
        machine_ids, product_ids = extract_machine_and_product_ids_from_table(table)
        extracted_machine_ids.extend(machine_ids)
        extracted_product_ids.extend(product_ids)

    return extracted_machine_ids, extracted_product_ids

def to_csv(machine_ids, product_ids, csv_file_path):
    if not machine_ids and not product_ids:
        print(f"No data found to save for {csv_file_path}.")
        return

    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['machine_id', 'product_id'])  # Write the headers
        max_length = max(len(machine_ids), len(product_ids))
        for i in range(max_length):
            machine_id = machine_ids[i] if i < len(machine_ids) else ""
            product_id = product_ids[i] if i < len(product_ids) else ""
            writer.writerow([machine_id, product_id])

    print(f"Data saved to {csv_file_path}")

# Directory paths
reports_directory = "HackathonDataAnalysis-master/HackathonDataAnalysis-master/reports/docx_files"
output_directory = "HackathonDataAnalysis-master/HackathonDataAnalysis-master/new_extracted_tables_from_docx"

# Create output directory if it doesn't exist
os.makedirs(output_directory, exist_ok=True)

# Loop through all files in the specified directory
for filename in os.listdir(reports_directory):
    if filename.endswith('.docx'):
        docx_file_path = os.path.join(reports_directory, filename)
        
        # Extract machine and product IDs from the document
        extracted_machine_ids, extracted_product_ids = table_from_docx(docx_file_path)

        # Create a corresponding CSV file path in the output folder
        csv_file_path = os.path.join(output_directory, f"{os.path.splitext(filename)[0]}.csv")

        # Save the extracted machine and product IDs to a CSV file
        to_csv(extracted_machine_ids, extracted_product_ids, csv_file_path)
