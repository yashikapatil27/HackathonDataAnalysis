import os
import shutil

def create_folder(directory):
    """
    Function to create a folder if it doesn't exist.
    
    Parameters:
    - directory (str): Path of the folder to be created.
    """
    if not os.path.exists(directory):
        os.makedirs(directory)

def move_files_by_extension(source_dir, extensions, dest_dir):
    """
    Function to move files with a given extension to a new folder.
    
    Parameters:
    - source_dir (str): The directory where the files are currently located.
    - extensions (list): List of file extensions to look for (e.g., ['.docx', '.pdf']).
    - dest_dir (str): The destination directory where files will be moved.
    """
    create_folder(dest_dir)
    
    for file_name in os.listdir(source_dir):
        file_path = os.path.join(source_dir, file_name)
        ext = os.path.splitext(file_name)[1].lower()  # Extract the file extension
        
        if ext in extensions:
            dest_path = os.path.join(dest_dir, file_name)
            shutil.move(file_path, dest_path)
            print(f"Moved {file_name} to {dest_dir}")

def separate_files_by_type(source_dir):
    """
    Function to separate .docx, .pdf, and .xlsx files into separate folders.

    Parameters:
    - source_dir (str): The directory where the files are currently located.
    """
    # Define the destination directories
    docx_dir = os.path.join(source_dir, 'docx_files')
    pdf_dir = os.path.join(source_dir, 'pdf_files')
    xlsx_dir = os.path.join(source_dir, 'xlsx_files')

    # Move .docx files to 'docx_files' folder
    move_files_by_extension(source_dir, ['.docx'], docx_dir)

    # Move .pdf files to 'pdf_files' folder
    move_files_by_extension(source_dir, ['.pdf'], pdf_dir)

    # Move .xlsx files to 'xlsx_files' folder
    move_files_by_extension(source_dir, ['.xlsx'], xlsx_dir)

# Example Usage
source_directory = "./synthetic_files/reports"  # Replace with the path to your folder containing the files
separate_files_by_type(source_directory)
