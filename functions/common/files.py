import os
from pathlib import Path
import shutil



def read_file_to_array(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found at: {file_path}")
    
    lines = []
    with open(file_path, 'r', encoding='utf-8') as file:  
        for line in file:
            stripped_line = line.strip()
            lines.append(stripped_line)
       
    return lines

def delete_folder(folder_path):
    try:
        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)
            print(f"Successfully deleted folder: {folder_path}")
        else:
            print(f"Folder does not exist: {folder_path}")
    except Exception as e:
        print(f"Error deleting folder {folder_path}: {e}")