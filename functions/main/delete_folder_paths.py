import shutil
import os

def delete_folder(folder_path):
    try:
        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)
            print(f"Successfully deleted folder: {folder_path}")
        else:
            print(f"Folder does not exist: {folder_path}")
    except Exception as e:
        print(f"Error deleting folder {folder_path}: {e}")

# Example usage
folders_to_delete = [
    'table_list',
    'temp',
    'view_table_list',
    'logs',
    'excel'
]

current_dir = os.getcwd()
for folder in folders_to_delete:
    full_path = os.path.join(current_dir, folder)
    delete_folder(full_path)