import re
import os
import logging
import configparser
from collections import defaultdict
from pathlib import Path
import functions as fnc
import datetime
import unicodedata

current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

current_root = Path().resolve()

config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
rootDir = current_root / Path(config['DEFAULT']['logFiledir'])
viewTableDir = current_root / Path(config['DEFAULT']['view_table_list_file_dir'])
tableListDir = current_root / Path(config['DEFAULT']['table_list_file_dir'])

rootTempDir = current_root / Path(config['DEFAULT']['tempDir'])

if not os.path.exists(rootTempDir):
    os.makedirs(rootTempDir)

# Open the log file

column_map = defaultdict(list)
column_map2 = defaultdict(list)
file_name = []

error_log_dir = current_root / Path(config['DEFAULT']['errorlogDir']) 
error_log_dir.mkdir(exist_ok=True)
read_file_error_log_path = error_log_dir / 'read_files_error.log' 
file_process_error_log_path = error_log_dir / 'file_process_error.log' 


def setup_logger(table_folder_path,table_name):
    # Create the directory if it doesn't exist
    if not os.path.exists(table_folder_path):
        os.makedirs(table_folder_path, exist_ok=True)

    logger = logging.getLogger(table_name.lower())  # Logger name based on table name
    logger.setLevel(logging.INFO)  # Set logging level to INFO

    # Create a file handler for the logger in the specified directory
    handler = logging.FileHandler(os.path.join(table_folder_path, f"{table_name}.txt"))
    handler.setLevel(logging.INFO)  # Set handler level to INFO

    # Create a formatter and set it for the handler
    formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    
    return logger

def extract_table_names(section_content, folder_path, view_tables = []):
    section_content = section_content.replace("Insert Table Name:", "Select Table Name:")
    section_content = section_content.replace("Update Table Name:", "Select Table Name:")
    section_content = section_content.replace("Delete Table Name:", "Select Table Name:")
    section_content = section_content.replace("Insert Columns:", "Select Columns:")
    section_content = section_content.replace("Set Columns:", "Select Columns:")
    section_content = section_content.replace("Where Columns:", "Select Columns:")
    section_content = section_content.replace("Order Columns:", "Select Columns:")
    section_content = section_content.replace("Group Columns:", "Select Columns:")
    view_table_name = ''
    target_view_table_name = ''
    # check for exists view or not
    if "View Table Name: " in section_content:
        view_table_name_pattern = re.compile(r'View Table Name:\s*(.*)', re.IGNORECASE)
        view_table_name = view_table_name_pattern.search(section_content)
        target_view_table_name = view_table_name.group(1).lower()
        target_view_table_name = unicodedata.normalize('NFKC', target_view_table_name)
     
    section_content = concat_column(section_content)
    select_table_name_pattern = re.compile(r'Select Table Name:\s*(.*)', re.IGNORECASE)
    select_column_name_pattern = re.compile(r'Select Columns:\s*(.*)', re.IGNORECASE)

    match = select_table_name_pattern.search(section_content)
    match_column = select_column_name_pattern.search(section_content)

    split_table_name = []
    write_table_name = ''   
    if match:
        table_name = match.group(1)
        split_table_name = re.split(r',\s*', table_name.lower())
       
        if target_view_table_name != '':
            folder_path = os.path.join(folder_path, target_view_table_name)
        
        for index,sp_name in enumerate(split_table_name):
            table_folder_path = folder_path
            table_name_split  = re.compile(r'\bAS\s+(\w+)', re.IGNORECASE)
            output_table      = table_name_split.sub('', sp_name)
            write_table_name  = output_table.strip()
            split_strings     = write_table_name.split('.')
            
            if len(split_strings) > 1:  # Combine the conditions
                write_table_name = split_strings[-1]
                db_folder_path = os.path.join(folder_path, *split_strings[:-1])
                table_folder_path = db_folder_path
            
            if write_table_name != '' and write_table_name.lower() != 'where' : 
                # convert full width to helf width
                write_table_name = unicodedata.normalize('NFKC', write_table_name)

                # if table name exists in view
                if write_table_name.lower() in view_tables:
                    table_folder_path = rootTempDir
                    full_remove_str = ' '.join(split_strings[:-1]) + '.'
                    full_temp_name = sp_name.replace(full_remove_str , '')
                    split_table_name[index] = full_temp_name
                     
                logger = setup_logger(table_folder_path,write_table_name.lower() )
    
    if match_column:
        split_name = ''
        column_name = match_column.group(1)

        if ',' in column_name:
            split_name = re.split(r',\s*', column_name)
            for sp_name in split_name:
                sp_name = sp_name.replace("[", "")
                sp_name = sp_name.replace("]", "")

                extract_column_names(sp_name, split_table_name, target_view_table_name)     
        else:
            extract_column_names(column_name, split_table_name, target_view_table_name)
                             
    return column_map

def concat_column(input_string):
    table_name = ""
    columns = []
    pattern = re.compile(r"Select Columns:\s*\[(.*?)\]")
    # Split input_string into lines and process each line
    for line in input_string.strip().splitlines():
        line = line.strip()
        if line.startswith("Select Table Name:"):
            table_name = line.split("Select Table Name:")[1].strip()
        elif line.startswith("Select Columns:"):
            match = pattern.search(line)
            if match:
                column_str = match.group(1)
                if column_str:  # Check if column_str is not empty
                    columns.append(column_str)

    columns = list(dict.fromkeys(columns))

    final_result = f"Select Table Name: {table_name}\n"
    if columns:
        final_result += "Select Columns: [" + "','".join(columns) + "]"

    return final_result

def extract_column_names(sp_name, split_table_name, target_view_table_name):

    column_name_split = re.compile(r'(.+?)\s*\.(\w+)', re.IGNORECASE)

    col_matches = column_name_split.findall(sp_name)

    # Regular expression pattern to match 'table_name AS alias_name'
    table_pattern = re.compile(r'(.+?)\s+AS\s+(\w+)\b', re.IGNORECASE)      
  
    if col_matches:
        parts = sp_name.split('.')
        col_name = parts[-1].strip("'").split(' as ')[0].strip()  # Last part is the column name
         
        # Joining the other parts with dot to form the alias
        col_alias = '.'.join(parts[:-1]).strip("'").replace("['",'')

        for table_name in split_table_name:
            table_name = unicodedata.normalize('NFKC', table_name)
            target_table_name = table_name
            if target_view_table_name != '':
                target_table_name = target_view_table_name + '.' + table_name
  
            output_table = table_pattern.findall(target_table_name)
           
            if output_table:
                 # for table_name As table_alais   
                for output_table_name, table_alias in output_table:
                    if col_alias == table_alias and col_name not in column_map[output_table_name]:
                        column_map[output_table_name].append(col_name)
            else:
                #  for table_name.col
                if col_alias.lower() == target_table_name.lower() and col_name not in column_map[target_table_name]:
                           
                    column_map[target_table_name].append(col_name)
            
    else:    
        for table_name in split_table_name:
            table_name = unicodedata.normalize('NFKC', table_name)
            target_table_name = table_name
            if target_view_table_name != '':
                target_table_name = target_view_table_name + '.' + table_name
   
            if ' as' in table_name.lower():
                output_table = table_pattern.findall(target_table_name)
                if output_table:
                    table_name = output_table[0][0]
                    if target_view_table_name != '':
                        target_table_name = target_view_table_name + '.' + table_name
               
            if sp_name not in column_map[target_table_name]:
                column_map[target_table_name].append(sp_name)  

            
def column_append(rootDir, value, target_log ,view_tables):
    sub_path = ''
    files = [file for file in rootDir.rglob('*') if file.is_file()]
    for file_path in files:
        parts = str(file_path).split('table_list', 1)
        if len(parts) > 1:
            sub_path = parts[1].lstrip('\\').lower()
        else:
            sub_path = file_path.name   
         
        if sub_path == target_log:
            exiting_data2 = []

            with file_path.open('a', encoding='utf-8') as file:
                for col in value:
                    col = unicodedata.normalize('NFKC', col)
                    exiting_data = fnc.read_file_to_array(file_path)
                    fixed_col = format_replace(col)
                    
                    if view_tables != []:
                        fixed_col = fixed_col.replace("*", "")
                        
                    if fixed_col not in exiting_data and fixed_col not in exiting_data2 and fixed_col != '':
                        exiting_data2.append(fixed_col)
                        file.write(fixed_col + '\n')

def format_replace(col):
    col = col.replace("'", "")
    col = col.replace("[", "")
    col = col.replace("]", "")
    col = col.replace("{", "")
    col = col.replace("}", "")
    col = col.replace("CAST", "")
    col = col.replace(")", "")
    col = col.replace("(", "")
    col = col.replace("&quot;%", "")
    col = col.replace("yyyy/mm/dd", "")
    col = col.replace("\n", "")
    col = col.replace(";", "")
    col = col.replace(",", "")

    return col.strip()

def analysis_for_view_files():
    files_dir1 = os.listdir(rootTempDir)

    lines = []
    view_table_lines = []
    sub_path = ''

    # Compare files
    for file in files_dir1:
        file1_path = os.path.join(rootTempDir, file)
        check_name = file.replace('.txt','').lower()
        
        # Check if the file exists in the second directory
        for file2 in files_dir1:
            check_file_name2 = file2.replace('.txt','').lower()
            if check_name == check_file_name2:
                lines = fnc.read_file_to_array(file1_path)
                rootDir = Path(os.path.join(viewTableDir, check_name))
                files = [file for file in rootDir.rglob('*') if file.is_file()]
                for view_file in files:
                    file2_path = os.path.join(viewTableDir, view_file)
                    view_table_lines = fnc.read_file_to_array(file2_path)

                    parts = str(file2_path).split(check_name, 1)
                    if len(parts) > 1:
                        sub_path = parts[1].lstrip('\\').lower()
                    else:
                        sub_path = file2_path.name
                    
                    for view_line in view_table_lines:
                        if view_line == '*':
                            for line in lines:
                               file_name = sub_path.replace('.txt','')
                               column_map2[file_name].append(line)

                    for line in lines:
                        if line in view_table_lines:
                            file_name = sub_path.replace('.txt','')
                            column_map2[file_name].append(line)

    create_file_and_append_coumns_for_views()                  

def create_file_and_append_coumns_for_views():
    for key,value in column_map2.items():
        split_names = key.rsplit('\\', 1)
        if len(split_names) != 1:
            path = split_names[0]
            folder_path = os.path.join(tableListDir, path)
            table_name = split_names[1]
        else:
            folder_path = tableListDir
            table_name = key

        logger = setup_logger(folder_path,table_name)
        target_log = f"{key}.txt".lower()

        column_append(tableListDir, value, target_log ,view_tables = [])

def force_delete_directory():
    """Force delete a directory using the system command."""
    try:
        os.system(f'rd /S /Q "{rootTempDir}"')
        print(f"Directory {rootTempDir} deleted successfully.")
    except Exception as e:
        with file_process_error_log_path.open('a', encoding='utf-8') as error_log:
            error_log.write(f'{current_time} - Error in log_file_analysis.py/force_delete_directory function: {e}\n')

def start_run(file_path, rootDir2, folder_path, view_tables = []):
    try:
        file_path = Path(file_path)

        with file_path.open('r', encoding='utf-8') as file:
            file_name.append(file_path.name)
            # Read each line in the file
            file_content = file.read()
            # Split the content by the delimiter
            sections = file_content.split('===================================')
            # Process each section
            column_map = {}
            for section in sections:
                section_content = section.strip()
                if section_content:  # Ensure the section is not empty
                    column_map = extract_table_names(section_content, folder_path, view_tables)

            for key,value in column_map.items(): 
                table_name = key.lower().replace('.', '\\')  
                target_log = f"{table_name}.txt".lower()
                view_table_name = key.split('.')
                
                if view_table_name[-1].lower() in view_tables:
                    column_append(rootTempDir, value, target_log, view_tables )
                elif key.lower() in view_tables:
                    column_append(rootTempDir, value, target_log, view_tables)
                else:
                    column_append(rootDir2, value, target_log, view_tables)
        
        if len(view_tables) != 0:
            analysis_for_view_files()
            force_delete_directory()

    except UnicodeDecodeError:
        with read_file_error_log_path.open('a', encoding='utf-8') as error_log:
            error_log.write(f'{current_time} - Error in log_file_analysis.py/start_run function: Unable to decode {file_path}\n')

    except Exception as e:
        with file_process_error_log_path.open('a', encoding='utf-8') as error_log:
            error_log.write(f'{current_time} - Error in log_file_analysis.py/start_run function: {e}\n')

                             
            