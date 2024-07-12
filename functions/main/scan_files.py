import os
import re
import time
import datetime
import logging
import configparser
import functions as fnc
from pathlib import Path

current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

current_root = Path().resolve()
# Create a folder if it doesn't exist
folder_path = "logs"
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

today_date = datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
# Set up logging
log_file = os.path.join(folder_path, f"query_analysis_{today_date}.log")
require_log_file = os.path.join(folder_path, f"require_query_{today_date}.log")

# Configure the logger for query_analysis.log
query_analysis_logger = logging.getLogger('query_analysis_logger')
query_analysis_logger.setLevel(logging.INFO)
query_analysis_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
query_analysis_handler.setFormatter(logging.Formatter('%(message)s'))
query_analysis_logger.addHandler(query_analysis_handler)

# Configure a separate logger for require_files_query_analysis.log
require_logger = logging.getLogger('require_files_logger')
require_logger.setLevel(logging.INFO)
require_handler = logging.FileHandler(require_log_file, mode='w', encoding='utf-8')
require_handler.setFormatter(logging.Formatter('%(message)s'))
require_logger.addHandler(require_handler)

config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
rootDir = Path(config['DEFAULT']['rootDir'])
rootTableListDir = current_root / Path(config['DEFAULT']['table_list_file_dir'])

table_folder_path = Path(config['DEFAULT']['table_list_file_dir'])


cfquery_pattern = re.compile(r'<cfquery\b.*?</cfquery>', re.DOTALL | re.IGNORECASE)
file_name = []

error_log_dir = current_root / Path(config['DEFAULT']['errorlogDir']) 
error_log_dir.mkdir(exist_ok=True)
read_file_error_log_path = error_log_dir / 'read_files_error.log' 
file_process_error_log_path = error_log_dir / 'file_process_error.log' 

def remove_all_comments(content):
    # Regular expression pattern to match nested HTML comments
    comment_pattern = re.compile(r'<!---(.*?)--->', re.DOTALL)
    nested_comment_pattern = re.compile(r'<!---(?:(?!<!---)(?:.|\\n))*?--->', re.DOTALL)
    
    # Remove nested comments
    cleaned_nested_content = nested_comment_pattern.sub('', content)
    cleaned_content = comment_pattern.sub('', cleaned_nested_content)

    return cleaned_content

def remove_sql_query_comment(content):
    # Remove all occurrences of C-style block comments (/* ... */)
    content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
    # Remove all occurrences of C++-style line comments (// ...)
    content = re.sub(re.compile("//.*?\n"), "", content)
    # Remove lines that start with '--'
    content = re.sub(re.compile("--.*?$", re.MULTILINE), "", content)
    
    # Remove ColdFusion tags
    content = re.sub(re.compile(r'<\/?cf[^>]*>', re.IGNORECASE), '', content)
    # Remove ColdFusion comments
    content = re.sub(re.compile(r'<!---.*?--->', re.DOTALL), '', content)
    # Remove ColdFusion expressions
    content = re.sub(re.compile(r'#[^#]*#'), '', content)
    
    return content

def scan_process_file(file_path):
    try:
        with file_path.open('r', encoding='utf-8') as file:
            file_contents = file.read()
            uncommented_contents = remove_all_comments(file_contents)  # Remove comments first
            matches = cfquery_pattern.findall(uncommented_contents)  # Find <cfquery> tags
            
            if matches:
                file_path_list = []
                file_name.append(file_path.name)
                query_analysis_logger.info(file_path)
                for match in matches:
                    #Remove <cfquery> tags from query
                    sql_regex = re.compile(r'<cfquery.*?>(.*?)</cfquery>', re.DOTALL)
                    result = sql_regex.search(match)
                    if result:
                        sql_query = result.group(1).strip()
                        is_select = fnc.has_select_query(sql_query)
                        is_insert = fnc.has_insert_query(sql_query)
                        is_update = fnc.has_update_query(sql_query)
                        is_delete = fnc.has_delete_query(sql_query)
                        column_map = {}
                        process_sql_query = remove_sql_query_comment(sql_query)
                        
                        if is_select:
                            query_pattern = fnc.validate_sql_pattern(process_sql_query)
                            if query_pattern == 'simple select':
                                column_map = fnc.extract_table_column_names(process_sql_query)
                            elif query_pattern == 'subquery select':
                                column_map = fnc.extract_table_column_names_with_sub_pat(process_sql_query)

                            elif query_pattern == 'Not sql exists':
                                if len(file_path_list) == 0:
                                    require_logger.info(file_path)
                                    file_path_list.append(file_path)

                                if file_path not in file_path_list:
                                    file_path_list = []
                                require_logger.info(process_sql_query + "\n")       
                            
                            for table, columns in column_map.items():
                                if 'join' in columns:
                                    if columns['join']:
                                        join_tables = ','.join(columns['join'])
                                        table += ',' + join_tables

                                query_analysis_logger.info(f"Select Table Name: {table}")                                
                                query_analysis_logger.info(f"Select Columns: {columns['select']}")
                                query_analysis_logger.info(f"Where Columns: {columns['where']}")
                                query_analysis_logger.info(f"Order Columns: {columns['order']}")
                                query_analysis_logger.info(f"Group Columns: {columns['group']}")
                                query_analysis_logger.info("===================================")
                        if is_insert:
                            table_name, columns = fnc.insert_table_column_names(process_sql_query)
                            if table_name and columns:
                                # Log table name and column names
                                query_analysis_logger.info("Insert Table Name: %s", table_name)
                                query_analysis_logger.info("Insert Columns: %s ", columns)
                                query_analysis_logger.info("===================================")
                        if is_update:
                            # Extract table name, set columns, and where columns
                            table_name, set_columns, where_columns = fnc.update_table_column_names(process_sql_query)
                            if table_name and set_columns and where_columns:
                                # Log table name, set columns, and where columns
                                query_analysis_logger.info("Update Table Name: %s", table_name)
                                query_analysis_logger.info("Set Columns: %s", set_columns)
                                query_analysis_logger.info("Where Columns: %s", where_columns)
                                query_analysis_logger.info("===================================")
                        if is_delete:
                            # Extract table name and where columns
                            table_name, where_columns = fnc.extract_delete_info(process_sql_query)
                            if table_name and where_columns:
                                # Log table name and where columns
                                query_analysis_logger.info("Delete Table Name: %s", table_name)
                                query_analysis_logger.info("Where Columns: %s", where_columns)
                                query_analysis_logger.info("===================================")
    except UnicodeDecodeError:
        with read_file_error_log_path.open('a', encoding='utf-8') as error_log:
            error_log.write(f'{current_time} - Error in scan_files.py/scan_process_file function: Unable to decode {file_path}\n')
    except Exception as e:
        with file_process_error_log_path.open('a', encoding='utf-8') as error_log:
            error_log.write(f'{current_time} - Error in scan_files.py/scan_process_file function: {e}\n')

def source_file_process(view_tables):
    files = [file for file in rootDir.rglob('*') if file.is_file()]
    for file in files:
        scan_process_file(file)
   
    fnc.start_run(log_file, rootTableListDir, table_folder_path, view_tables)
    return file_name

    
