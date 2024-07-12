import re
import functions as fnc

def extract_delete_info(delete_query):
    # Regular expressions to extract table name and where columns
    table_name_pattern = r'DELETE\s+FROM\s+(\w+)'
    where_columns_pattern = r'WHERE\s+(.*)'
    # Regular expression to split by "=", "IN", or "NOT IN"
    split_pattern = r'\s*(=|IN|NOT IN|>)\s*'
    # Regular expression to split by "AND"
    and_pattern = r'\s+AND\s+'
    # Remove "AND NOT (" and replace it with "AND ("
    and_not_pattern = r'\s+AND\s+NOT\s*\('

    # Skip queries necessary
    if "<cfif LOOPCOUNT EQ 1>" in delete_query:
        return None, None

    # Extract table name
    table_match = re.search(table_name_pattern, delete_query, re.IGNORECASE)
    if not table_match:
        return None, None
    table_name = table_match.group(1)

    # Extract where columns
    where_columns_match = re.search(where_columns_pattern, delete_query, re.IGNORECASE | re.DOTALL)
    if not where_columns_match:
        return table_name, None
    where_columns_str = where_columns_match.group(1)
    where_columns_str_modified = re.sub(and_not_pattern, ' AND (', where_columns_str)
    where_columns = [fnc.extract_column_name(re.split(split_pattern, col)[0].strip().lstrip('(')) for col in re.split(and_pattern, where_columns_str_modified)]
    
    return table_name, where_columns
