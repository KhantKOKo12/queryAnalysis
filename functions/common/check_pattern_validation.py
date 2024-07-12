import re

def has_select_query(query):
    select_pattern = re.compile(r'\bselect\b', re.IGNORECASE)
    return bool(select_pattern.search(query))

def has_insert_query(query):
    insert_pattern = r'^\s*INSERT\s+INTO\b'
    return bool(re.match(insert_pattern, query, re.IGNORECASE))

def has_update_query(query):
    update_pattern = r'^\s*UPDATE\b'
    return bool(re.match(update_pattern, query, re.IGNORECASE))

def has_delete_query(query):
    delete_pattern = r'^\s*DELETE\s+FROM\b'
    return bool(re.match(delete_pattern, query, re.IGNORECASE))


def is_table_name_present(query):
    table_search = r'\bFROM\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*(?:AS\s+([a-zA-Z_][a-zA-Z0-9_]*))?'
    join_table_name_regex = r'(JOIN)\s+(.+?)\s+(ON|$)'
    from_table = re.search(table_search, query, re.IGNORECASE)
    join_table = re.findall(join_table_name_regex, query, re.IGNORECASE) 
   
    if from_table:
        table_name = from_table.group(1)
        if table_name.upper() == 'AS' or table_name.upper() == 'WHERE' or len(table_name) == 1 and len(join_table) == 0:
            return False
        else:
            return True
    return False

def is_valid_join_table_name(word):
    table_name_pattern = re.compile(r"^[a-zA-Z_][\w]*$")
    return bool(table_name_pattern.match(word))

def validate_sql_pattern(query):
    # Check for exactly one occurrence of SELECT/select (case-insensitive)
    select_pattern = re.compile(r'\bselect\b', re.IGNORECASE)
    select_matches = select_pattern.findall(query)
    
    # Check for absence of JOIN/join (case-insensitive)
    # join_pattern = re.compile(r'\bjoin\b', re.IGNORECASE)
    # join_matches = join_pattern.findall(query)

    # Validate conditions
    query_pattern = ''
    if len(select_matches) == 1:
        query_pattern = 'simple select'
    elif len(select_matches) > 1:
        query_pattern = 'subquery select'
    else:
        query_pattern = 'Not sql exists' 
    
    return query_pattern

# Function to extract column names
def extract_column_name(col):
    # Split by dot and take the last part
    parts = col.split('.')
    column_name = parts[-1]
    return column_name