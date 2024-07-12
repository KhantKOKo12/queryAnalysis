import re
from collections import defaultdict
import functions as fnc
import sqlparse
from sqlparse.sql import IdentifierList, Identifier
from sqlparse.tokens import Keyword, DML, Punctuation,Wildcard

subqueries = []
# Function to recursively extract and print CASE statements
def extract_case_statements(token_list):
    case_statements = []
    for token in token_list:
        if isinstance(token, sqlparse.sql.Case):
            case_statements.append(token.value.strip())
        elif isinstance(token, sqlparse.sql.TokenList):
            case_statements.extend(extract_case_statements(token.tokens))
        elif isinstance(token, sqlparse.sql.IdentifierList):
            for identifier in token.get_identifiers():
                case_statements.extend(extract_case_statements(identifier.tokens))
        elif isinstance(token, sqlparse.sql.Identifier):
            case_statements.extend(extract_case_statements(token.tokens))
        elif isinstance(token, sqlparse.sql.Function):
            case_statements.extend(extract_case_statements(token.tokens))
    return case_statements

def remove_case_statements(token_list):
    new_query_parts = []
    for token in token_list:
        if isinstance(token, sqlparse.sql.Case):
            continue  # Skip CASE statements
        elif isinstance(token, sqlparse.sql.TokenList):
            new_query_parts.extend(remove_case_statements(token))
        else:
            new_query_parts.append(token)
    return new_query_parts

def case_when_pattern_analysis(case_statements):
    columns = []
    case_when_columns = []
    column_pattern = r'\b(?:[a-zA-Z_][a-zA-Z0-9_]*|[ぁ-んァ-ヶ亜-熙訁-鿋０-９ー]+)\.(?:[a-zA-Z_][a-zA-Z0-9_]*|[ぁ-んァ-ヶ亜-熙訁-鿋０-９ー]+)\b'

    if case_statements:
        for case_statement in case_statements:
            columns = re.findall(column_pattern, case_statement)
            case_when_columns.append(columns)

    return case_when_columns

def extract_where_column_names(sql_query,is_group_by_exit,is_order_by_exit):
    not_order_group = False
    if is_order_by_exit:
        where_clause_regex = re.compile(r'WHERE\s+(.*?)(ORDER\s+BY|$)', re.IGNORECASE | re.DOTALL)
    elif is_group_by_exit:
        where_clause_regex = re.compile(r'WHERE\s+(.*?)(GROUP\s+BY|$)', re.IGNORECASE | re.DOTALL)
    else :
        where_clause_regex = re.compile(r'WHERE\s+(.*)', re.IGNORECASE | re.DOTALL)
        not_order_group = True
    
    where_clause = where_clause_regex.findall(sql_query)
    where_columns = ''
    
    if where_clause:
        if not_order_group == False:
            where_columns = where_clause[0][0]
        else:
            where_columns = where_clause[0]

    return where_columns.lower() 

def remove_row_number_segment(sql_query):
    order_by_regex = re.compile(r'ORDER\s+BY\s+(.*?)(?=\))', re.IGNORECASE | re.DOTALL)
    pattern1 = re.compile(r"ROW_NUMBER\(\) OVER\(ORDER BY .*?\) AS rownum,", re.IGNORECASE | re.DOTALL)
    pattern2 = re.compile(r'ROW_NUMBER\(\) OVER\(.*?\) AS rownum', re.IGNORECASE | re.DOTALL)
    pattern3 = re.compile(r'(ROW_NUMBER\(\) OVER \(ORDER BY .+?\) AS rownum\s*,?\s*)', re.IGNORECASE | re.DOTALL)
    pattern4 = re.compile(r'ROW_NUMBER\(\) OVER\s*\((.*?)\)\s*AS\s+(\w+)', re.IGNORECASE | re.DOTALL)
    
    result1 = pattern1.findall(sql_query)
    result2 = pattern2.findall(sql_query)
    result3 = pattern3.findall(sql_query)
    result4 = pattern4.findall(sql_query)
    modified_query = sql_query
    if result1:
        order_by_clouds = order_by_regex.search(result1[0])
        modified_query = pattern1.sub('', sql_query)
        modified_query += order_by_clouds.group(0)  
    elif result2:
        order_by_clouds = order_by_regex.search(result2[0])
        modified_query = pattern2.sub('', sql_query)
        if order_by_clouds:
            modified_query += order_by_clouds.group(0)   
    elif result3:
        modified_query = pattern3.sub('', sql_query)
        
    elif result4:
        modified_query = pattern4.sub('', sql_query)

    return modified_query

def anlaysis_select_clause_column_conditions(col):
    column = re.sub(r"'([^']*)'", r"\1", col)
    column = re.sub(r'\s+AS\s+\w+', '', column, flags=re.IGNORECASE).strip()
    column = re.sub(r'^As\s\w+$', '', column, flags=re.IGNORECASE).strip()
    column = re.sub(r"\s+AS\s+[a-zA-Z_]\w*", "", column, flags=re.IGNORECASE)
    column = re.sub(r'[-+*/]\s*\d+(\.\d+)?', '', column)

    return column.strip()

def contains_select_all(query):
    parsed = sqlparse.parse(query)[0]
    select_seen = False
    for token in parsed.tokens:
        # Skip comments and whitespaces
        if token.is_whitespace or token.ttype in [Keyword, DML] and token.value.upper() == 'SELECT':
            select_seen = True
        elif select_seen:
            # Look for identifiers after SELECT
            if isinstance(token, IdentifierList):
                for identifier in token.get_identifiers():
                    if identifier.ttype == Wildcard:
                        return True
            elif isinstance(token, Identifier) and token.ttype == Wildcard:
                return True
            elif token.ttype == Wildcard:
                return True
            # Stop checking once we hit the FROM clause
            elif token.ttype is Keyword and token.value.upper() == 'FROM':
                break
    return False

def extract_table_and_alias(query):
    table_name_regex = re.compile(r'\bFROM\b\s+(.*)', re.IGNORECASE)
    stop_keywords = re.compile(r'\b(?:WHERE|\s+INNER|\s+LEFT|\s+FULL|GROUP\s+BY|ORDER\s+BY)\b', re.IGNORECASE)

    match = table_name_regex.findall(query)
    tables = []
    if match:
        table_names_part = stop_keywords.split(match[0])

        for table_def in table_names_part[0].split(','):
            table_def = ' '.join(table_def.strip().split())
            table_def = table_def.replace(';','')
            full_table_name = replace_as_for_table_name(table_def)
            tables.append(full_table_name)
    
    output = ",".join(tables)

    return output

def replace_as_for_table_name(table_name):
    split_table_names = table_name.split(' ')
    if len(split_table_names) == 2 and not table_name.upper().startswith("AS"):
        table_name = table_name.replace(' ', ' AS ')
    return table_name


def extract_table_column_names(sql_query):
    distinct_regex = re.compile(r'\bDISTINCT\b', re.IGNORECASE)
    sql_query = distinct_regex.sub('', sql_query)
    parsed_query = sqlparse.parse(sql_query)[0]
    new_tokens = remove_case_statements(parsed_query.tokens)
    sql_query = ''.join(str(token) for token in new_tokens)
    case_statements = extract_case_statements(parsed_query)
    case_when_columns = case_when_pattern_analysis(case_statements)

    sql_query = remove_row_number_segment(sql_query)
    sql_query = remove_parentheses_and_brackets(sql_query)
    
    join_table_name_regex = re.compile(r'(JOIN)\s+(.+?)\s+(ON|$)', re.IGNORECASE)
    select_table_name = extract_table_and_alias(sql_query)

    # vaildation_for_view_table(select_table_name, view_tables)  
    join_table_names = join_table_name_regex.findall(sql_query)


    column_map = defaultdict(lambda: {'select': [], 'where': [], 'join': [], 'order': [], 'group': []})
    if join_table_names:
        full_join_table_names = []
        if select_table_name.upper().startswith("AS") or select_table_name.upper().startswith("WHERE"):
            join_table_names = [item[1] for item in join_table_names]
            select_table_name = ''
            for join_table_name in join_table_names:
               full_table_name = replace_as_for_table_name(join_table_name)
               full_join_table_names.append(full_table_name)

            select_table_name = ','.join(full_join_table_names)
        else:
            for table_name in join_table_names:
                full_table_name = replace_as_for_table_name(table_name[1])
                column_map[select_table_name]['join'].append(full_table_name)  
    else:
        if select_table_name.upper().startswith("AS") or select_table_name.upper().startswith("WHERE"):
            return column_map    
                    
    column_list = []

    if select_table_name:
        for col_list in case_when_columns:
            for col in col_list:
                column = anlaysis_select_clause_column_conditions(col)
                column_map[select_table_name]['select'].append(column)
    
        select_clause_regex = re.compile(r'SELECT\s+(.*?)\s+FROM', re.IGNORECASE | re.DOTALL)
        select_clause = select_clause_regex.findall(sql_query)

        if select_clause:
            select_column = split_columns(select_clause[0])

            if select_column:
                for col_match in select_column:
                
                    column = anlaysis_select_clause_column_conditions(col_match)
                    column_map[select_table_name]['select'].append(column)   


    split_regex = re.compile(r'\b(?:LEFT\s+OUTER|INNER|RIGHT|LEFT|FULL\s+OUTER)?\s+JOIN\b', re.IGNORECASE)

    # Define the regex pattern to identify keywords where, group by, order by to truncate the query
    stop_keywords = re.compile(r'\b(?:WHERE|GROUP\s+BY|ORDER\s+BY)\b', re.IGNORECASE)
    query = stop_keywords.split(sql_query)[0]
    join_clauses = split_regex.split(query)

    if len(join_clauses) > 0:
        for join_clause in join_clauses[1:]:
            pattern = re.compile(r'(\b[\w\.]+)\s*=\s*([\w\.]+\b)')
            matches = pattern.findall(join_clause)

            for match in matches:
                if match[0] not in column_list:
                    column_map[select_table_name]['select'].append(match[0])
                if match[1] not in column_list:
                    column_map[select_table_name]['select'].append(match[1])

    # Extract WHERE clause columns
    asc_desc_regex =  re.compile(r"(?i)\b ASC\b|\b DESC\b", re.IGNORECASE | re.DOTALL)
    no_asc_desc_query = re.sub(asc_desc_regex, '', sql_query)
    order_by_regex = re.compile(r'ORDER\s+BY\s+(.*)', re.IGNORECASE | re.DOTALL)
    order_by_clauses = order_by_regex.findall(no_asc_desc_query)

    group_by_regex = re.compile(r'GROUP\s+BY\s+(.*)', re.IGNORECASE | re.DOTALL)
    group_by_clauses = group_by_regex.findall(sql_query)
    
    stop_keywords_for_where = re.compile(r'\b(?:GROUP\s+BY|ORDER\s+BY)\b', re.IGNORECASE)
    where_query = stop_keywords_for_where.split(sql_query)[0]
    result = extract_where_column_names(where_query,order_by_clauses,group_by_clauses)

    where_regex = re.compile(r'([\w\.]+)\s*(?:=|<>|!=|>=|<=|>|<|\s+NOT IN|\s+IN|\s+NOT LIKE|\s+LIKE)\s*([^=\n]*?)\s*$')

    between_regex = re.compile(r"(.*?)\s+BETWEEN\s+(.*?)\s+AND\s+(.*)")
    columns_with_where = []
    columns_with_between = between_regex.findall(result)
    no_between_sql = between_regex.sub('',result)

    split_where_columns_for_where = no_between_sql.upper().split('WHERE')
    if len(split_where_columns_for_where) > 1:
       no_between_sql = ' AND '.join(s.strip() for s in split_where_columns_for_where)

    # split_where_columns_for_and = no_between_sql.upper().split('AND'|'OR')
    split_where_columns = re.split(r'\sAND\s|\sOR\s|\sFOR\s', no_between_sql.upper())
    for where_column in split_where_columns:
        cols = where_regex.findall(where_column.strip())
        if cols:
            col1 = cols[0][0]
            col2 = cols[0][1]
            if not (col1.startswith("'") and col1.endswith("'")):
                columns_with_where.append(col1.strip().lower())
            if not (col2.startswith("'") and col2.endswith("'")):
                columns_with_where.append(col2.strip().lower())


    split_between_regex = re.compile(r"\b(\w+\.\w+)\b")

    # Extracting the matching parts
    extracted_data = []
    for col in columns_with_between:
        matches = split_between_regex.findall(col[0])
        if matches:
            extracted_data.extend(matches)

    for column in columns_with_where:
        column_map[select_table_name]['where'].append(column.strip())

    for column in extracted_data:
        column_map[select_table_name]['where'].append(column.strip())
   
    # Extract Order By Or Group By clause columns    
    if order_by_clauses:
        for column in order_by_clauses:
            order_column_regex = re.compile(r'(\b[\w.]+)\s*(=|!=|<>|>|<|>=|<=|IS NULL|IS NOT NULL)\s*', re.IGNORECASE)
            split_order_regex = r'\bAND\b|\bOR\b|\bORDER BY\b|,'
            col_parts = re.split(split_order_regex, column, flags=re.IGNORECASE)
            if col_parts:
                for col in col_parts:
                    col_results = order_column_regex.search(col.strip())

                    if col_results:
                        col_result = col_results.group(1).replace(')', '')
                        column_map[select_table_name]['order'].append(col_result.strip())
                    else:
                        col = col.replace(')', '')
                        col = col.replace('\n', '')
                        col = col.replace(',', '')
                        column_map[select_table_name]['order'].append(col.strip())
            else:
                column = column[0].replace(')', '')
                column = column.replace('\n', '')
                column = column.replace(',', '')
                column_map[select_table_name]['order'].append(column.strip())        
      
    if contains_select_all(sql_query):           
        if group_by_clauses:
            for column in group_by_clauses:
                column = column.replace(')', '')
                column = column.strip().replace(' ', '').replace('\n', '').replace('\t', '')
                column = column.replace('havingcount*>1', '')
                column_map[select_table_name]['group'].append(column)
                
    column_map = ensure_unique_values(column_map)
    return column_map

def split_columns(select_clause):
    columns = []
    bracket_level = 0
    current_column = []
    
    for char in select_clause:
        if char == ',' and bracket_level == 0:
            columns.append(''.join(current_column).strip())
            current_column = []
        else:
            if char == '{':
                bracket_level += 1
            elif char == '}':
                bracket_level -= 1
            current_column.append(char)
    
    if current_column:
        columns.append(''.join(current_column).strip())
    
    return columns
        
def ensure_unique_values(data_dict):
    for table, columns in data_dict.items():
        for key, values in columns.items():
            seen = set()
            unique_ordered_values = []
            for value in values:
                value = filter_built_in_functions_from_column(value)
                split_values = value.lower().split('+')
                
                if len(split_values) > 1:
                    for value in split_values:
                        value = value.replace("=", '')
                        value = value.replace("'", "")
                        value = value.strip()
                        if not value.isdigit() and value and not value.lower() == "count *" and not value.lower() == "and" and not value.isdigit() and value.strip() and value != "''" and value != "':'" and value != ":" and value != 'as' and value not in seen:
                            seen.add(value)
                            unique_ordered_values.append(value)
                else:
                    value = value.replace("=", "")
                    value = value.replace("'", "")
                    value = value.strip()
                    if not value.isdigit() and value and not value.lower() == "count *" and not value.lower() == "and" and not value.isdigit() and value.strip() and value != "''" and value != "':'" and value != ":" and value != 'as' and value not in seen:
                        seen.add(value) 
                        unique_ordered_values.append(value)          
            data_dict[table][key] = unique_ordered_values

    return data_dict

def filter_built_in_functions_from_column(column):
    built_in_functions_regex = re.compile(
    r"\b(SOME\s+|COUNT\s+|NOT LIKE|SUM\s+|AVERAGE\s+|MIN\s+|MAX\s+|CAST\s+|IS NULL\s+|ISNULL\s+|MONTH\s+|DAY\s+|LEFT\s+|ISNUMERIC\s+|LTRIM\s+|\s+STR\s+|SUBSTRING\s+|datetime\s+|STR\s+|CONVERT\s+|VARCHAR\s+|TOP\s+\d+)\b|\bTOP\s*\(\s*\d+\s*\)",
    re.IGNORECASE
    )
    dateformat_regex = re.compile(r"dateformat{\s*([^,']+)", re.IGNORECASE | re.DOTALL)
    no_built_in_col = built_in_functions_regex.sub("", column)
    if no_built_in_col:
        match = re.search(dateformat_regex, no_built_in_col)
        if match:
           column = match.group(1).strip()
        else:
           column = no_built_in_col
    else:
        match = re.search(dateformat_regex, column)
        if match:
           column = match.group(1).strip() 

    return column.lower()
     

def combine_data_list(data_list):
    remove_pattern = re.compile(r'\b\d+\+[^\s]+\b', re.IGNORECASE | re.DOTALL)
    combined_data = defaultdict(lambda: defaultdict(list))
    
    key_name = ''
    for data_dict in data_list:
        for key in data_dict:
            if key_name == '':
                key_name = key
            else:   
                key_name += ',' + key   

    for data_dict in data_list:        
        for alias, columns_info in data_dict.items():
            for key, value in columns_info.items():
                cleaned_values = [remove_pattern.sub('', v) for v in value]
                if cleaned_values:
                    combined_data[key_name][key].extend(cleaned_values)
                else:    
                    combined_data[key_name][key].extend(value)

    combined_data[key_name]['select'] = list(set(combined_data[key_name]['select']))
    combined_data[key_name]['where'] = list(set(combined_data[key_name]['where']))
    combined_data[key_name]['order'] = list(set(combined_data[key_name]['order']))
    combined_data[key_name]['group'] = list(set(combined_data[key_name]['group']))
    column_map = ensure_unique_values(combined_data)
    return column_map

def remove_parentheses_and_brackets(content):
    # Compile patterns that match parentheses and square brackets
    parentheses_pattern = re.compile(r'[\(\)]')
    brackets_pattern = re.compile(r'\[(.*?)\]')
    
    # Substitute all matches with their content, removing the brackets
    content = parentheses_pattern.sub(' ', content)  # Remove spaces around parentheses

    content = brackets_pattern.sub(r'\1', content)

    return content
                
def is_subselect(parsed):
    if not parsed.is_group:
        return False
    for item in parsed.tokens:
        if item.ttype is DML and item.value.upper() == 'SELECT':
            return True
    return False

def extract_subqueries(tokens):
    """
    Recursively extract subqueries from the token list.
    """
    subqueries = []
    for token in tokens:
        if is_subselect(token):
            subqueries.append(token)
            subqueries.extend(extract_subqueries(token.tokens))
        elif token.is_group:
            subqueries.extend(extract_subqueries(token.tokens))
    return subqueries

def extract_nested_subqueries(sql):
    """
    Parse the SQL and extract all nested subqueries.
    """
    parsed = sqlparse.parse(sql)
    statement = parsed[0]
    return extract_subqueries(statement.tokens)

def remove_nested_subqueries(remove_subqueries,subquery_str):
    for remove_subquery in remove_subqueries:
        remove_subquery = str(remove_subquery).strip()
        
        if remove_subquery != subquery_str:  
            remove_subquery = remove_subquery
            subquery_str = subquery_str.replace(remove_subquery, "")
            
    return subquery_str

def generate_sql_fragments(sql):
    # Function to extract nested subqueries (not implemented here)
    nested_subqueries = extract_nested_subqueries(sql)
    
    # Regular expression to match 'SELECT' keyword
    select_regex = re.compile(r'\bselect\b', re.IGNORECASE)
    
    # Append the main SQL query to the nested subqueries list
    nested_subqueries.append(sql)
    
    sql_fragments = []
    
    # Iterate through each subquery (including the main SQL query)
    for subquery in nested_subqueries:
        subquery_str = str(subquery).strip()
        
        # Count occurrences of 'SELECT' in the subquery
        select_matches = select_regex.findall(subquery_str)
        
        # If there is not exactly one 'SELECT', handle the subquery
        if len(select_matches) != 1:
            # Assuming remove_nested_subqueries is properly defined elsewhere
            subquery_str = remove_nested_subqueries(nested_subqueries, subquery_str)
        
        # Add unique subquery to sql_fragments list
        if subquery_str not in sql_fragments:
            sql_fragments.append(subquery_str)
            # Remove subquery from main SQL string to avoid duplication
            sql = sql.replace(subquery_str, '')
    
    # Return list of unique SQL fragments
    return sql_fragments

def generate_sql_fragments(sql):
    nested_subqueries = extract_nested_subqueries(sql)
    select_regex = re.compile(r'\bselect\b', re.IGNORECASE)
    nested_subqueries.append(sql)
    sql_fragments = []
    
    for subquery in nested_subqueries:
        subquery_str = str(subquery).strip()    
        select_matches = select_regex.findall(subquery_str)

        if len(select_matches) != 1:     
            subquery_str = remove_nested_subqueries(nested_subqueries,subquery_str)  

        if subquery_str not in sql_fragments:
            sql_fragments.append(subquery_str)
            sql = sql.replace(subquery_str, '')
    return sql_fragments


def remove_invalid_joins(query):
    join_patterns = [
        # LEFT JOIN
        r"LEFT\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?\s*ON\s*\(([\s\S]*?)\)",
        # LEFT OUTER JOIN
        r"LEFT\s+OUTER\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?\s*ON\s*\(([\s\S]*?)\)",
        # INNER JOIN
        r"INNER\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?\s*ON\s*\(([\s\S]*?)\)",
        # RIGHT JOIN
        r"RIGHT\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?\s*ON\s*\(([\s\S]*?)\)",
        # RIGHT OUTER JOIN
        r"RIGHT\s+OUTER\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?\s*ON\s*\(([\s\S]*?)\)",
        # OUTER JOIN
        r"OUTER\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?\s*ON\s*\(([\s\S]*?)\)",
        # CROSS JOIN
        r"CROSS\s+JOIN\s+(\w+)?(\s+AS\s+\w+)?",
    ]

    # Compile the join patterns with case insensitivity
    join_regex = re.compile("|".join(join_patterns), re.IGNORECASE)
    
    sql_query = join_regex.sub(validate_and_keep, query)
    
    return sql_query

def validate_and_keep(match):
    table_name = match.group(1)
    if table_name and fnc.is_valid_join_table_name(table_name):
        return match.group(0)  # Keep the valid JOIN clause
    return ""  # Remove the invalid JOIN clause

def extract_table_column_names_with_sub_pat(sql_query):
    column_map = {}
    temp_column_map = []
    sql_fragments = generate_sql_fragments(sql_query)

    if sql_fragments:       
        for subquery in sql_fragments:
            split_queries = re.split(r'\bUNION\s+ALL\b|\bUNION\b', subquery, flags=re.IGNORECASE)
            if split_queries:
                for query in split_queries:
                    if fnc.is_table_name_present(query):
                        no_invaild_join_query = remove_invalid_joins(query)
                        column_map1 = extract_table_column_names(no_invaild_join_query)
                        temp_column_map.append(column_map1)     

    if temp_column_map:
        temp_column_map.append(column_map)
        column_map = combine_data_list(temp_column_map)
    subqueries.clear()

    return column_map
