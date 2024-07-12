import re

def insert_table_column_names(insert_query):
    # Regular expression to find <cfif ...> ... </cfif>
    cfif_pattern = re.compile(r"<cfif[^>]*>.*?</cfif>", re.DOTALL)

    # Skip queries with placeholders
    if "#tableName#" in insert_query or "#noColumnName#" in insert_query:
        return None, None
    
    # Extract table name
    table_name_start = insert_query.index("INTO") + len("INTO")
    table_name_end = insert_query.index("(")
    table_name = insert_query[table_name_start:table_name_end].strip()

    # Extract column names
    columns_start = insert_query.index("(") + 1
    columns_end = insert_query.index(")")
    columns_str = insert_query[columns_start:columns_end].strip()
    columns_name = cfif_pattern.sub('', columns_str)
    columns = [col.strip() for col in columns_name.split(",")
               ]
    return table_name,columns

