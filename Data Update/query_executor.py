from api import AthenaQueryExecutor

def execute_query():
    query_executor = AthenaQueryExecutor()
    
    query= """
    
    """
    
    query_exec_id = query_executor.execute_query(query)
    query_status = None
    while query_status in (None, 'QUEUED', 'RUNNING'):
        query_status = query_executor.get_query_status(query_exec_id)
        print(query_status)
    
    if query_status == 'SUCCEEDED':
        query_executor.save_results_to_csv(query_exec_id, '<NAME-FILE>')
    else: 
        print('A consulta falhou')

if __name__ == "__main__":
    execute_query()