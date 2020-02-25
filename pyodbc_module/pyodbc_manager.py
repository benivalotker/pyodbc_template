import pyodbc
from connection_config import pyodbc_params	


'''
fast execute many statment  
    - insert_statement (string):    "INSERT INTO table_name (...) VALUES(?,?,?);"
    - values (list of tuples):      [ (row1), (row2), (row3) .... ] 
'''
def sql_insert(insert_statement, values):
    try:
        with pyodbc.connect('DRIVER={driver};SERVER={server};DATABASE={data_base};UID={uid};PWD={pwd};Trusted_Connection={trusted_conn};timeout=0;'
        .format(**pyodbc_params)) as conn:   

            cursor = conn.cursor()
            cursor.fast_executemany = True
            cursor.executemany(insert_statement, values)
            conn.commit()

    except pyodbc.OperationalError as ex:   # errors that are related to the database's operation and not necessarily under the control of the programmer
        pass
    except pyodbc.InterfaceError as ex:     #  errors that are related to the database interface rather than the database itself
        pass
    except Exception as ex:                 # other exception
        pass


'''
execute stored procedure 
    - sp (string):      "exec sp_name  @param_name=?;"
    - values (tuple):   (ex, ) 
'''
def run_sp(sp, values=None):
    try:
        with pyodbc.connect('DRIVER={driver};SERVER={server};DATABASE={data_base};UID={uid};PWD={pwd};Trusted_Connection={trusted_conn};timeout=0;'
        .format(**pyodbc_params)) as conn: 

            cursor = conn.cursor()
            cursor.fast_executemany = True
            cursor.execute(sp, values)
            conn.commit()

    except pyodbc.OperationalError as ex:   # errors that are related to the database's operation and not necessarily under the control of the programmer
        pass
    except pyodbc.InterfaceError as ex:     #  errors that are related to the database interface rather than the database itself
        pass
    except Exception as ex:                 # other exception
        pass

