import pandas as pd
from sqlalchemy import create_engine
import time



def Excel2sqlAttr(file_name, seconds):
    server = 'localhost'
    database = 'Database1'

    # create a SQLalchemy engine
    engine = create_engine(f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server")
    # Define the table name
    table_name = 'real_time_table'
    #read the excel file into dataframe
    df = pd.read_excel(file_name)
    #make new empity dataframe
    df_output = pd.DataFrame()
    for index, row in df.iterrows():
        #add new row every cycle
        df_output = df_output.append(row, ignore_index=True).tail(1)
        #update the database
        if index==0:
            df_output.to_sql(table_name, engine, if_exists='replace', index=False)
        else:
            df_output.to_sql(table_name, engine, if_exists='append', index=False)
        #slow the process
        time.sleep(seconds)

time.sleep(2)
print ('important note: your excel file must have only one sheet')
time.sleep(2)
print('don\'t forget to add extension of your file | example "your file name.xlsx"')
time.sleep(2)
file= input('Enter the name of your Excel sheet...')
time.sleep(2)
seconds=input('Enter the time(seconds) between adding new rows to the database  (the speed of updating)...')
time.sleep(1)
print('updating start')
Excel2sqlAttr(file_name =file, seconds=int(float(seconds)))
print('finished')
time.sleep(3)
