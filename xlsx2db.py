import pandas as pd
from openpyxl import load_workbook 
import re
import glob
import os
import sys

eng = 'oracle+cx_oracle://ABC:1236@ABC.com:8080/DB' 
path='/'
merge_name = 'PRE_NAME'
extension = '.xlsx'

excel_files = glob.glob(os.path.join(path, "*"+extension)) 

for names in excel_files:

    output_name = (names.split("\\")[-1])
    name_of_sheet = "Data" #output_name.split(extension) [8]
    full_path = path + output_name
    table_name = merge_name+ output_name.split(".")[0][:4]
    print "\nFile Name:", output_name, ", Name of Sheet:", name_of_sheet 
    print "Table Name:", table_name, full path
    
    wb = load_workbook(full_path, read_only=True)
    ws = wb[name_of_sheet]
    
    data = ws.values
    
    # Get the first line in file as a header line
    columns = next(data)[0:]
    print columns
    1 = []
    
    # Cleaning columns using regular expression
    
    for c in columns:
        s= (re.sub ('[^A-Za-z8-9]+','', c)) [:30].upper()
        1.append(s)
    
    c = tuple(1)
    print c
    
    # Creating a DataFrame based on the second and subsequent lines of data
    df = pd.DataFrame (data, columns=c)
    df = df.replace('NA', '')
    
    #df = df.infer objects Q
    #print df
    
    from sqlalchemy import create_engine
    # THIS FOR REFERENCE AND NOT USED IN CODE
    credentials = {
    
    'username': 'abc',
    'password' '123',
    'host' : 'abc.com',
    'database' : 'db',
    'port' : '8080'}
    
    engine = create_engine(eng, encoding='UTF-8')
    
    from sqlalchemy.types import *
    
    print "Converting to DB..."
    # Write records stored in a DataFrame to a SQL database
    df.to_sql(table_name.lower(), con=engine, if_exists='replace', index=False, chunksize=180, dtype=String(
            Length=256))
    print "Table-", table_name, "added to database"
