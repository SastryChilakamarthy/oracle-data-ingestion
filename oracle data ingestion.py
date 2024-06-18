#!/usr/bin/env python
# coding: utf-8

# In[61]:


import pandas as pd
import cx_Oracle

# Load the XLSX file into a DataFrame
file_path = r'‪C:\Users\wissen\Downloads\houses.xlsx'
file_path = file_path.replace('\u202a', '').replace('\u202b', '')
df = pd.read_excel(file_path)

# Define Oracle connection details
hostname = 'DESKTOP-LBDBHDI'
port = '1521'
service_name = 'XE'  # Ensure this is the correct service name
username = 'SYSTEM'
password = '1234'

# Create a DSN (Data Source Name)
dsn_tns = cx_Oracle.makedsn(hostname, port, service_name=service_name)

# Establish the connection
connection = cx_Oracle.connect(user=username, password=password, dsn=dsn_tns)

# Prepare insert statement
table_name = 'YOUR_TABLE'  # Replace 'YOUR_TABLE' with your actual table name
columns = ', '.join([f'"{col}" VARCHAR2(100)' for col in df.columns])  # Ensure column definitions are properly formatted
placeholders = ', '.join([f':{i+1}' for i in range(len(df.columns))])
insert_columns = ', '.join([f'"{col}"' for col in df.columns])
insert_sql = f'INSERT INTO {table_name} ({insert_columns}) VALUES ({placeholders})'

print(f"Insert SQL: {insert_sql}")  # Debugging: Print the insert SQL statement

# Check if the table exists, if not, create it
cursor = connection.cursor()
try:
    cursor.execute(f"SELECT COUNT(*) FROM user_tables WHERE table_name = '{table_name.upper()}'")
    table_exists = cursor.fetchone()[0] == 1
    if not table_exists:
        # Create the table dynamically based on DataFrame columns
        create_table_sql = f"CREATE TABLE {table_name} ({columns})"
        print(f"Create Table SQL: {create_table_sql}")  # Debugging: Print the create table SQL statement
        cursor.execute(create_table_sql)
        print(f"Table '{table_name}' created successfully.")
except cx_Oracle.DatabaseError as e:
    error, = e.args
    print(f"Error: {error.message}")

# Insert DataFrame into Oracle table
try:
    for _, row in df.iterrows():
        cursor.execute(insert_sql, row.tolist())  # Execute the insert statement with row values
    print("Data inserted successfully.")
except cx_Oracle.DatabaseError as e:
    error, = e.args
    print(f"Error: {error.message}")

# Commit and close the connection
connection.commit()
cursor.close()
connection.close()


# In[ ]:





# In[ ]:




