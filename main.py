import pandas as pd
import os
import psycopg2
from sqlalchemy import create_engine
import shutil
from urllib.parse import quote_plus
 
# Function to load Excel files and insert data into PostgreSQL table
def import_excel_to_postgres(excel_files, table_name, connection_uri):
    # Create SQLAlchemy engine
    engine = create_engine(connection_uri)
   
    # Loop through each Excel file
    for file in excel_files:
 
        try:
            # Load Excel file into pandas DataFrame
            df = pd.read_excel(os.path.join(folder_path, file))
             # Insert data into PostgreSQL table
            df.to_sql(table_name, engine, if_exists='append', index=False)
           
            os.remove(os.path.join(folder_path, file))
 
        except Exception as e:
            print(f"Error processing '{file}': {e}")
 
def get_excel_files(folder_path):
    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            excel_files.append(os.path.join(folder_path, file))
    return excel_files
 
# Folder path containing Excel files
#folder_path = r"C:\Users\manmohan.d\OneDrive - SAKSOFT LIMITED\Desktop\backup\Subash\executeKL"
folder_path = r"C:\Users\manmohan.d\OneDrive - SAKSOFT LIMITED\Desktop\backup\Subash\upload"

#destination_folder_path = r"C:\Users\subash.kb\Desktop\OneDrive_2024-04-18\Aggregated Level 1\Uploaded Files"
 
# Get all Excel files from the folder
excel_files = get_excel_files(folder_path)
#db_password = "Spi@123"
db_password = "ciYlRgEm"
encoded_password = quote_plus(db_password)
# PostgreSQL connection URI
#connection_uri = 'postgresql://sipuser:{db_password}@172.25.1.23:5432/sip_dev'
#connection_uri =f"postgresql://sipuser:{encoded_password}@172.25.1.23:5432/sip_dev"

#PROD
#connection_uri =f"postgresql://sipuser:{encoded_password}@sipwebapp-prod.cfuq8m22uktz.ap-south-1.rds.amazonaws.com:5432/sip_prod"

#UAT
connection_uri =f"postgresql://sipuser:{encoded_password}@sipwebapp-uat.czfoko8d3qhs.ap-south-1.rds.amazonaws.com:5432/sip_uat"
         
 
# Table name in PostgreSQL
table_name = 'TempFranchisee_OD'
#table_name = 'Temp_1722_Discontinued_students_30082024_KL'

 
# Import Excel files to PostgreSQL
import_excel_to_postgres(excel_files, table_name, connection_uri)