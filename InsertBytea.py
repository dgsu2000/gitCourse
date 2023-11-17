import os
import sys
import psycopg2
from psycopg2 import sql
import base64


# Connect to the PostgreSQL database
conn1 = psycopg2.connect(database="dare1u",user="dare_uat", password="Passw0rd",host="10.68.2.32",port="5544")
cursor1 = conn1.cursor()

#sql_file = 'extraction_sql.sql'
file_path = '/tmp/IMG_1761.JPG'


def main():
    Open_DB_Connection()
    Process()
    Close_DB_Connection()

def Open_DB_Connection():
    sqltext = 'SELECT CURRENT_TIMESTAMP'
    cursor1.execute(sqltext)
        
    for now_text in cursor1:
        print("Now is:", now_text)

def Close_DB_Connection():
    conn1.close()	


def insert_pdf_into_database(file_path):
    try:
        # Read the binary data from the PDF file
        with open(file_path, 'rb') as file:
            pdf_data = file.read()

        # Insert the binary data into the database table
        insert_query = sql.SQL("INSERT INTO opom.test_blob (fname,fdata) VALUES ('2023020380667.pdf',%s)")

        cursor1.execute(insert_query, (pdf_data,))

        # Commit the transaction
        conn1.commit()
        print("PDF file inserted successfully.")
    except Exception as e:
        print(f"Error: {e}")

def extract_pdf_from_database(file_path):
    try:
        query = "select ATTACHFILE from ic.gr_reqattach where recid = 2017051373396"
        cursor1.execute(query)
        data = cursor1.fetchall()
        #file_binary=data[0][0].tobytes()
        #with open(file_path,'wb') as file:
        #    file.write(base64.b64decode(file_binary))

        with open(file_path,'wb') as file:
            file.write(data[0][0].tobytes())  #for BLOB
 
            
        print("PDF file extracted successfully.")
    except Exception as e:
        print(f"Error: {e}")
    
def Process():

    #insert_pdf_into_database(file_path)
    extract_pdf_from_database(file_path)
      
    sqltext = 'SELECT CURRENT_TIMESTAMP'
    cursor1.execute(sqltext)    
    for now_text in cursor1:
        print("End of Process Now is:", now_text)
    

if __name__ == '__main__':
    sys.exit(main())