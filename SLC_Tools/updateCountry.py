from pymysql.cursors import DictCursor
import xlrd
import pymysql
import os
import sys
from datetime import datetime
import db


# Open the workbook and define the worksheetp
filename = r"C:/Ruta/donde/esta/archivo/SCRUM/archivo_pais.xls"
if not os.path.exists(filename):
    print("No encontré el archivo")
    sys.exit()

book = xlrd.open_workbook(filename)
sheet = book.sheet_by_name("Hoja1")
lista = []


database = pymysql.connect(
    host = db.app.config['MYSQL_DATABASE_HOST'],
    user = db.app.config['MYSQL_DATABASE_USER'],
    passwd = db.app.config['MYSQL_DATABASE_PASSWORD'],
    database = db.app.config['MYSQL_DATABASE_DB']
)

print(" ")
print("**********************CONSTRUYENDO COMUNIDAD SCRUM LATAM (updates de países)*************************************")
print(" ")
print("Loading..................................")
print(" ")

number_country = 0

for row in range(0, sheet.nrows):

    email = sheet.cell(row, 0).value
    emailComparar = sheet.cell(row, 0).value
    print(emailComparar)
    country = sheet.cell(row, 1).value

    id_C = 0
    id_today = datetime.now()
    id_random = int(id_today.strftime ('%m%d%H%S'))
  
    cursor = database.cursor()
    query_country = """SELECT ID_Country FROM SLC_Country WHERE Country ='"""+ country +"""'"""
    cursor.execute(query_country)
    record = cursor.fetchall()
   
    for row in record:
        id_C = int(row[0])
    
    if not id_C:
        query_insert = """INSERT INTO SLC_Country (ID_Country, Country) SELECT %s,%s WHERE NOT EXISTS (SELECT * FROM SLC_Country WHERE Country = '"""+ country +"""')"""
        values_insert = (id_random, country)
        cursor.execute(query_insert, values_insert)
        query_country = """SELECT ID_Country FROM SLC_Country WHERE Country ='"""+ country +"""'"""
        cursor.execute(query_country)
        record = cursor.fetchall()
    
        for row in record:
            id_C = int(row[0])
        
    query = """UPDATE SLC_community_all SET ID_Country = %s WHERE email = %s and ID_Country = 11131748"""
    values = (id_C, email)
    cursor.execute(query, values)

    cursor.close()
    
    database.commit()

database.close()

print("")
print("Done! ")
print("")
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print("Acabo de cargar", columns, "columnas y", rows, "filas de datos de excel para MySQL!")
