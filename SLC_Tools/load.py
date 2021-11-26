from pymysql.cursors import DictCursor
import xlrd
import pymysql
import os
import sys
from datetime import datetime
import db

# Open the workbook and define the worksheet
filename = r"C:/Ruta/donde/esta/archivo/SCRUM/archivo.xls"
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
print("************************************CONSTRUYENDO COMUNIDAD SCRUM LATAM*************************************")
print(" ")
print("Loading..................................")
print(" ")

number_country = 0

for row in range(0, sheet.nrows):

    email = sheet.cell(row, 0).value
    emailComparar = sheet.cell(row, 0).value
    print(emailComparar)
    name = sheet.cell(row, 1).value
    last_name = sheet.cell(row, 2).value

    webinar = sheet.cell(row, 3).value
    webinarComparar = sheet.cell(row, 3).value
    date = sheet.cell(row, 4).value
    member_type = sheet.cell(row, 5).value
    country = sheet.cell(row, 6).value

    id_C = 0
    id_today = datetime.now()
    id_random = int(id_today.strftime ('%m%d%H%S'))
    """ for row in lista:
        print(row + " + " + email)
        if row != email: """
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
        

    today = datetime.now()
    today_date = str(today.date())

    query = """INSERT INTO SLC_community_all (email, name, last_name, ID_country, Registry_date) SELECT %s, %s, %s, %s, %s WHERE NOT EXISTS (SELECT * FROM SLC_community_all WHERE email = '"""+emailComparar +"""')"""
    values = (email, name, last_name, id_C, today_date)
    cursor.execute(query, values)

    today_date_N = today_date + "_WTEST" #Se coloca W seguido de la fecha del webinar(EJ:2311)

    query_1 = """INSERT INTO SLC_webinar (id_webinar, webinar, date) SELECT %s, %s, %s WHERE NOT EXISTS (SELECT * FROM SLC_webinar WHERE webinar = '"""+webinarComparar +"""')"""
    values_1 = (today_date_N, webinar, date)
    cursor.execute(query_1, values_1)

    query_2 = """INSERT INTO SLC_Participants (id_webinar, email, member_type) SELECT %s, %s, %s WHERE NOT EXISTS (SELECT * FROM SLC_Participants WHERE id_webinar = '"""+today_date_N +"""' and email = '"""+emailComparar +"""')"""
    values_2 = (today_date_N, email, member_type)
    cursor.execute(query_2, values_2)    

    cursor.close()
    
    database.commit()

database.close()

print("")
print("Done! ")
print("")
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print("Acabo de cargar", columns, "columnas y", rows, "filas de datos de excel para MySQL!")
