# -*- coding: utf-8 -*-


import psycopg2
import psycopg2.extensions
import pyodbc
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)

#===========================================================================================
#Writer : Arda Çetinkaya
#Initiation Date : 16-12-2017 13:50
#===========================================================================================

#DB Parameters==============================================================================
host = 'localhost'
user = 'postgres'
password = 'kalman'
dbName = 'DataCollecitonDB'
#DB Parameters==============================================================================

#Definitions================================================================================


#ConnString=================================================================================
try:
    conn = psycopg2.connect("dbname="+dbName+" user="+user+" host="+host+" password="+password)
    conn.set_client_encoding('UTF-8')
    print "Connected Succesfully!"
except BaseException as Be:
    print "I am unable to connect to the database"
    print Be.message
#ConnString=================================================================================



#TheCode-PSQL QUERY=========================================================================
try:
    cur = conn.cursor()
    cur.execute("""SELECT * from yapidata""")
    rows = cur.fetchall()
    # for row in rows:
    #     print row[0]
except BaseException as Be:
    print Be.message

#TheCode-MDB QUERY=========================================================================
try:
    mdb = "A:\kaip_ornek_parsel\TASKINPASA_KAIP.mdb"
    drv = "{Microsoft Access Driver (*.mdb)}"
    PWD = "pw"

    mdb_con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
    mdb_cur = mdb_con.cursor()
    mdb_rows = cur.execute('Select * from BINA').fetchall()
    mdb_cur.close()
    mdb_con.close()

except BaseException as Be:
    print Be.message