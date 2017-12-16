# -*- coding: utf-8 -*-


import psycopg2
import psycopg2.extensions
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


#ConnString=================================================================================
try:
    conn = psycopg2.connect("dbname="+dbName+" user="+user+" host="+host+" password="+password)
    conn.set_client_encoding('UTF-8')
    print "Connected Succesfully!"
except BaseException as Be:
    print "I am unable to connect to the database"
    print Be.message
#ConnString=================================================================================



#TheCode====================================================================================
try:
    cur = conn.cursor()
    cur.execute("""SELECT yapi_avlu_ici_yerlesim from yapidata_son""")
    rows = cur.fetchall()
    for row in rows:
        print row[0]
except BaseException as Be:
    print Be.message