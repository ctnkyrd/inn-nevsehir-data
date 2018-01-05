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
    print "Unable to connect to the database"
    print Be.message
#ConnString=================================================================================

#Dictionary=================================================================================
global tableNoDict
tableNoDict = {}
tableNoDict[1] = 0
tableNoDict[2] = 0
tableNoDict[3] = 0
tableNoDict[4] = 0
tableNoDict[5] = 0
tableNoDict[6] = 0
tableNoDict[7] = 0
tableNoDict[8] = 0
tableNoDict[9] = 1
tableNoDict[10] = 1
tableNoDict[11] = 0
tableNoDict[12] = 1
tableNoDict[13] = 1
tableNoDict[14] = 0
tableNoDict[15] = 1
tableNoDict[16] = 0
tableNoDict[17] = 0
tableNoDict[18] = 1
tableNoDict[19] = 0
tableNoDict[20] = 0
tableNoDict[21] = 0
tableNoDict[22] = 0
tableNoDict[23] = 0
tableNoDict[24] = 0
tableNoDict[25] = 0
tableNoDict[26] = 0
tableNoDict[27] = 0
tableNoDict[28] = 1
tableNoDict[29] = 1
tableNoDict[30] = 1
tableNoDict[31] = 1
tableNoDict[32] = 1
tableNoDict[33] = 1
tableNoDict[34] = 1
tableNoDict[35] = 1
tableNoDict[36] = 1
tableNoDict[37] = 1
tableNoDict[38] = 0
tableNoDict[39] = 0
tableNoDict[40] = 1
tableNoDict[41] = 1
tableNoDict[42] = 1
tableNoDict[43] = 0
tableNoDict[44] = 0
tableNoDict[45] = 0
tableNoDict[46] = 0
#Dictionary=================================================================================



#TheCode====================================================================================
try:
    cur = conn.cursor()
    cur.execute("""SELECT * from yapidata_son""")
    rows = cur.fetchall()
    rowCounter = 0
    for row in rows:
        print row[0]
        rowCounter += 1
except BaseException as Be:
    print Be.message