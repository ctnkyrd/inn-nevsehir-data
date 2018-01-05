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

global tableNameDict
tableNameDict = {}
tableNameDict[1] = "id"
tableNameDict[2] = "projectid"
tableNameDict[3] = "groupname"
tableNameDict[4] = "koyname"
tableNameDict[5] = "ada"
tableNameDict[6] = "parsel"
tableNameDict[7] = "sokak"
tableNameDict[8] = "kapino"
tableNameDict[9] = "yapitipi"
tableNameDict[10] = "yapi_alttipi"
tableNameDict[11] = "yaklasik_yapim_tarihi"
tableNameDict[12] = "yapi_ozgun_islev"
tableNameDict[13] = "yapi_mevcut_islev"
tableNameDict[14] = "yapi_avlu_ici_yerlesim"
tableNameDict[15] = "avlu_duvari_yapim_teknigi"
tableNameDict[16] = "avlu_bahce_duvari_yukseklik"
tableNameDict[17] = "avlu_bahce_duvari_yapisal_durum"
tableNameDict[18] = "avlu_bahce_zemini"
tableNameDict[19] = "avlu_bahce_degismisligi"
tableNameDict[20] = "avlu_duvari_mudahale_onerisi"
tableNameDict[21] = "cati_yapisal_durum"
tableNameDict[22] = "cati_degismisligi"
tableNameDict[23] = "yapisal_durum"
tableNameDict[24] = "yapi_degismisligi"
tableNameDict[25] = "yapi_yerinde_mudahale_onerisi"
tableNameDict[26] = "cevresel_olcekte_yapi_degeri"
tableNameDict[27] = "yapi_fotografi"
tableNameDict[28] = "ozgun_avlu_elemanlari"
tableNameDict[29] = "ozgun_servis_birimleri"
tableNameDict[30] = "cati_kaplama_malzemesi"
tableNameDict[31] = "yapim_teknigi"
tableNameDict[32] = "ozgun_mimari_elemanlar"
tableNameDict[33] = "duvar_malzemesi"
tableNameDict[34] = "tavan_malzemesi"
tableNameDict[35] = "doseme_malzemesi"
tableNameDict[36] = "ozgun_mimari_elemanlar_ic"
tableNameDict[37] = "cati_tipi_formu"
tableNameDict[38] = "kat_sayisi"
tableNameDict[39] = "tesisat"
tableNameDict[40] = "doku_ile_uyum"
tableNameDict[41] = "mimari_deger"
tableNameDict[42] = "yeni_yapilara_yonelik_mudahale_onerileri"
tableNameDict[43] = "status"
tableNameDict[44] = "timestamp"
tableNameDict[45] = "createdAt"
tableNameDict[46] = "updatedAt"
#Dictionary=================================================================================

def row_processes(row):
    cur2 = conn.cursor()
    theSentence = []
    for i in range(46):
        theRow = []
        if(tableNoDict[i] == 0):
            theRow.append(row[i])
        else:
            rowLength = row[i].split(',')
            for j in range(rowLength):
                cur2.execute = ("select deger from kodyapidata where column_name = '"+tableNameDict[i]+"' and kod = "+str(row[i].split(',')[j])
                theRow.append()
            pass
            # cur2.execute("select deger from kodyapidata where column_name='")
            

<<<<<<< HEAD
#TheCode====================================================================================
try:
    cur = conn.cursor()
    cur.execute("""SELECT * from yapidata_son""")
    rows = cur.fetchall()
    for row in rows:
        row_processes(row)
=======
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

>>>>>>> eac69a108ce40666f5a289b62c1f92bc766c1f73
except BaseException as Be:
    print Be.message