# -*- coding: utf-8 -*-

import sys, os
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
def is_number(s):
    if s == None:
        return False
    else:
        try:
            float(s)
            return True
        except ValueError:
            pass
    
        try:
            import unicodedata
            unicodedata.numeric(s)
            return True
        except (TypeError, ValueError):
            pass


def updateDecoded(column_name, value,row_id):
    cur3 = conn.cursor()
    cur3.execute("UPDATE yapidata_decode SET "+column_name+" =%s WHERE id = %s", (value,row_id))
    conn.commit()
    cur3.close()



def row_processes(row):
    try:
        cur2 = conn.cursor()
        theSentence = ""
        rowId = row[0]
        
        print rowId, "Started!"
        for i in range(1,47):
            x=i-1
            if(tableNoDict[i] == 0):
                continue
            else:
                if is_number(row[x]):
                    multiColumnWrite = ""
                    kod = str(row[x])
                    cur2.execute ("select deger from kodyapidata where column_name = '"+tableNameDict[i]+"' and kod = "+kod)
                    a = cur2.fetchone()
                    multiColumnWrite = a[0]
                    updateDecoded(tableNameDict[i], multiColumnWrite, rowId)
                elif row[x] is not None:
                    rowLength = len(row[x].split(','))
                    multiColumnWrite = ""
                    for j in range(rowLength):
                        kod = str(row[x].split(',')[j].encode('utf-8'))
                        if kod.isdigit():
                            cur2.execute ("select deger from kodyapidata where column_name = '"+tableNameDict[i]+"' and kod = "+kod)
                            a = cur2.fetchone()
                            if(len(multiColumnWrite) == 0):
                                multiColumnWrite = a[0]
                            else:
                                multiColumnWrite = multiColumnWrite + ","+a[0]
                        else:
                            b = str(row[x].encode('utf-8')).split(',')[j]
                            if (len(b)!=0):
                                if(len(multiColumnWrite) == 0):
                                    multiColumnWrite = b
                                else:
                                    multiColumnWrite = multiColumnWrite + ","+b.decode('utf-8')
                    updateDecoded(tableNameDict[i], multiColumnWrite, rowId)                  
                else:
                    multiColumnWrite = ""
                    updateDecoded(tableNameDict[i], multiColumnWrite, rowId)
        print rowId, "Completed!"
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno, rowId, tableNameDict[i])            


#TheCode====================================================================================
try:
    cur = conn.cursor()
    cur.execute("""SELECT * from yapidata_son""")
    rows = cur.fetchall()
    for row in rows:
        row_processes(row)
except Exception as e:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)

print "Done!"
