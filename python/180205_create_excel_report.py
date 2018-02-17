# -*- coding: utf-8 -*-
import sys, os, xlsxwriter,arcpy


#================================DEMO-VARIABLES====================================
#=========YAPI TYPE======================
yapi_tipi                   = u"-"
yapi_alttipi                = u"-"
#=========YAPI MAIN======================
ada                         = u"-"
parsel                      = u"-"
yapi_kodu                   = u"-"
koy_adi                     = u"-"
#=========REMAINING======================
ozgun_islev                 = u"-"
mevcut_islev                = u"-"
yapim_tarihi                = u"-"
avlu_ici_yerlesim           = u"-"
kat_sayisi                  = u"-"
yapim_teknigi               = u"-"
yapisal_durum               = u"-" 
degismislik                 = u"-"
ozgun_mimari_elemanlar      = u"-"

ozgun_avlu_elemanlari       = u"-"
ozgun_servis_birimleri      = u"-"
duvar_yapisal_durum         = u"-"
avlu_degismislik            = u"-"
duvar_deger_grubu           = u"-"
duvar_karar_grubu           = u"-"

cati_tipi                   = u"-"
cati_kaplama                = u"-"
cati_yapisal_durum          = u"-"
cati_degismislik            = u"-"

deger_grubu                 = u"-"
mudahale_onerisi            = u"-"
tescil_durumu               = u"-"
tescil_onerisi              = u"-"
karar_grubu                 = u"-"
#===========YENI_YAPI====================
doku_ile_uyum               = u"-"

#=============================================================================
#Variables, directory of the excel file and name of it
workingDrive = "E:\\"
reportFolder = workingDrive+"yapi_fisleri"
mapExortPath =  reportFolder+"\\"+"map_export"

ayvali_fis_folder = reportFolder+"\\AYVALI\\"
taskinpasa_fis_folder = reportFolder+"\\TASKINPASA\\"
cemil_fis_folder = reportFolder+"\\CEMIL\\"

workBookLocation_ayvali = ayvali_fis_folder
workBookLocation_cemil = cemil_fis_folder
workBookLocation_taskinpasa = taskinpasa_fis_folder

workBookName_geleneksel_ana = "01_Geleneksel_Ana_Yapi.xlsx"
workBookName_geleneksel_ek = "02_Geleneksel_Ek_Yapi.xlsx"
workBookName_yeni_ana = "03_Yeni_Ana_Yapi.xlsx"
workBookName_yeni_ek = "03_Yeni_Ana_Yapi.xlsx"

workbook_ayvali_geleneksel_ana = xlsxwriter.Workbook(workBookLocation_ayvali+workBookName_geleneksel_ana)
workbook_cemil_geleneksel_ana = xlsxwriter.Workbook(workBookLocation_cemil+workBookName_geleneksel_ana)
workbook_taskinpasa_geleneksel_ana = xlsxwriter.Workbook(workBookLocation_taskinpasa+workBookName_geleneksel_ana)

workbook_ayvali_geleneksel_ek = xlsxwriter.Workbook(workBookLocation_ayvali+workBookName_geleneksel_ek)
workbook_cemil_geleneksel_ek = xlsxwriter.Workbook(workBookLocation_cemil+workBookName_geleneksel_ek)
workbook_taskinpasa_geleneksel_ek = xlsxwriter.Workbook(workBookLocation_taskinpasa+workBookName_geleneksel_ek)

workbook_ayvali_yeni_ana = xlsxwriter.Workbook(workBookLocation_ayvali+workBookName_yeni_ana)
workbook_cemil_yeni_ana = xlsxwriter.Workbook(workBookLocation_cemil+workBookName_yeni_ana)
workbook_taskinpasa_yeni_ana = xlsxwriter.Workbook(workBookLocation_taskinpasa+workBookName_yeni_ana)

workbook_ayvali_yeni_ek = xlsxwriter.Workbook(workBookLocation_ayvali+workBookName_yeni_ek)
workbook_cemil_yeni_ek = xlsxwriter.Workbook(workBookLocation_cemil+workBookName_yeni_ek)
workbook_taskinpasa_yeni_ek = xlsxwriter.Workbook(workBookLocation_taskinpasa+workBookName_yeni_ek)

excel_path = workingDrive+"excele"
arcpy.env.workspace = excel_path
arcpy.env.overwriteOutput=True
mxd_path = os.path.join(excel_path,"excele_mxd.mxd")
yapi_data_gis = r"E:\excele\16.02.2018\yapidata-gis 09022018\nevsehir_gis.gdb\YAPI_DATA_GIS"

if not os.path.exists(mapExortPath):
    os.makedirs(mapExortPath)
if not os.path.exists(ayvali_fis_folder):
    os.makedirs(ayvali_fis_folder)
if not os.path.exists(taskinpasa_fis_folder):
    os.makedirs(taskinpasa_fis_folder)
if not os.path.exists(cemil_fis_folder):
    os.makedirs(cemil_fis_folder)

#yapi_objectid_den uretilen yeni kod kullanilacak
def export_map(map_document,layerName,yapi_id):
    try:
        imagePath = mapExortPath+"\\"+str(yapi_id)+".jpg"
        mxd = arcpy.mapping.MapDocument(map_document)
        df = arcpy.mapping.ListDataFrames(mxd, "AYVALI")[0]
        lyr = arcpy.mapping.ListLayers(mxd, layerName, df)[0]
        arcpy.SelectLayerByAttribute_management(lyr,"NEW_SELECTION","YAPI_KODU = "+str(yapi_id))
        df.zoomToSelectedFeatures()
        df.scale *= 1.5
        arcpy.RefreshActiveView()
        arcpy.mapping.ExportToJPEG(mxd,imagePath,df,df_export_width=363,df_export_height=300)
    except BaseException as be:
        print be.message

def create_geleneksel_ws(wbName, ada, parsel):
    try:
        worksheet = wbName.add_worksheet(ada.decode('utf-8')+"-"+parsel.decode('utf-8'))

        worksheet.set_portrait()
        worksheet.set_page_view()
        worksheet.set_paper(9)
        worksheet.set_margins(left=0.1,right=0.1)

        ##############################################################################
        #
        #set column widths
        #
        worksheet.set_column('A:A', 17.8)
        worksheet.set_column('B:B', 2.0)
        worksheet.set_column('C:C', 17.8)
        worksheet.set_column('D:D', 11.26)
        worksheet.set_column('E:E', 0.76)
        worksheet.set_column('F:F', 13)
        worksheet.set_column('G:G', 16.6)
        worksheet.set_column('H:H', 8.5)
        worksheet.set_row(1,8)
        worksheet.set_row(3,8)
        worksheet.set_row(5,8)
        worksheet.set_row(12,28)
        worksheet.set_row(24,26)
        #
        #
        #header1 format_h1#############################################################
        #
        #
        format_h1 = wbName.add_format()
        format_h1.set_bg_color('#808080')
        format_h1.set_font_color('white')
        format_h1.set_bold(True)
        format_h1.set_align('center')
        format_h1.set_font_name('Arial Narrow')
        format_h1.set_border(2)
        format_h1.set_font_size(10)
        #
        #
        #header2 format_h2#############################################################
        #
        #
        format_h2 = wbName.add_format()
        format_h2.set_bg_color('#d9d9d9')
        format_h2.set_font_color('black')
        format_h2.set_bold(True)
        format_h2.set_align('vcenter')
        format_h2.set_font_name('Arial Narrow')
        format_h2.set_border(1)
        format_h2.set_font_size(8.5)
        #
        #
        #header3 format_h3#############################################################
        #
        #
        format_h3 = wbName.add_format()
        format_h3.set_bg_color('#bfbfbf')
        format_h3.set_font_color('black')
        format_h3.set_bold(True)
        format_h3.set_align('center')
        format_h3.set_align('vcenter')
        format_h3.set_font_name('Arial Narrow')
        format_h3.set_border(1)
        format_h3.set_font_size(8.5)
        #
        #
        #datacell format_dc#############################################################
        #
        #
        format_dc = wbName.add_format()
        format_dc.set_font_color('black')
        format_dc.set_align('center')
        format_dc.set_font_name('Arial Narrow')
        format_dc.set_border(1)
        format_dc.set_font_size(8.5)
        #
        
        format_df = wbName.add_format()
        format_df.set_font_color('black')
        format_df.set_align('left')
        format_df.set_font_name('Arial Narrow')
        format_df.set_border(1)
        format_df.set_text_wrap()
        format_df.set_font_size(8.5)
        #h1_merge_format
        

        ##############################################################################

        worksheet.write('C1', u'ADA/PARSEL',format_h1)
        worksheet.write('A5', u'ÖZGÜN İŞLEV',format_h2)
        worksheet.write('F5', u'MEVCUT İŞLEV',format_h2)
        worksheet.write('H1', u'YAPI KODU',format_h1)
        worksheet.merge_range('A3:I3',u'GELENEKSEL ANA YAPI',format_h1)
        worksheet.merge_range('A7:D7',u'YAPI ÖZELLİKLERİ',format_h3)
        worksheet.merge_range('F7:I7',u'ÇATI ÖZELLİKLERİ',format_h3)
        worksheet.write('A8', u'YAPIM TARİHİ',format_h2)
        worksheet.write('A9', u'AVLU İÇİ YERLEŞİM',format_h2)
        worksheet.write('A10', u'KAT SAYISI',format_h2)
        worksheet.write('A11', u'YAPIM TEKNİĞİ',format_h2)
        worksheet.merge_range('A12:A13',u'YAPISAL DURUM',format_h2)
        worksheet.merge_range('A14:A15',u'DEĞİŞMİŞLİK',format_h2)
        worksheet.merge_range('A16:A17',u'ÖZGÜN MİMARİ ELEMANLAR',format_h2)
        worksheet.write('F8', u'TİPİ',format_h2)
        worksheet.write('F9', u'KAPLAMA',format_h2)
        worksheet.merge_range('F10:F11',u'YAPISAL DURUM',format_h2)
        worksheet.merge_range('F12:F13',u'DEĞİŞMİŞLİK',format_h2)
        worksheet.merge_range('F15:I15',u'DEĞERLENDİRME/MÜDAHALE',format_h3)

        worksheet.write('F16',u'DEĞER GRUBU',format_h2)
        worksheet.write('F17',u'MÜDAHALE ÖNERİSİ',format_h2)
        worksheet.write('F18', u'TESCİL DURUMU',format_h2)
        worksheet.write('F19', u'TESCİL ÖNERİSİ',format_h2)
        worksheet.write('F20', u'KARAR GRUBU',format_h2)

        worksheet.merge_range('A19:D19',u'AVLU ÖZELLİKLERİ/DEĞERLENDİRME',format_h3)
        worksheet.write('A20', u'ÖZGÜN AVLU ELEMANLARI',format_h2)
        worksheet.write('A21', u'ÖZGÜN SERVİS BİRİMLERİ',format_h2)
        worksheet.merge_range('A22:A23',u'DUVAR YAPISAL DURUMU',format_h2)
        worksheet.merge_range('A24:A25',u'DEĞİŞMİŞLİK',format_h2)
        worksheet.write('A26', u'DUVAR DEĞER GRUBU',format_h2)
        worksheet.write('A27', u'DUVAR KARAR GRUBU',format_h2)


        worksheet.write('A1', koy_adi.decode('utf-8'),format_h1)
        worksheet.write('I1',yapi_kodu,format_dc)
        worksheet.write('D1',ada+"-"+parsel,format_dc)
        worksheet.merge_range('B5:D5',ozgun_islev,format_df)
        worksheet.merge_range('G5:I5',mevcut_islev,format_df)
        worksheet.merge_range('B8:D8',yapim_tarihi,format_df)
        worksheet.merge_range('B9:D9',avlu_ici_yerlesim,format_df)
        worksheet.merge_range('B10:D10',kat_sayisi,format_df)
        worksheet.merge_range('B11:D11',yapim_teknigi,format_df)
        worksheet.merge_range('B12:D13',yapisal_durum,format_df)
        worksheet.merge_range('B14:D15',degismislik,format_df)
        worksheet.merge_range('B16:D17',ozgun_mimari_elemanlar,format_df)
        worksheet.merge_range('B20:D20',ozgun_avlu_elemanlari,format_df)
        worksheet.merge_range('B21:D21',ozgun_servis_birimleri,format_df)
        worksheet.merge_range('B22:D23',duvar_yapisal_durum,format_df)
        worksheet.merge_range('B24:D25',avlu_degismislik,format_df)
        worksheet.merge_range('B26:D26',duvar_deger_grubu,format_df)
        worksheet.merge_range('B27:D27',duvar_karar_grubu,format_df)
        worksheet.merge_range('G8:I8',cati_tipi,format_df)
        worksheet.merge_range('G9:I9',cati_kaplama,format_df)
        worksheet.merge_range('G10:I11',cati_yapisal_durum,format_df)
        worksheet.merge_range('G12:I13',cati_degismislik,format_df)

        worksheet.merge_range('G16:I16',deger_grubu,format_df)
        worksheet.merge_range('G17:I17',mudahale_onerisi,format_df)
        worksheet.merge_range('G18:I18',tescil_durumu,format_df)
        worksheet.merge_range('G19:I19',tescil_onerisi,format_df)
        worksheet.merge_range('G20:I20',karar_grubu,format_df)

        worksheet.insert_image('A29', mapExortPath+"\\"+str(yapi_kodu)+".jpg")
    except BaseException as Be:
        print Be.message

def create_ek_ws(wbName, ada, parsel):
    try:
        worksheet = wbName.add_worksheet(ada.decode('utf-8')+"-"+parsel.decode('utf-8'))

        worksheet.set_portrait()
        worksheet.set_page_view()
        worksheet.set_paper(9)
        worksheet.set_margins(left=0.1,right=0.1)

        ##############################################################################
        #
        #set column widths
        #
        worksheet.set_column('A:A', 17.8)
        worksheet.set_column('B:B', 2.0)
        worksheet.set_column('C:C', 17.8)
        worksheet.set_column('D:D', 11.26)
        worksheet.set_column('E:E', 0.76)
        worksheet.set_column('F:F', 13)
        worksheet.set_column('G:G', 16.6)
        worksheet.set_column('H:H', 8.5)
        worksheet.set_row(1,8)
        worksheet.set_row(3,8)
        worksheet.set_row(5,8)
        worksheet.set_row(12,28)
        #
        #
        #header1 format_h1#############################################################
        #
        #
        format_h1 = wbName.add_format()
        format_h1.set_bg_color('#808080')
        format_h1.set_font_color('white')
        format_h1.set_bold(True)
        format_h1.set_align('center')
        format_h1.set_font_name('Arial Narrow')
        format_h1.set_border(2)
        format_h1.set_font_size(10)
        #
        #
        #header2 format_h2#############################################################
        #
        #
        format_h2 = wbName.add_format()
        format_h2.set_bg_color('#d9d9d9')
        format_h2.set_font_color('black')
        format_h2.set_bold(True)
        format_h2.set_align('vcenter')
        format_h2.set_font_name('Arial Narrow')
        format_h2.set_border(1)
        format_h2.set_font_size(8.5)
        #
        #
        #header3 format_h3#############################################################
        #
        #
        format_h3 = wbName.add_format()
        format_h3.set_bg_color('#bfbfbf')
        format_h3.set_font_color('black')
        format_h3.set_bold(True)
        format_h3.set_align('center')
        format_h3.set_align('vcenter')
        format_h3.set_font_name('Arial Narrow')
        format_h3.set_border(1)
        format_h3.set_font_size(8.5)
        #
        #
        #datacell format_dc#############################################################
        #
        #
        format_dc = wbName.add_format()
        format_dc.set_font_color('black')
        format_dc.set_align('center')
        format_dc.set_font_name('Arial Narrow')
        format_dc.set_border(1)
        format_dc.set_font_size(8.5)
        #
        
        format_df = wbName.add_format()
        format_df.set_font_color('black')
        format_df.set_align('left')
        format_df.set_font_name('Arial Narrow')
        format_df.set_border(1)
        format_df.set_text_wrap()
        format_df.set_font_size(8.5)
        #h1_merge_format
        

        ##############################################################################

        worksheet.write('C1', u'ADA/PARSEL',format_h1)
        worksheet.write('A5', u'ÖZGÜN İŞLEV',format_h2)
        worksheet.write('F5', u'MEVCUT İŞLEV',format_h2)
        worksheet.write('H1', u'YAPI KODU',format_h1)
        worksheet.merge_range('A3:I3',u'GELENEKSEL EK YAPI',format_h1)
        worksheet.merge_range('A7:D7',u'YAPI ÖZELLİKLERİ',format_h3)
        worksheet.merge_range('F7:I7',u'ÇATI ÖZELLİKLERİ',format_h3)
        worksheet.write('A8', u'YAPIM TARİHİ',format_h2)
        worksheet.write('A9', u'AVLU İÇİ YERLEŞİM',format_h2)
        worksheet.write('A10', u'KAT SAYISI',format_h2)
        worksheet.write('A11', u'YAPIM TEKNİĞİ',format_h2)
        worksheet.merge_range('A12:A13',u'YAPISAL DURUM',format_h2)
        worksheet.merge_range('A14:A15',u'DEĞİŞMİŞLİK',format_h2)
        worksheet.merge_range('A16:A17',u'ÖZGÜN MİMARİ ELEMANLAR',format_h2)
        worksheet.write('F8', u'TİPİ',format_h2)
        worksheet.write('F9', u'KAPLAMA',format_h2)
        worksheet.merge_range('F10:F11',u'YAPISAL DURUM',format_h2)
        worksheet.merge_range('F12:F13',u'DEĞİŞMİŞLİK',format_h2)
        worksheet.merge_range('A19:D19',u'DEĞERLENDİRME/MÜDAHALE',format_h3)

        worksheet.write('A20',u'DEĞER GRUBU',format_h2)
        worksheet.write('A21',u'MÜDAHALE ÖNERİSİ',format_h2)
        worksheet.write('A22', u'TESCİL DURUMU',format_h2)
        worksheet.write('A23', u'TESCİL ÖNERİSİ',format_h2)
        worksheet.write('A24', u'KARAR GRUBU',format_h2)



        worksheet.write('A1', koy_adi.decode('utf-8'),format_h1)
        worksheet.write('I1',yapi_kodu,format_dc)
        worksheet.write('D1',ada+"-"+parsel,format_dc)
        worksheet.merge_range('B5:D5',ozgun_islev,format_df)
        worksheet.merge_range('G5:I5',mevcut_islev,format_df)
        worksheet.merge_range('B8:D8',yapim_tarihi,format_df)
        worksheet.merge_range('B9:D9',avlu_ici_yerlesim,format_df)
        worksheet.merge_range('B10:D10',kat_sayisi,format_df)
        worksheet.merge_range('B11:D11',yapim_teknigi,format_df)
        worksheet.merge_range('B12:D13',yapisal_durum,format_df)
        worksheet.merge_range('B14:D15',degismislik,format_df)
        worksheet.merge_range('B16:D17',ozgun_mimari_elemanlar,format_df)
       
        worksheet.merge_range('G8:I8',cati_tipi,format_df)
        worksheet.merge_range('G9:I9',cati_kaplama,format_df)
        worksheet.merge_range('G10:I11',cati_yapisal_durum,format_df)
        worksheet.merge_range('G12:I13',cati_degismislik,format_df)

        worksheet.merge_range('B20:D20',deger_grubu,format_df)
        worksheet.merge_range('B21:D21',mudahale_onerisi,format_df)
        worksheet.merge_range('B22:D22',tescil_durumu,format_df)
        worksheet.merge_range('B23:D23',tescil_onerisi,format_df)
        worksheet.merge_range('B24:D24',karar_grubu,format_df)

        worksheet.insert_image('A26', mapExortPath+"\\"+str(yapi_kodu)+".jpg")
    except BaseException as Be:
        print Be.message

def create_yeni_ana_ws(wbName, ada, parsel):
    try:
        worksheet = wbName.add_worksheet(ada.decode('utf-8')+"-"+parsel.decode('utf-8'))

        worksheet.set_portrait()
        worksheet.set_page_view()
        worksheet.set_paper(9)
        worksheet.set_margins(left=0.1,right=0.1)

        ##############################################################################
        #
        #set column widths
        #
        worksheet.set_column('A:A', 17.8)
        worksheet.set_column('B:B', 2.0)
        worksheet.set_column('C:C', 17.8)
        worksheet.set_column('D:D', 11.26)
        worksheet.set_column('E:E', 0.76)
        worksheet.set_column('F:F', 13)
        worksheet.set_column('G:G', 16.6)
        worksheet.set_column('H:H', 8.5)
        worksheet.set_row(1,8)
        worksheet.set_row(3,8)
        worksheet.set_row(5,8)
        worksheet.set_row(12,28)
        #
        #
        #header1 format_h1#############################################################
        #
        #
        format_h1 = wbName.add_format()
        format_h1.set_bg_color('#808080')
        format_h1.set_font_color('white')
        format_h1.set_bold(True)
        format_h1.set_align('center')
        format_h1.set_font_name('Arial Narrow')
        format_h1.set_border(2)
        format_h1.set_font_size(10)
        #
        #
        #header2 format_h2#############################################################
        #
        #
        format_h2 = wbName.add_format()
        format_h2.set_bg_color('#d9d9d9')
        format_h2.set_font_color('black')
        format_h2.set_bold(True)
        format_h2.set_align('vcenter')
        format_h2.set_font_name('Arial Narrow')
        format_h2.set_border(1)
        format_h2.set_font_size(8.5)
        #
        #
        #header3 format_h3#############################################################
        #
        #
        format_h3 = wbName.add_format()
        format_h3.set_bg_color('#bfbfbf')
        format_h3.set_font_color('black')
        format_h3.set_bold(True)
        format_h3.set_align('center')
        format_h3.set_align('vcenter')
        format_h3.set_font_name('Arial Narrow')
        format_h3.set_border(1)
        format_h3.set_font_size(8.5)
        #
        #
        #datacell format_dc#############################################################
        #
        #
        format_dc = wbName.add_format()
        format_dc.set_font_color('black')
        format_dc.set_align('center')
        format_dc.set_font_name('Arial Narrow')
        format_dc.set_border(1)
        format_dc.set_font_size(8.5)
        #
        
        format_df = wbName.add_format()
        format_df.set_font_color('black')
        format_df.set_align('left')
        format_df.set_font_name('Arial Narrow')
        format_df.set_border(1)
        format_df.set_text_wrap()
        format_df.set_font_size(8.5)
        #h1_merge_format
        

        ##############################################################################

        worksheet.write('C1', u'ADA/PARSEL',format_h1)
        worksheet.write('A5', u'MEVCUT İŞLEV',format_h2)
        worksheet.write('H1', u'YAPI KODU',format_h1)
        worksheet.merge_range('A3:I3',u'YENİ ANA YAPI',format_h1)
        worksheet.merge_range('A8:D8',u'YAPI ÖZELLİKLERİ',format_h3)
        worksheet.merge_range('F5:I5',u'ÇATI ÖZELLİKLERİ',format_h3)
        worksheet.write('A9', u'YAPIM TARİHİ',format_h2)
        worksheet.write('A10', u'AVLU İÇİ YERLEŞİM',format_h2)
        worksheet.write('A11', u'KAT SAYISI',format_h2)
        worksheet.write('A12', u'YAPIM TEKNİĞİ',format_h2)
        worksheet.merge_range('A13:A14',u'YAPISAL DURUM',format_h2)



        worksheet.write('F6', u'TİPİ',format_h2)
        worksheet.write('F7', u'KAPLAMA',format_h2)

        worksheet.merge_range('F9:I9',u'DEĞERLENDİRME/MÜDAHALE',format_h3)
        worksheet.merge_range('F10:I11',u'DOKU İLE UYUM',format_h2)
        
        worksheet.write('F12',u'DEĞER GRUBU',format_h2)
        worksheet.write('F13',u'MÜDAHALE ÖNERİSİ',format_h2)
        worksheet.write('F14', u'KARAR GRUBU',format_h2)



        worksheet.write('A1', koy_adi.decode('utf-8'),format_h1)
        worksheet.write('I1',yapi_kodu,format_dc)
        worksheet.write('D1',ada+"-"+parsel,format_dc)

        worksheet.merge_range('B5:D5',mevcut_islev,format_df)
        worksheet.merge_range('B9:D9',yapim_tarihi,format_df)
        worksheet.merge_range('B10:D10',avlu_ici_yerlesim,format_df)
        worksheet.merge_range('B11:D11',kat_sayisi,format_df)
        worksheet.merge_range('B12:D12',yapim_teknigi,format_df)
        worksheet.merge_range('B13:D14',yapisal_durum,format_df)

       
        worksheet.merge_range('G6:I6',cati_tipi,format_df)
        worksheet.merge_range('G7:I7',cati_kaplama,format_df)
        

        worksheet.merge_range('G10:I11',doku_ile_uyum,format_df)
        worksheet.merge_range('G12:I12',deger_grubu,format_df)
        worksheet.merge_range('G13:I13',mudahale_onerisi,format_df)

        worksheet.merge_range('G14:I14',karar_grubu,format_df)
    except BaseException as Be:
        print Be.message
#=============================================================================
#Workbook Formatting...


cursor = arcpy.SearchCursor(yapi_data_gis,"Y_YAPI_ID not in (0,-4,-9)")

for row in cursor:
    yapi_kodu = int(row.getValue("YAPI_KODU"))
    yapi_tipi = row.getValue("Y_YAPITIPI")
    yapi_alttipi = row.getValue("Y_YAPI_ALT")
    ada = row.getValue("M_ADA")
    parsel = row.getValue("M_PARSEL")
    koy_adi = row.getValue("M_KOYNAME")
    ozgun_islev = row.getValue("Y_YAPI_OZG")
    mevcut_islev = row.getValue("Y_YAPI_MEV")
    avlu_ici_yerlesim = row.getValue("Y_YAPI_AVL")
    kat_sayisi = row.getValue("Y_KAT_SAYI")
    yapim_teknigi = row.getValue("Y_YAPIM_TE")
    yapisal_durum = row.getValue("Y_YAPISAL_")
    degismislik = row.getValue("Y_YAPI_DEG")
    ozgun_mimari_elemanlar = row.getValue("Y_OZGUN_MI")
    ozgun_avlu_elemanlari = row.getValue("Y_OZGUN_AV")
    ozgun_servis_birimleri = row.getValue("Y_OZGUN_SE")
    duvar_yapisal_durum = row.getValue("Y_AVLU__01")
    avlu_degismislik = row.getValue("Y_AVLU__03")

    avlu_karar_deger_temp = row.getValue("Y_AVLU__04")
    # avlu duvar deger/karar grubu
    if avlu_karar_deger_temp is not None:
        if(avlu_karar_deger_temp.split(' ')[0] == '1'):
            duvar_deger_grubu = u"1-Nitelikli"
            duvar_karar_grubu = u"1-Korunacak"
        elif(avlu_karar_deger_temp.split(' ')[0] == '2'):
            duvar_deger_grubu = u"2-Niteliksiz"
            duvar_karar_grubu = u"2-Kaldırılacak"
        elif(avlu_karar_deger_temp.split(' ')[0] == '3'):
            duvar_deger_grubu = u"3-Az Nitelikli"
            duvar_karar_grubu = u"3-Onarılaca"
        else:
            duvar_deger_grubu = avlu_karar_deger_temp
    else:
        duvar_deger_grubu = "-"

    cati_tipi = row.getValue("Y_CATI_TIP")
    cati_kaplama = row.getValue("Y_CATI_KAP")
    cati_yapisal_durum = row.getValue("Y_CATI_YAP")
    cati_degismislik = row.getValue("Y_CATI_DEG")

    deger_grubu = row.getValue("DEGER_GRUBU")
    mudahale_onerisi = row.getValue("Y_YAPI_YER")
    tescil_durumu = row.getValue("TESCİL_DURUMU")
    tescil_onerisi = row.getValue("TESCIL_ONERISI")
    karar_grubu = row.getValue("KARAR_GRUBU")

    doku_ile_uyum = row.getValue("Y_YENI_YAP")

    if(yapi_tipi == u"Geleneksel" and yapi_alttipi == u"Ana Yapı" and koy_adi == u"Cemil"):
        export_map(mxd_path,"YAPI_DATA_GIS",yapi_kodu)    
        create_geleneksel_ws(workbook_cemil_geleneksel_ana,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    elif (yapi_tipi == u"Geleneksel" and yapi_alttipi == u"Ana Yapı" and koy_adi == u"Ayvali"):
        export_map(mxd_path,"YAPI_DATA_GIS",yapi_kodu)    
        create_geleneksel_ws(workbook_ayvali_geleneksel_ana,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    elif (yapi_tipi == u"Geleneksel" and yapi_alttipi == u"Ana Yapı" and koy_adi == u"Taskinpasa"):
        export_map(mxd_path,"YAPI_DATA_GIS",yapi_kodu)    
        create_geleneksel_ws(workbook_taskinpasa_geleneksel_ana,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel

    elif (yapi_tipi == u"Geleneksel" and yapi_alttipi == u"Ek Yapı" and koy_adi == u"Cemil"):
        export_map(mxd_path,"YAPI_DATA_GIS",yapi_kodu)    
        create_ek_ws(workbook_cemil_geleneksel_ek,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    elif (yapi_tipi == u"Geleneksel" and yapi_alttipi == u"Ek Yapı" and koy_adi == u"Ayvali"):
        export_map(mxd_path,"YAPI_DATA_GIS",yapi_kodu)    
        create_ek_ws(workbook_ayvali_geleneksel_ek,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    elif (yapi_tipi == u"Geleneksel" and yapi_alttipi == u"Ek Yapı" and koy_adi == u"Taskinpasa"):
        export_map(mxd_path,"YAPI_DATA_GIS",yapi_kodu)    
        create_ek_ws(workbook_taskinpasa_geleneksel_ek,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    
    elif (yapi_tipi == u"Yeni Yapı" and yapi_alttipi == u"Ana Yapı" and koy_adi == u"Cemil"):
        create_yeni_ana_ws(workbook_cemil_yeni_ana,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    elif (yapi_tipi == u"Yeni Yapı" and yapi_alttipi == u"Ana Yapı" and koy_adi == u"Ayvali"):
        create_yeni_ana_ws(workbook_ayvali_yeni_ana,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    elif (yapi_tipi == u"Yeni Yapı" and yapi_alttipi == u"Ana Yapı" and koy_adi == u"Taskinpasa"):
        create_yeni_ana_ws(workbook_taskinpasa_yeni_ana,ada,parsel)
        print yapi_kodu,"-",koy_adi,"-",ada,"-",parsel
    else:
        continue

workbook_ayvali_geleneksel_ana.close()
workbook_cemil_geleneksel_ana.close()
workbook_taskinpasa_geleneksel_ana.close()
workbook_ayvali_geleneksel_ek.close()
workbook_cemil_geleneksel_ek.close()
workbook_taskinpasa_geleneksel_ek.close()
workbook_ayvali_yeni_ana.close()
workbook_cemil_yeni_ana.close()
workbook_taskinpasa_yeni_ana.close()
workbook_ayvali_yeni_ek.close()
workbook_cemil_yeni_ek.close()
workbook_taskinpasa_yeni_ek.close()