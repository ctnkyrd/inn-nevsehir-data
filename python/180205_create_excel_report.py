# -*- coding: utf-8 -*-
import sys, os, xlsxwriter

#================================VARIABLES====================================

#=========YAPI TYPE======================
yapi_tipi                   = "Geleneksel"
yapi_alttipi                = "Ana Yapı"
#=========YAPI MAIN======================
ada_parsel                  = "32/34v21a"
yapi_kodu                   = "1232"
koy_adi                     = "Ayvalı"
#=========REMAINING======================
ozgun_islev                 = "Konut,Ticaret"
mevcut_islev                = "Depo,Ahır"
yapim_tarihi                = "19.yy"
avlu_ici_yerlesim           = "Avlu/Bahçe İçinde – Bitişik"
kat_sayisi                  = "Z+Seçim Yapılmadı"
yapim_teknigi               = "Yığma Kesme Taş,Betonarme,Yığma Tuğla Briket"
yapisal_durum               = "4 HARABE- Yapının bazı mekânında ya da tamamında çökme var" 
degismislik                 = "2 Cephe ve kütle organizasyonu okunabiliyor, açıklıkların form, boyut ve sayılarında, malzemelerde değişme var."
ozgun_mimari_elemanlar      = "Parmaklık,Saçak,Çörten,Kapı Silmesi,tepe penceresi"

ozgun_avlu_elemanlari       = "Tespit Edilemedi"
ozgun_servis_birimleri      = "Ahır"
duvar_yapisal_durum         = "HARABE- Yapının bazı mekânında ya da tamamında çökme var"
avlu_degismislik            = "2 Konum ve boyut tamamen korunmuş, elemanlar, malzeme ve formda ciddi değişiklikler var."
duvar_deger_grubu           = "1 Nitelikli, olduğu gibi korunacak avlu duvarı."
duvar_karar_grubu           = "1 Nitelikli, olduğu gibi korunacak avlu duvarı."

cati_tipi                   = "Düz,Teras,birkismi 1 katli avlu duvari mekan yapilmis arkada"
cati_kaplama                = "Oluklu Sac,Şap"
cati_yapisal_durum          = "2 ORTA-Basit onarım ve bakıma ihtiyacı var"
cati_degismislik            = "3 Konum, boyut ya da form değişmiş, çatı sistemi, malzemeleri, kaplaması kısmen ya da tamamen değiştirilmiş, özgün çatı okunamıyor"

deger_grubu                 = "2 Cephe ve kütle organizasyonu okunabiliyor, açıklıkların form, boyut ve sayılarında, malzemelerde değişme var."
mudahale_onerisi            = "4 Kütle ve malzeme özellikleri ile dokuya uyumlu, diğer özellikleri doku ile uyumlu hale getirilerek korunacak yapı"
tescil_durumu               = ""
tescil_onerisi              = "Tescile Önerilen"
karar_grubu                 = "2 Cephe ve kütle organizasyonu okunabiliyor, açıklıkların form, boyut ve sayılarında, malzemelerde değişme var."
#=============================================================================
#Variables, directory of the excel file and name of it
workBookLocation = "E:\\"
workBookName = "demo.xlsx"

#=============================================================================
#Workbook Formatting...
try:
    workbook = xlsxwriter.Workbook(workBookLocation+workBookName)
    worksheet = workbook.add_worksheet()

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
    worksheet.set_column('F:F', 11.3)
    worksheet.set_column('G:G', 16.6)
    worksheet.set_column('H:H', 8.5)
    worksheet.set_row(1,8)
    worksheet.set_row(3,8)
    worksheet.set_row(6,8)
    worksheet.set_row(25,8)
    #
    #
    #header1 format_h1#############################################################
    #
    #
    format_h1 = workbook.add_format()
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
    format_h2 = workbook.add_format()
    format_h2.set_bg_color('#d9d9d9')
    format_h2.set_font_color('black')
    format_h2.set_bold(True)
    format_h2.set_font_name('Arial Narrow')
    format_h2.set_border(1)
    format_h2.set_font_size(8.5)
    #
    #
    #header3 format_h3#############################################################
    #
    #
    format_h3 = workbook.add_format()
    format_h3.set_bg_color('#bfbfbf')
    format_h3.set_font_color('black')
    format_h3.set_bold(True)
    format_h3.set_align('center')
    format_h3.set_font_name('Arial Narrow')
    format_h3.set_border(1)
    format_h3.set_font_size(8.5)
    #
    #
    #datacell format_dc#############################################################
    #
    #
    format_dc = workbook.add_format()
    format_dc.set_font_color('black')
    format_dc.set_align('center')
    format_dc.set_font_name('Arial Narrow')
    format_dc.set_border(1)
    format_dc.set_font_size(8.5)
    #
    #
    #h1_merge_format
    

    ##############################################################################
except BaseException as Be:
    print Be.message


worksheet.write('A1', u'[Ayvalı]',format_h1)
worksheet.write('C1', u'ADA/PARSEL',format_h1)
worksheet.write('A5', u'ÖZGÜN İŞLEV',format_h2)
worksheet.merge_range('A3:I3','GELENEKSEL ANA YAPI',format_h1)
worksheet.merge_range('B5:D5',u'KONUT',format_dc)
workbook.close()