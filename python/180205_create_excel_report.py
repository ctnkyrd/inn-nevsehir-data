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
    format_h2.set_align('vcenter')
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
    format_h3.set_align('vcenter')
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
    
    format_df = workbook.add_format()
    format_df.set_font_color('black')
    format_df.set_align('left')
    format_df.set_font_name('Arial Narrow')
    format_df.set_border(1)
    format_df.set_text_wrap()
    format_df.set_font_size(8.5)
    #h1_merge_format
    

    ##############################################################################
except BaseException as Be:
    print Be.message


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
worksheet.write('F16', u'DEĞER GRUBU',format_h2)
worksheet.write('F17', u'MÜDAHALE ÖNERİSİ',format_h2)
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
worksheet.write('D1',ada_parsel.decode('utf-8'),format_dc)
worksheet.merge_range('B5:D5',ozgun_islev.decode('utf-8'),format_df)
worksheet.merge_range('G5:I5',mevcut_islev.decode('utf-8'),format_df)
worksheet.merge_range('B8:D8',yapim_tarihi.decode('utf-8'),format_df)
worksheet.merge_range('B9:D9',avlu_ici_yerlesim.decode('utf-8'),format_df)
worksheet.merge_range('B10:D10',kat_sayisi.decode('utf-8'),format_df)
worksheet.merge_range('B11:D11',yapim_teknigi.decode('utf-8'),format_df)
worksheet.merge_range('B12:D13',yapisal_durum.decode('utf-8'),format_df)
worksheet.merge_range('B14:D15',degismislik.decode('utf-8'),format_df)
worksheet.merge_range('B16:D17',ozgun_mimari_elemanlar.decode('utf-8'),format_df)
worksheet.merge_range('B20:D20',ozgun_avlu_elemanlari.decode('utf-8'),format_df)
worksheet.merge_range('B21:D21',ozgun_servis_birimleri.decode('utf-8'),format_df)
worksheet.merge_range('B22:D23',duvar_yapisal_durum.decode('utf-8'),format_df)
worksheet.merge_range('B24:D25',avlu_degismislik.decode('utf-8'),format_df)
worksheet.merge_range('B26:D26',duvar_deger_grubu.decode('utf-8'),format_df)
worksheet.merge_range('B27:D27',duvar_karar_grubu.decode('utf-8'),format_df)
worksheet.merge_range('G8:I8',cati_tipi.decode('utf-8'),format_df)
worksheet.merge_range('G9:I9',cati_kaplama.decode('utf-8'),format_df)
worksheet.merge_range('G10:I11',cati_yapisal_durum.decode('utf-8'),format_df)
worksheet.merge_range('G12:I13',cati_degismislik.decode('utf-8'),format_df)
worksheet.merge_range('G16:I16',deger_grubu.decode('utf-8'),format_df)
worksheet.merge_range('G17:I17',mudahale_onerisi.decode('utf-8'),format_df)
worksheet.merge_range('G18:I18',tescil_durumu.decode('utf-8'),format_df)
worksheet.merge_range('G19:I19',tescil_onerisi.decode('utf-8'),format_df)
worksheet.merge_range('G20:I20',karar_grubu.decode('utf-8'),format_df)




















workbook.close()