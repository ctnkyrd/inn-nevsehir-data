# -*- coding: utf-8 -*-
import sys, os, xlsxwriter

workBookLocation = "E:\\"
workBookName = "demo.xlsx"

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
##############################################################################



worksheet.write('A1', u'[Ayvalı]',format_h1)
worksheet.write('C1', u'ADA/PARSEL',format_h1)
worksheet.write('A5', u'ÖZGÜN İŞLEV',format_h2)

workbook.close()