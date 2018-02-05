# -*- coding: utf-8 -*-
import codecs
import sys, os, xlsxwriter

workBookLocation = "E:\\"
workBookName = "demo.xlsx"

workbook = xlsxwriter.Workbook(workBookLocation+workBookName)
worksheet = workbook.add_worksheet()

worksheet.set_portrait()
worksheet.set_page_view()
worksheet.set_paper(9)

worksheet.set_column('A:A', 20)
bold = workbook.add_format({'bold': True})

format = workbook.add_format()
format.set_bg_color('#808080')
format.set_font_color('white')
format.set_bold(True)


worksheet.write('A1', u'[Ayvalı]',format)


workbook.close()