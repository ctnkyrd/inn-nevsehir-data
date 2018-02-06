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

worksheet.set_column('A:A', 17.8)
worksheet.set_column('B:B', 2.0)
worksheet.set_column('C:C', 17.8)
worksheet.set_column('D:D', 11.26)
worksheet.set_column('E:E', 0.76)
worksheet.set_column('F:F', 11.3)
worksheet.set_column('G:G', 16.6)



format = workbook.add_format()
format.set_bg_color('#808080')
format.set_font_color('white')
format.set_bold(True)
format.set_align('center')
format.set_font_name('Arial Narrow')
format.set_border(2)
format.set_font_size(10)



worksheet.write('A1', u'[Ayvalı]',format)
worksheet.write('C1', u'ADA/PARSEL',format)


workbook.close()