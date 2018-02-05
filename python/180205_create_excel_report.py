# -*- coding: utf-8 -*-

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
worksheet.write('A1', 'Hello')
worksheet.write('A2', 'World', bold)
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)
workbook.close()