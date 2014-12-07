#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from datetime import date
from xlwt import *
fnt = Font()
fnt.name = 'Arial'
'''
borders = Borders()
borders.left = Borders.THICK
borders.right = Borders.THICK
borders.top = Borders.THICK
borders.bottom = Borders.THICK

pattern = Pattern()
pattern.pattern = Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = 0x21

style = XFStyle()
style.num_format_str='YYYY-MM-DD'
style.font = fnt
style.borders = borders
style.pattern = pattern
'''
def main_table(sheet):
    pattern = Pattern()
    pattern.pattern = Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 0x09
    #pattern.easyxf("align: horiz center")
    first_col = sheet.col(0)
    first_col.width = 256 * 6
    for i in xrange(1, 12):
        first_col = sheet.col(i)
        first_col.width = 256 * 15
        if i <= 6:
            first_row = sheet.row(i)
            first_row.height = 256 * 2

    alignment = Alignment()
    alignment.horz = Alignment.HORZ_CENTER


    style = XFStyle()
    style.pattern = pattern
    style.alignment = alignment

    print sheet.col(0)

    

    sheet.write(1,0,'012', style)
    sheet.write(2,0,'Mon', style)
    sheet.write(3,0,'Tue', style)
    sheet.write(4,0,'Wed', style)
    sheet.write(5,0,'Thu', style)
    sheet.write(6,0,'Fri', style)

    sheet.write(1,1,'8.00-9.00', style)
    sheet.write(1,2,'9.00- \n 10.00', style)
    sheet.write(1,3,'10.00-11.00', style)
    sheet.write(1,4,'11.00-12.00', style)
    sheet.write(1,5,'12.00-13.00', style)
    sheet.write(1,6,'13.00-14.00', style)
    sheet.write(1,7,'14.00-15.00', style)
    sheet.write(1,8,'15.00-16.00', style)
    sheet.write(1,9,'16.00-17.00', style)
    sheet.write(1,10,'17.00-18.00', style)
    sheet.write(1,11,'18.00-19.00', style)


book = Workbook(encoding='UTF-8')
sheet = book.add_sheet('A Date')
#sheet.write(1,1,'00')#,style)
#sheet.write(2,2,' ก็ไม่รู้สินะ')#,style)
main_table(sheet)
book.save('date.xls')
