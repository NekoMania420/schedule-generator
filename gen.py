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
    borders = Borders()
    borders.left = Borders.THICK
    borders.right = Borders.THICK
    borders.top = Borders.THICK
    borders.bottom = Borders.THICK

    pattern = Pattern()
    pattern.pattern = Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 0x17
    #pattern.easyxf("align: horiz center")
    first_col = sheet.col(0)
    first_col.width = 256 * 6
    for i in xrange(1, 25):
        first_col = sheet.col(i)
        first_col.width = 256 * 5
        if i <= 6:
            first_row = sheet.row(i)
            #first_row.height_mismatch = True
            first_row.height = 256 * 5

    alignment = Alignment()
    alignment.horz = Alignment.HORZ_CENTER
    alignment.vert = Alignment.VERT_CENTER


    style = XFStyle()
    style.borders = borders
    style.pattern = pattern
    style.alignment = alignment
    style.alignment.wrap = 1

    print sheet.col(0)

    #sheet.write_merge(0, 0, 0, 1, 'Long Cell', style)

    sheet.write(1,0,'012', style)
    sheet.write(2,0,'Mon', style)
    sheet.write(3,0,'Tue', style)
    sheet.write(4,0,'Wed', style)
    sheet.write(5,0,'Thu', style)
    sheet.write(6,0,'Fri', style)

    sheet.write_merge(1,1,1,2,'8.00-9.00', style)
    sheet.write_merge(1,1,3,4,'9.00-10.00', style)
    sheet.write_merge(1,1,5,6,'10.00-11.00', style)
    sheet.write_merge(1,1,7,8,'11.00-12.00', style)
    sheet.write_merge(1,1,9,10,'12.00-13.00', style)
    sheet.write_merge(1,1,11,12,'13.00-14.00', style)
    sheet.write_merge(1,1,13,14,'14.00-15.00', style)
    sheet.write_merge(1,1,15,16,'15.00-16.00', style)
    sheet.write_merge(1,1,17,18,'16.00-17.00', style)
    sheet.write_merge(1,1,19,20,'17.00-18.00', style)
    sheet.write_merge(1,1,21,22,'18.00-19.00', style)
    sheet.write_merge(1,1,23,24,'19.00-20.00', style)

    pattern2 = Pattern()
    pattern2.pattern = Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = 0x0C
    style2 = XFStyle()
    style2.pattern = pattern2
    style2.alignment = alignment
    sheet.write_merge(2,2,1,4,'BUSINESS FUNDAMENTALS FOR INFORMATION TECHNOLOGY \n (ไอที M23)', style2)

book = Workbook(encoding='UTF-8')
sheet = book.add_sheet('A Date')

#sheet.write(1,1,'00')#,style)
#sheet.write(2,2,' ก็ไม่รู้สินะ')#,style)
main_table(sheet)
sheet.portrait = False
book.save('date.xls')
