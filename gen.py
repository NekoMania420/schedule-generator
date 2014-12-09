#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# ssss
'''
จำนวนวิชา
เวลา(list)
วัน(list)
ชื่อวิชา+ตึก+ห้อง(list)
'''
from datetime import date
from xlwt import *
import random
list_color= [0x21, 0x0B, 0x35, 0x11, 0x016, 0x012, 0x0D, 0x10, 0x17, 0x0C, 0x0F]
fnt = Font()
fnt.name = 'Arial'

num = 8
list_time = ['09:00-12:00', '13:00-15:00', '13:00-16:00', '09:00-11:00', '15:00-18:00', '09:00-12:00', '14:00-16:00', '09:00-11:00']
list_day = ['Thu', 'Thu', 'Tue', 'Wed', 'Mon', 'Mon', 'Fri', 'Fri']
list_name = ['MATHEMATICS FOR INFORMATION TECHNOLOGY'
    , 'COMPUTER PROGRAMMING', 'COMPUTER SYSTEMS ORGANIZATION AND OPERATING SYSTEM'
    , 'MULTIMEDIA AND WEB TECHNOLOGY', 'THAI GEOSOCIAL DESIGN'
    , 'FOUNDATION ENGLISH 2', 'COMPUTER PROGRAMMING[Lab]', 'MULTIMEDIA AND WEB TECHNOLOGY[Lab]']

main_day = ['0', '1', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri']
main_time = ['0', '08:00', '08:30', '09:00', '09.30', '10.00'
            , '10:30', '11:00', '11:30', '12:00', '12:30'
            , '13:00', '13:30', '14:00', '14:30', '15:00'
            , '15:30', '16:00', '16:30', '17:00', '17:30'
            , '18:00', '18:30', '19:00', '19:30', '20:00']
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
    pattern.pattern_fore_colour = random.choice(list_color)
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

    sheet.write_merge(1,1,1,2,'08.00-09.00', style)
    sheet.write_merge(1,1,3,4,'09.00-10.00', style)
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
    '''
    pattern2 = Pattern()
    pattern2.pattern = Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = random.choice(list_color)
    style2 = XFStyle()
    style2.pattern = pattern2
    style2.alignment = alignment
    sheet.write_merge(2,2,1,4,'BUSINESS FUNDAMENTALS FOR INFORMATION TECHNOLOGY \n (ไอที M23)', style2)
    '''
    list_test=[]
    for i in xrange(num):
        pattern2 = Pattern()
        pattern2.pattern = Pattern.SOLID_PATTERN
        pattern2.pattern_fore_colour = random.choice(list_color)
        style2 = XFStyle()
        style2.pattern = pattern2
        style2.alignment = alignment
        style2.borders = borders
        sheet.write_merge(main_day.index(list_day[i]),main_day.index(list_day[i]),main_time.index(list_time[i][:5]),main_time.index(list_time[i][6:])-1,list_name[i], style2)
        list_test.append([main_day.index(list_day[i]),main_day.index(list_day[i]),main_time.index(list_time[i][:5]),main_time.index(list_time[i][6:])-1])
    print sorted(list_test)
    list_test = sorted(list_test)
    pattern3 = Pattern()
    pattern3.pattern = Pattern.SOLID_PATTERN
    pattern3.pattern_fore_colour = 0x09
    style3 = XFStyle()
    style3.pattern = pattern3
    style3.alignment = alignment
    style3.borders = borders
    i=0
    for j in xrange(2, 7):
        cou = 1
        #print 2
        while True:
            #print 3
            if cou>24:
                break
            if list_test[i][0] != j:
                break
            if list_test[i][2] >= cou+2:
                sheet.write_merge(j,j,cou,cou+1,'', style3)
                #print '1'
                cou += 2
            else:
                cou = list_test[i][3]+1
                i+=1
                if i >= len(list_test):
                    
                    for m in xrange(cou, 24, 2):
                        sheet.write_merge(j,j,m,m+1,'', style3)
                    break
                if list_test[i][0] != j:
                    for m in xrange(cou, 24, 2):
                        sheet.write_merge(j,j,m,m+1,'', style3)
                    break
            #print list_test[i]

    
book = Workbook(encoding='UTF-8')
sheet = book.add_sheet('A Date')

#sheet.write(1,1,'00')#,style)
#sheet.write(2,2,' ก็ไม่รู้สินะ')#,style)
main_table(sheet)
sheet.portrait = False
book.save('date.xls')
