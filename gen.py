#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# check color
'''
จำนวนวิชา
เวลา(list)
วัน(list)
ชื่อวิชา+ตึก+ห้อง(list)
'''
#from datetime import date
import sys
sys.path.append("xlwt")

from xlwt import *
import random
import tkFileDialog

def gen():

    book = Workbook(encoding='UTF-8')
    sheet = book.add_sheet('A Date')
    list_color= [0x21, 0x0B, 0x35, 0x11, 0x016, 0x012, 0x0D, 0x10, 0x17, 0x0C, 0x0F]
    fnt = Font()
    fnt.name = 'Arial'
    list_name = []
    list_day=[]
    list_time=[]
    change_day = ['จ.', 'Mon', 'อ.', 'Tue', 'พ.' ,'Wed', 'พฤ.', 'Thu', 'ศ.', 'Fri']
    with open("filtered.txt", "r+") as f:
        #for i in xrange()
        #aaa = f.read().index('\n')
        aaa = f.read().split('\n')

        num = (len(aaa)/5)
        for i in xrange(len(aaa)-3):
            if i % 5==0:
                list_name.append(aaa[i]+' (' +(aaa[i+3].translate(None, ' '))+', '+ aaa[i+4]+') ['+aaa[i+1]+']')
                list_day.append(change_day[change_day.index(aaa[i+2][:aaa[i+2].index(' ')])+1])
                list_time.append((aaa[i+2][aaa[i+2].index(' ')+1:len(aaa[i+2])-5]))

        f.close()
    print num
    print list_name
    print list_time
    print list_day

    for i in xrange(len(list_name)):

        if list_name[i].count('INFORMATION TECHNOLOGY FUNDAMENTALS')>=1:
            list_name.append('INFORMATION TECHNOLOGY FUNDAMENTALS [Lab]')
            list_day.append('Thu')
            num += 1
            if list_name[i].count('[1]')>=1:
                list_time.append('09:00-11:00')
            elif list_name[i].count('[2]')>=1:
                list_time.append('12:00-14:00')
            else:
                list_time.append('14:00-16:00')
            list_name[i]=list_name[i][:len(list_name[i])-(list_name[i].index(']')-list_name[i].index('[')+1)]
        elif list_name[i].count('PROBLEM SOLVING IN INFORMATION TECHNOLOGY')>=1:
            list_name.append('PROBLEM SOLVING IN INFORMATION TECHNOLOGY [Lab]')
            list_day.append('Thu')
            num += 1
            if list_name[i].count('[1]')>=1:
                list_time.append('14:00-16:00')
            elif list_name[i].count('[2]')>=1:
                list_time.append('09:00-11:00')
            else:
                list_time.append('12:00-14:00')
            list_name[i]=list_name[i][:len(list_name[i])-(list_name[i].index(']')-list_name[i].index('[')+1)]
        elif list_name[i].count('INTRODUCTION TO COMPUTER SYSTEMS')>=1:
            list_name.append('INTRODUCTION TO COMPUTER SYSTEMS [Lab]')
            list_day.append('Thu')
            num += 1
            if list_name[i].count('[1]')>=1:
                list_time.append('12:00-14:00')
            elif list_name[i].count('[2]')>=1:
                list_time.append('14:00-16:00')
            else:
                list_time.append('09:00-11:00')
            list_name[i]=list_name[i][:len(list_name[i])-(list_name[i].index(']')-list_name[i].index('[')+1)]
        elif list_name[i].count('COMPUTER PROGRAMMING')>=1:
            list_name.append('COMPUTER PROGRAMMING [Lab]')
            list_day.append('Fri')
            num += 1
            if list_name[i].count('[1]')>=1:
                list_time.append('09:00-11:00')
            elif list_name[i].count('[2]')>=1:
                list_time.append('12:00-14:00')
            else:
                list_time.append('14:00-16:00')
            list_name[i]=list_name[i][:len(list_name[i])-(list_name[i].index(']')-list_name[i].index('[')+1)]
        elif list_name[i].count('MULTIMEDIA AND WEB TECHNOLOGY')>=1:
            list_name.append('MULTIMEDIA AND WEB TECHNOLOGY [Lab]')
            list_day.append('Fri')
            num += 1
            if list_name[i].count('[1]')>=1:
                list_time.append('12:00-14:00')
            elif list_name[i].count('[2]')>=1:
                list_time.append('14:00-16:00')
            else:
                list_time.append('09:00-11:00')
            list_name[i]=list_name[i][:len(list_name[i])-(list_name[i].index(']')-list_name[i].index('[')+1)]
        else:
            list_name[i]=list_name[i][:len(list_name[i])-(list_name[i].index(']')-list_name[i].index('[')+1)]

    #print list_name
    main_day = ['0', '1', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    main_time = ['0', '08:00', '08:30', '09:00', '09:30', '10:00'
                , '10:30', '11:00', '11:30', '12:00', '12:30'
                , '13:00', '13:30', '14:00', '14:30', '15:00'
                , '15:30', '16:00', '16:30', '17:00', '17:30'
                , '18:00', '18:30', '19:00', '19:30', '20:00']

    borders = Borders()
    borders.left = Borders.THICK
    borders.right = Borders.THICK
    borders.top = Borders.THICK
    borders.bottom = Borders.THICK

    pattern = Pattern()
    pattern.pattern = Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = random.choice(list_color)
    print pattern.pattern_fore_colour
    color = pattern.pattern_fore_colour , '+'
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

    #print sheet.col(0)
    sheet.write(1,0,'', style)
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
    list_test=[]
    print num
    for i in xrange(num):
        pattern2 = Pattern()
        pattern2.pattern = Pattern.SOLID_PATTERN
        while True:
            pattern2.pattern_fore_colour = random.choice(list_color)
            if color.count(pattern2.pattern_fore_colour)== 0:
                color = pattern2.pattern_fore_colour , '+'
                break
        style2 = XFStyle()
        style2.pattern = pattern2
        style2.alignment = alignment
        style2.borders = borders
        print 1
        sheet.write_merge(main_day.index(list_day[i]),main_day.index(list_day[i]),main_time.index(list_time[i][:5]),main_time.index(list_time[i][6:])-1,list_name[i], style2)
        list_test.append([main_day.index(list_day[i]),main_day.index(list_day[i]),main_time.index(list_time[i][:5]),main_time.index(list_time[i][6:])-1])
    #print sorted(list_test)
    list_test = sorted(list_test)
    print list_test
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
        list_test.append([j, j, 25, 26])
        list_test = sorted(list_test)
        #sheet.write_merge(j,j,25,26,'', style4)
        while True:

            if cou > 24:
                break
            if list_test[i][0] != j:
                break
            if list_test[i][2] >= cou+2:
                sheet.write_merge(j,j,cou,cou+1,'', style3)
                cou += 2
            elif list_test[i][2] >= cou+1:
                sheet.write_merge(j,j,cou,cou,'', style3)
                cou += 1
            else:
                if list_test[i][3] % 2 != 0:
                    cou = list_test[i][3]+1
                    #sheet.write_merge(j,j,cou,cou,'', style3)
                    cou +=1
                else:
                    cou = list_test[i][3]+1
                i+=1
        i+=1
    #sheet.write_merge(3,3,24,24,'', style3)

    sheet.portrait = False
    book.save('Schedule.xls')
gen()