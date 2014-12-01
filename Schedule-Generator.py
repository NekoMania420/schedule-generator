#!/usr/bin/env python
# -*- coding: UTF-8 -*-

'''
Project: Schedule.GEN
Version: 0.01
App by:
    Suchaj Jongprasit (57070132)
    Seksan Neramitthanasombat (57070132)
IT@KMITL
'''

from Tkinter import *
import mechanize
import getpass

class RequestRegistra(object):


    def request_data(self):
        '''Get raw schedule table from KMITL Registra.'''

        br = mechanize.Browser()
        br.set_handle_robots(False)
        br.set_handle_refresh(False)
        br.addheaders = [('User-Agent', 'Mozilla/5.0')]

        br.open('https://www.reg.kmitl.ac.th/user/index.php')

        br.select_form('usermm')
        br['user_id'] = self.username
        br['password'] = self.passwd
        br.method = 'POST'
        br.submit()

        br.open('http://www.reg.kmitl.ac.th/u_student/report_studytable.php?close_header=1')

        br.select_form('edit')
        #year = br.possible_items('year')
        #semester = br.possible_items('semester')
        br['year'] = ['2557']
        br['semester'] = ['1']
        response = br.submit()
        self.write_to_file(response.get_data())

    def write_to_file(self, data):
        '''Write data to file'''

        with open('data.html','w+') as f:
            f.write(data)
            f.close()

class MainWindow(object):

    def __init__(self, master):
        '''Initial function.'''
        self.master = master
        master.title('Login')
        master.geometry('300x200')
        master.resizable(width=False, height=False)
        master.configure(bg='#FB8525')

        self.header = Label(master, font=(None, 20), text='ล็อกอินเข้าสู่ระบบ')
        self.header.pack()

if __name__ == '__main__':
    #req = RequestRegistra()
    #req.username = raw_input()
    #req.passwd = getpass.getpass()
    #req.request_data()
    root = Tk()
    app = MainWindow(root)
    root.mainloop()
