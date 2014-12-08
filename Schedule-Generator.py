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
import tkMessageBox
import tkFont
import mechanize
import getpass
import time
from bs4 import BeautifulSoup

class RequestRegistra(object):


    def request_data(self):
        '''Get raw schedule table from KMITL Registra.'''
        self.start_time = time.time()
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

        soup = BeautifulSoup(response.get_data())
        response.set_data(soup.prettify('utf-8'))

        self.write_to_file(response.get_data())

    def write_to_file(self, data):
        '''Write data to file'''

        with open('data.html','w+') as f:
            f.write(data)
            f.close()

        tkMessageBox.showinfo('Done!', '%.3f s.' % (time.time()-self.start_time))


class MainWindow(object):

    def __init__(self, master):
        '''Initial function.'''
        self.master = master
        master.title('Schedule.GEN')
        master.geometry('290x120')
        master.resizable(width=False, height=False)

        # Menubar
        menubar = Menu(master)

        filemenu = Menu(menubar, tearoff=False)
        filemenu.add_command(label="Quit", command=quit, underline=0, accelerator="Ctrl+Q")
        menubar.add_cascade(label="File", menu=filemenu, underline=0)

        helpmenu = Menu(menubar, tearoff=False)
        helpmenu.add_command(label="How to use?", command=self.how_to_use, underline=0)
        helpmenu.add_separator()
        helpmenu.add_command(label="About...", command=self.call_about, underline=0)
        menubar.add_cascade(label="Help", menu=helpmenu, underline=0)

        master.config(menu=menubar)

        master.bind_all("<Control-q>", self.quit)

        # Contents
        self.header = Label(master, font=(None, 20), text='ล็อกอินเข้าสู่ระบบ')
        self.header.grid(row=0, columnspan=2)

        self.username_label = Label(master, text='Username:')
        self.username_label.grid(row=1, sticky=E)

        self.username_input = Entry(master)
        self.username_input.grid(row=1, column=1, sticky=W)

        self.passwd_label = Label(master, text='Password:')
        self.passwd_label.grid(row=2, sticky=E)

        self.passwd_input = Entry(master, show="*")
        self.passwd_input.grid(row=2, column=1, sticky=W)

        self.submit_button = Button(master, text='Get!', command=self.send_data)
        self.submit_button.grid(row=3, columnspan=2)

        for col in xrange(2):
            master.grid_columnconfigure(col, weight=1)

    def send_data(self):
        req = RequestRegistra()
        req.username = self.username_input.get()
        req.passwd = self.passwd_input.get()
        req.request_data()

    def quit(self, event):
    	root.destroy()

    def call_about(self):
    	About().mainloop()

    def how_to_use(self):
    	HowToUse().mainloop()


class About(Tk):

	def __init__(self):
		Tk.__init__(self)

		self.title("About")
		self.geometry("400x300")
		self.resizable(width=False, height=False)

		window_frame = Frame(self)
		window_frame.pack(expand=1)

		Label(window_frame, text="Schedule.GEN", font="None 20").pack()
		Label(window_frame, text="Generate schedule table from KMITL Registra").pack()
		Label(window_frame, text="Version 0.01").pack(pady=(20, 0))

		author_frame = Label(window_frame)
		author_frame.pack(pady=(20, 0))

		Label(author_frame, text="Creators:", font=tkFont.Font(weight="bold")).pack()
		Label(author_frame, text="Suchaj Jongprasit (57070132)").pack()
		Label(author_frame, text="Seksan Neramitthanasombat (57070137)").pack()

		button_area = Frame(window_frame)
		button_area.pack(pady=10)

		Button(button_area, text="License", command=self.call_license).pack(side=LEFT)
		Button(button_area, text="Close", command=self.destroy).pack(side=LEFT)

	def call_license(self):
		License().mainloop()


class HowToUse(Tk):

	def __init__(self):
		Tk.__init__(self)

		self.title("How to use?")
		self.geometry("300x70")
		self.resizable(width=False, height=False)

		label_frame = Frame(self)
		label_frame.pack(expand=1)

		Label(label_frame, text="1. Input your username and password").pack(anchor=W)
		Label(label_frame, text="2. Press 'Get!' button").pack(anchor=W)


class License(Tk):

	def __init__(self):
		Tk.__init__(self)

		self.title("License")

		string = open("LICENSE", "r").read()

		scrollbar = Scrollbar(self)
		scrollbar.pack(side=RIGHT, fill=Y)

		textarea = Text(self, width=80, yscrollcommand=scrollbar.set)
		textarea.insert(END, string)
		textarea.config(state=DISABLED)
		textarea.pack(fill=BOTH, expand=1)

		scrollbar.config(command=textarea.yview)


if __name__ == '__main__':
    root = Tk()
    app = MainWindow(root)
    root.mainloop()
