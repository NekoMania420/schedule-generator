#!/usr/bin/env python
# -*- coding: UTF-8 -*-

"""
Project: Schedule.GEN
Version: 0.1
App by:
    Suchaj Jongprasit (57070132)
    Seksan Neramitthanasombat (57070137)

    Faculty of Information Technology
    King Mongkut's Institute of Technology Ladkrabang
"""

import sys
sys.path.append("mechanize")

from Tkinter import *
import tkMessageBox
import tkFont
import mechanize
import time

class RequestRegistra(object):

    def request_data(self):
        """Get raw schedule table from KMITL Registra."""

        self.start_time = time.time()
        br = mechanize.Browser()
        br.set_handle_robots(False)
        br.set_handle_refresh(False)
        br.addheaders = [("User-Agent", "Mozilla/5.0")]

        br.open("https://www.reg.kmitl.ac.th/user/index.php")

        if self.username == "" or self.passwd == "":
            tkMessageBox.showerror("Login error", "Please complete fill your username and password.")
        else:
            br.select_form("usermm")
            br["user_id"] = self.username
            br["password"] = self.passwd
            br.method = "POST"
            br.submit()

            br.open("http://www.reg.kmitl.ac.th/u_student/report_studytable.php?close_header=1")

            if self.year == "" or self.semester == "":
                tkMessageBox.showerror("Login error", "Please complete fill year and semester.")
            else:
                br.select_form("edit")
                br["year"] = [self.year]
                br["semester"] = [self.semester]
                response = br.submit()

                #soup = BeautifulSoup(response.get_data())
                #response.set_data(soup.prettify("utf-8"))

                self.write_to_file(response.get_data())

                self.export_string()

    def write_to_file(self, data):
        """Write data to file."""

        with open("data.html", "w") as f:
            f.write(data)
            f.close()

    def export_string(self):
        """Export HTML code to custom modules."""

        import string_filter
        
        with open("data.html", "r") as f:
            string_filter.string_filter(f)

        import gen

        gen.gen()

        tkMessageBox.showinfo("Done!", "Generate complete!\nTime: %.3f s.\n\nOpen 'Schedule.xls' to see schedule table.\n\nPlease use 'LibreOffice Calc' to open file." % (time.time()-self.start_time))


class MainWindow(object):

    def __init__(self, master):
        """Initial function."""

        self.master = master
        master.title("Schedule.GEN")
        master.geometry("290x220")
        master.resizable(0, 0)

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
        self.header = Label(master, font=(None, 20), text="Schedule.GEN")
        self.header.grid(row=0, columnspan=2, pady=20)

        self.username_label = Label(master, text="Username:")
        self.username_label.grid(row=1, sticky=E)

        self.username_input = Entry(master)
        self.username_input.grid(row=1, column=1, sticky=W)

        self.passwd_label = Label(master, text="Password:")
        self.passwd_label.grid(row=2, sticky=E)

        self.passwd_input = Entry(master, show="*")
        self.passwd_input.grid(row=2, column=1, sticky=W)

        option_area = Frame(master)
        option_area.grid(row=3, columnspan=2)

        self.year_label = Label(option_area, text="Year:")
        self.year_label.pack(side=LEFT)
        self.year_input = Entry(option_area, width=4)
        self.year_input.pack(side=LEFT)
        self.year_input.insert(END, "2557")

        self.semester_label = Label(option_area, text="Semester:")
        self.semester_label.pack(side=LEFT)
        self.semester_input = Entry(option_area, width=1)
        self.semester_input.pack(side=LEFT)
        self.semester_input.insert(END, "1")

        self.submit_button = Button(master, text="Get!", command=self.send_data)
        self.submit_button.grid(row=4, columnspan=2, pady=20)

        # Element configuration
        master.config(bg="#111")
        menubar.config(bg="#111", fg="#fff", activebackground="#fff", borderwidth=0)
        filemenu.config(bg="#111", fg="#fff", activebackground="#fff", activeforeground="#111")
        helpmenu.config(bg="#111", fg="#fff", activebackground="#fff", activeforeground="#111")
        self.header.config(bg="#111", fg="#fff")
        self.username_label.config(bg="#111", fg="#fff")
        self.username_input.config(bg="#111", fg="#fff", insertbackground="#fff")
        self.passwd_label.config(bg="#111", fg="#fff")
        self.passwd_input.config(bg="#111", fg="#fff", insertbackground="#fff")
        option_area.config(bg="#111")
        self.year_label.config(bg="#111", fg="#fff")
        self.year_input.config(bg="#111", fg="#fff", insertbackground="#fff")
        self.semester_label.config(bg="#111", fg="#fff")
        self.semester_input.config(bg="#111", fg="#fff", insertbackground="#fff")
        self.submit_button.config(bg="#111", fg="#fff", activebackground="#111", activeforeground="#fff", height=2, width=10)

        for col in xrange(2):
            master.grid_columnconfigure(col, weight=1)

    def send_data(self):
        """Call RequestRegistra() to request data."""

        req = RequestRegistra()
        req.username = self.username_input.get()
        req.passwd = self.passwd_input.get()
        req.year = self.year_input.get()
        req.semester = self.semester_input.get()
        try:
            req.request_data()
        except:
            tkMessageBox.showerror("Not found!", "Please check your username and password and try again.")

    def quit(self, event):
        """Quit program."""

        root.destroy()

    def call_about(self):
        """Open 'About' window."""

        About().mainloop()

    def how_to_use(self):
        """Open 'How to use?' window."""

        HowToUse().mainloop()


class About(Tk):

    def __init__(self):
        """Initial function."""

        Tk.__init__(self)

        self.title("About")
        self.geometry("400x300")
        self.resizable(0, 0)
        self.config(bg="#111")

        window_frame = Frame(self)
        window_frame.pack(expand=1)
        window_frame.config(bg="#111")

        Label(window_frame, text="Schedule.GEN", font="None 20", bg="#111", fg="#fff").pack()
        Label(window_frame, text="Generate schedule table from KMITL Registra",bg="#111", fg="#fff").pack()
        Label(window_frame, text="Version 0.1", bg="#111", fg="#fff").pack(pady=(20, 0))

        author_frame = Label(window_frame)
        author_frame.pack(pady=(20, 0))
        author_frame.config(bg="#111")

        Label(author_frame, text="Creators:", font=tkFont.Font(weight="bold"), bg="#111", fg="#fff").pack()
        Label(author_frame, text="Suchaj Jongprasit (57070132)", bg="#111", fg="#fff").pack()
        Label(author_frame, text="Seksan Neramitthanasombat (57070137)", bg="#111", fg="#fff").pack()

        button_area = Frame(window_frame)
        button_area.pack(pady=10)
        button_area.config(bg="#111")

        Button(button_area, text="License", command=self.call_license, bg="#111", fg="#fff", activebackground="#111", activeforeground="#fff").pack(side=LEFT)
        Button(button_area, text="Close", command=self.destroy, bg="#111", fg="#fff", activebackground="#111", activeforeground="#fff").pack(side=LEFT)

    def call_license(self):
        """Open 'License' window."""

        License().mainloop()


class HowToUse(Tk):

    def __init__(self):
        """Initial function."""

        Tk.__init__(self)

        self.title("How to use?")
        self.geometry("300x100")
        self.resizable(0, 0)
        self.config(bg="#111")

        label_frame = Frame(self)
        label_frame.pack(expand=1)
        label_frame.config(bg="#111")

        Label(label_frame, text="1. Input your username and password", bg="#111", fg="#fff").pack(anchor=W)
        Label(label_frame, text="2. Press 'Get!' button", bg="#111", fg="#fff").pack(anchor=W)

        Button(label_frame, text="Close", command=self.destroy, bg="#111", fg="#fff", activebackground="#111", activeforeground="#fff").pack()


class License(Tk):

    def __init__(self):
        """Initial function."""

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


if __name__ == "__main__":
    root = Tk()
    app = MainWindow(root)
    root.mainloop()
