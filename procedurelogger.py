from selenium import webdriver
from procedureloggertools import *
import tkinter as tk
import tkinter.simpledialog
from tkinter.messagebox import showinfo

import time


tk.Tk().withdraw()
# username =  tkinter.simpledialog.askstring("Username", "Enter username:")
# password = tkinter.simpledialog.askstring("Password", "Enter password:", show='*')

# excel_file = r"E:\OneDrive\Documents\Residency\procedurelog.xlsx"
excel_file = r"C:\Users\Jon\Documents\GitHub\procedurelogger\procedurelog_jl2025.xlsx"

database = get_log_data(excel_file)

browser = webdriver.Firefox(executable_path=r'geckodriver.exe')
# browser = webdriver.Chrome()
browser.get("https://www.new-innov.com/login/Login.aspx")

test = showinfo(title='Login',message='Please log in and then press OK')

i=0
for data in database:
    i = i+1
    print(str(i) + "/" + str(len(database)))

    if data[8]=='no':
        try:
            log_procedure(browser, data)
        except Exception as e:
            print('***FAILURE START***')
            print(e)
            print(data)
            print('***FAILURE END***')

browser.close()