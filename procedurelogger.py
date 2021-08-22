from selenium import webdriver
from procedureloggertools import *
import tkinter as tk
import tkinter.simpledialog

username = 'jonlin'
tk.Tk().withdraw()
password = tkinter.simpledialog.askstring("Password", "Enter password:", show='*')

excel_file = "E:\OneDrive\Documents\procedurelog.xlsx"

database = get_log_data(excel_file)

browser = webdriver.Firefox(executable_path=r'C:\geckodriver\geckodriver.exe')
browser.get("https://www.new-innov.com/login/Login.aspx")

login(browser, username, password)
for data in database:
    if data[8]=='no':
        log_procedure(browser, data)

browser.close()