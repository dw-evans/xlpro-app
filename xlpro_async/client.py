import win32com.client
import datetime, os, threading

from server import xlproServerAsync

xlpro = win32com.client.Dispatch(xlproServerAsync._reg_clsid_)
print(xlpro.add_data("1"))

xlpro2 = win32com.client.Dispatch(xlproServerAsync._reg_clsid_)
print(xlpro.add_data("2"))


pass