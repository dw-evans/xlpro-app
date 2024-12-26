import win32com.client

xlpro = win32com.client.Dispatch('xlproServer.Application')

# xlapp = win32com.client.Dispatch("Excel.Application")
# wb = xlapp.ActiveWorkbook

xlpro.register_functions_in_self()
xlpro.execute_function("add_numbers", 12, 13)
xlpro.executeFunction("add_numbers", 12, 13)





pass