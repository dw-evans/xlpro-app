import win32com.client

app = win32com.client.Dispatch("Python.PersistingCOMServer")

print(app.store_data("abc", 123))
print(app.store_data("abcd", 1234))

app2 = win32com.client.Dispatch("Python.PersistingCOMServer")
print(app2.get_data("abc"))


pass