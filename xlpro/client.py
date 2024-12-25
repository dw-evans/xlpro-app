import win32com.client

xlpro = win32com.client.Dispatch('xlproServer.Application')

print(xlpro.getpid())
xlpro.testImports()


pass