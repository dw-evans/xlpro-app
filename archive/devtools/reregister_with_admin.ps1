Start-Process powershell -ArgumentList "-Command", "cd 'C:\Users\Daniel Evans\projects\xlpro'; 'tools_async/reregister.bat'; Start-Sleep -Seconds 0.5" -Verb RunAs
# ./tools async/refresh_excel.bat