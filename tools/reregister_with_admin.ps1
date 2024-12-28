Start-Process powershell -ArgumentList "-Command", "cd 'C:\Users\Daniel Evans\projects\xlpro'; tools/reregister.bat; Start-Sleep -Seconds 0.5" -Verb RunAs
./tools/refresh_excel.bat