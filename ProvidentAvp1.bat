@echo off
c:\Windows\System32\cmd.exe /q /c start /min ProvidentAvp1.pl AVPSalesReport.csv
set yy=%date:~-4%
set dd=%date:~-7,2%
set mm=%date:~-10,2%
set MYDATE=%yy%_%mm%_%dd%
start firefox %CD%\%MYDATE%\ProvidentAvpFile.html
