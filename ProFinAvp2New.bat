@echo off
set yy=%date:~-4%
set dd=%date:~-7,2%
set mm=%date:~-10,2%
set MYDATE=%yy%_%mm%_%dd%
c:\Windows\System32\cmd.exe /q /c start /min python ProFinAvp3-Python.py %CD%\%MYDATE%\ProvidentAvp_output.csv 
