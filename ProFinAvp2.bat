@echo off
set yy=%date:~-4%
set dd=%date:~-7,2%
set mm=%date:~-10,2%
set MYDATE=%yy%_%mm%_%dd%
echo %CD%
rem timeout /t 20
c:\Windows\System32\cmd.exe /q /c start /min python -m ProFinAvp3-Python.py %CD%\%MYDATE%\ProvidentAvp_output.csv 
rem start chrome %CD%\%MYDATE%\ProvidentAvpFinal.html

