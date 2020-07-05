@echo off
set yy=%date:~-4%
set dd=%date:~-7,2%
set mm=%date:~-10,2%
set MYDATE=%yy%_%mm%_%dd%
REM - DEBUG LAUNCH -> c:\Windows\System32\cmd.exe /c start perl ProvidentAvp3.pl %CD%\%MYDATE%\ProvidentAvp_output.csv Provident
c:\Windows\System32\cmd.exe /q /c start /min perl ProvidentAvp3.pl %CD%\%MYDATE%\ProvidentAvp_output.csv Provident
