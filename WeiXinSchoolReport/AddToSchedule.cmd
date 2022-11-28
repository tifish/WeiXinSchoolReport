schtasks /create /f /tn "My\WeXin School Report" /sc daily /st 07:30 /tr "%~dp0WeiXinSchoolReport.exe"
pause
