@Echo Off
ECHO Running AD Health Checks - Notepad will open after completion
set logfile=%userprofile%\Desktop\HW.txt
echo. >> %logfile%
echo. >> %logfile%
REM Finds system boot time
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                   System Boot Time >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
systeminfo >> %logfile%
notepad %logfile%