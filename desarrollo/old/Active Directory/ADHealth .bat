@Echo Off
ECHO Running AD Health Checks - Notepad will open after completion
set logfile=%userprofile%\Desktop\ADHealth.txt
echo You can share this log using http://pastie.org/pastes/new > %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM Finds system boot time
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                   System Boot Time >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
systeminfo | find "System Boot Time:" >> %logfile%
systeminfo | find "System Up Time:" >> %logfile%
systeminfo | find "Original Install Date:" >> %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM Displays all current TCP/IP network configuration values
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                     IPCONFIG >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
ipconfig /all >> %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM Analyse the state of domain controllers in a forest and reports any problems to assist in troubleshooting
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                     DCDIAG >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
dcdiag /a >> %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM The replsummary operation quickly summarizes the replication state and relative health
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                  Replsummary >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
repadmin /replsummary >> %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM Displays the replication partners for each directory partition on the specified domain controller
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                     Showrepl >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
repadmin /showrepl >> %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM Query FSMO roles
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                               NETDOM Query FSMO >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
netdom query fsmo >> %logfile%
echo. >> %logfile%
echo. >> %logfile%
REM Query Global Catalogs
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
echo                                                  List Global Catalogs >> %logfile%
echo ------------------------------------------------------------------------------------------------------------------------ >> %logfile%
for /f "tokens=2" %%a in ('systeminfo ^| findstr Domain:') do set domain=%%a
nslookup -querytype=srv _gc._tcp.%domain% >> %logfile%
notepad %logfile%