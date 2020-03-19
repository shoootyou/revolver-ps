mkdir C:\MSLOG
netsh ras set tracing * enabled
netsh trace start scenario=lan globallevel=0xff capture=yes maxsize=1024 tracefile=C:\MSLOG\%COMPUTERNAME%_wired_cli.etl
wevtutil.exe sl Microsoft-Windows-CAPI2/Operational /e:true
wevtutil sl Microsoft-Windows-CAPI2/Operational /ms:104857600