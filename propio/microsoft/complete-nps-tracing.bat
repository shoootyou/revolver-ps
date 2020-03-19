netsh trace stop
netsh ras set tracing * disabled
wevtutil.exe sl Microsoft-Windows-CAPI2/Operational /e:false
wevtutil.exe epl Microsoft-Windows-CAPI2/Operational C:\MSLOG\%COMPUTERNAME%_CAPI2.evtx
xcopy %Systemroot%\Tracing C:\MSLOG\Tracing /s /i /e /h /r /k /y
gpresult /H C:\MSLOG\%COMPUTERNAME%_gpresult.htm
msinfo32 /report c:\MSLOG\%COMPUTERNAME%_msinfo32.txt
ipconfig /all > c:\MSLOG\%COMPUTERNAME%_ipconfig.txt
route print > c:\MSLOG\%COMPUTERNAME%_route_print.txt
wevtutil epl Application c:\MSLOG\%COMPUTERNAME%_Application.evtx
wevtutil epl System c:\MSLOG\%COMPUTERNAME%_System.evtx
wevtutil epl Security c:\MSLOG\%COMPUTERNAME%_Security.evtx
wevtutil epl Microsoft-Windows-GroupPolicy/Operational C:\MSLOG\%COMPUTERNAME%_GroupPolicy_Operational.evtx
wevtutil epl "Microsoft-Windows-WLAN-AutoConfig/Operational" c:\MSLOG\%COMPUTERNAME%_Microsoft-Windows-WLAN-AutoConfig-Operational.evtx
wevtutil epl "Microsoft-Windows-Wired-AutoConfig/Operational" c:\MSLOG\%COMPUTERNAME%_Microsoft-Windows-Wired-AutoConfig-Operational.evtx
wevtutil epl Microsoft-Windows-CertificateServicesClient-CredentialRoaming/Operational c:\MSLOG\%COMPUTERNAME%_CertificateServicesClient-CredentialRoaming_Operational.evtx
wevtutil epl Microsoft-Windows-CertPoleEng/Operational c:\MSLOG\%COMPUTERNAME%_CertPoleEng_Operational.evtx
wevtutil epl Microsoft-Windows-CertificateServicesClient-Lifecycle-System/Operational c:\MSLOG\%COMPUTERNAME%_CertificateServicesClient-Lifecycle-System_Operational.evtx
wevtutil epl Microsoft-Windows-CertificateServicesClient-Lifecycle-User/Operational c:\MSLOG\%COMPUTERNAME%_CertificateServicesClient-Lifecycle-User_Operational.evtx
wevtutil epl Microsoft-Windows-CertificateServices-Deployment/Operational c:\MSLOG\%COMPUTERNAME%_CertificateServices-Deployment_Operational.evtx
certutil -v -silent -store MY > c:\MSLOG\%COMPUTERNAME%_cert-Personal-Registry.txt
certutil -v -silent -store ROOT > c:\MSLOG\%COMPUTERNAME%_cert-TrustedRootCA-Registry.txt
certutil -v -silent -store -grouppolicy ROOT > c:\MSLOG\%COMPUTERNAME%_cert-TrustedRootCA-GroupPolicy.txt
certutil -v -silent -store -enterprise ROOT > c:\MSLOG\%COMPUTERNAME%_TrustedRootCA-Enterprise.txt
certutil -v -silent -store TRUST > c:\MSLOG\%COMPUTERNAME%_cert-EnterpriseTrust-Reg.txt
certutil -v -silent -store -grouppolicy TRUST > c:\MSLOG\%COMPUTERNAME%_cert-EnterpriseTrust-GroupPolicy.txt
certutil -v -silent -store -enterprise TRUST > c:\MSLOG\%COMPUTERNAME%_cert-EnterpriseTrust-Enterprise.txt
certutil -v -silent -store CA > c:\MSLOG\%COMPUTERNAME%_cert-IntermediateCA-Registry.txt
certutil -v -silent -store -grouppolicy CA > c:\MSLOG\%COMPUTERNAME%_cert-IntermediateCA-GroupPolicy.txt
certutil -v -silent -store -enterprise CA > c:\MSLOG\%COMPUTERNAME%_cert-Intermediate-Enterprise.txt
certutil -v -silent -store AuthRoot > c:\MSLOG\%COMPUTERNAME%_cert-3rdPartyRootCA-Registry.txt
certutil -v -silent -store -grouppolicy AuthRoot > c:\MSLOG\%COMPUTERNAME%_cert-3rdPartyRootCA-GroupPolicy.txt
certutil -v -silent -store -enterprise AuthRoot > c:\MSLOG\%COMPUTERNAME%_cert-3rdPartyRootCA-Enterprise.txt
certutil -v -silent -store SmartCardRoot > c:\MSLOG\%COMPUTERNAME%_cert-SmartCardRoot-Registry.txt
certutil -v -silent -store -grouppolicy SmartCardRoot > c:\MSLOG\%COMPUTERNAME%_cert-SmartCardRoot-GroupPolicy.txt
certutil -v -silent -store -enterprise SmartCardRoot > c:\MSLOG\%COMPUTERNAME%_cert-SmartCardRoot-Enterprise.txt
certutil -v -silent -store -enterprise NTAUTH > c:\MSLOG\%COMPUTERNAME%_cert-NtAuth-Enterprise.txt
certutil -v -silent -user -store MY > c:\MSLOG\%COMPUTERNAME%_cert-User-Personal-Registry.txt
certutil -v -silent -user -store ROOT > c:\MSLOG\%COMPUTERNAME%_cert-User-TrustedRootCA-Registry.txt
certutil -v -silent -user -store -enterprise ROOT > c:\MSLOG\%COMPUTERNAME%_cert-User-TrustedRootCA-Enterprise.txt
certutil -v -silent -user -store TRUST > c:\MSLOG\%COMPUTERNAME%_cert-User-EnterpriseTrust-Registry.txt
certutil -v -silent -user -store -grouppolicy TRUST > c:\MSLOG\%COMPUTERNAME%_cert-User-EnterpriseTrust-GroupPolicy.txt
certutil -v -silent -user -store CA > c:\MSLOG\%COMPUTERNAME%_cert-User-IntermediateCA-Registry.txt
certutil -v -silent -user -store -grouppolicy CA > c:\MSLOG\%COMPUTERNAME%_cert-User-IntermediateCA-GroupPolicy.txt
certutil -v -silent -user -store Disallowed > c:\MSLOG\%COMPUTERNAME%_cert-User-UntrustedCertificates-Registry.txt
certutil -v -silent -user -store -grouppolicy Disallowed > c:\MSLOG\%COMPUTERNAME%_cert-User-UntrustedCertificates-GroupPolicy.txt
certutil -v -silent -user -store AuthRoot > c:\MSLOG\%COMPUTERNAME%_cert-User-3rdPartyRootCA-Registry.txt
certutil -v -silent -user -store -grouppolicy AuthRoot > c:\MSLOG\%COMPUTERNAME%_cert-User-3rdPartyRootCA-GroupPolicy.txt
certutil -v -silent -user -store SmartCardRoot > c:\MSLOG\%COMPUTERNAME%_cert-User-SmartCardRoot-Registry.txt
certutil -v -silent -user -store -grouppolicy SmartCardRoot > c:\MSLOG\%COMPUTERNAME%_cert-User-SmartCardRoot-GroupPolicy.txt
certutil -v -silent -user -store UserDS > c:\MSLOG\%COMPUTERNAME%_cert-User-UserDS.txt
netsh lan show interfaces > c:\MSLOG\%computername%_lan_interfaces.txt
netsh lan show profiles > c:\MSLOG\%computername%_lan_profiles.txt
netsh lan show settings > c:\MSLOG\%computername%_lan_settings.txt
netsh lan export profile folder=c:\MSLOG\