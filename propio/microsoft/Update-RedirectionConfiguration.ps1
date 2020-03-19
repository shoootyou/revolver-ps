$DB_FOR = Import-Csv C:\Windows\Scripts\ForwardUsers.csv
$LG_PRO = "C:\Windows\Scripts\" + (Get-Date -Format "yyyyMMdd-hhmm") + "-ChangeProgress.log"
$TX_OOF = Get-Content C:\Windows\Scripts\Engie-OOF.txt

$FL_TST = $true #Flag para activar prueba o desactivarla
$US_TST = 'mah@cam-la.com'

"Configured on date,PrimarySmtpAddress,Forward status,OOF status" | Out-File $LG_PRO -Append

if($FL_TST){
	$USR_TST = $DB_FOR | ? { $_.PrimarySmtpAddress -eq $US_TST}
	if($USR_TST.ForwardingAddress){
		try{
			Set-Mailbox -Identity $USR_TST.Identity -ForwardingAddress $null
			$FORWARD = $false
		}
		catch{
			$FORWARD = $true
		}
	}
	if(!$FORWARD){
		Set-MailboxAutoReplyConfiguration -Identity $USR_TST.Identity -AutoReplyState Enabled -InternalMessage $TX_OOF -ExternalMessage $TX_OOF
		(Get-Date -Format "yyyyMMdd-hhmm") + "," + $USR_TST.PrimarySmtpAddress + ",Forward deactivated,OOF configured" | Out-File $LG_PRO -Append
	}
	else{
		(Get-Date -Format "yyyyMMdd-hhmm") + "," + $USR_TST.PrimarySmtpAddress + ",Forward can't deactivate,OOF not' configured" | Out-File $LG_PRO -Append
	}
}
else{
	foreach($USR_TST in $DB_FOR){
		if($USR_TST.ForwardingAddress){
			try{
				Set-Mailbox -Identity $USR_TST.Identity -ForwardingAddress $null
				$FORWARD = $false
			}
			catch{
				$FORWARD = $true
			}
		}
		if(!$FORWARD){
			Set-MailboxAutoReplyConfiguration -Identity $USR_TST.Identity -AutoReplyState Enabled -InternalMessage $TX_OOF -ExternalMessage $TX_OOF
			(Get-Date -Format "yyyyMMdd-hhmm") + "," + $USR_TST.PrimarySmtpAddress + ",Forward deactivated,OOF configured" | Out-File $LG_PRO -Append
		}
		else{
			(Get-Date -Format "yyyyMMdd-hhmm") + "," + $USR_TST.PrimarySmtpAddress + ",Forward can't deactivate,OOF not' configured" | Out-File $LG_PRO -Append
		}
	}
}