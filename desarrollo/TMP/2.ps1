$Address=((Get-mailbox rodolfo.castelo@tecnofor.pe |Select *).PrimarySmtpAddress).toString()

$Address.Split("@")[1]

