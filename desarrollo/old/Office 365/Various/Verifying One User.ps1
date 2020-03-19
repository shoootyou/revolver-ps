$usuario = Read-Host '¿Qué usuario deseas verificar?'
Get-Mailbox $usuario | fl Name, Archive*
Get-Mailbox $usuario | Select RetentionPolicy