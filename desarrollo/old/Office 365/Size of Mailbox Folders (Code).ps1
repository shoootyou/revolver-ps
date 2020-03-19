Get-MigrationBatch -Identity adupont@s10peru.com | fl *

Get-MailboxFolderStatistics -identity adupont@s10peru.com | Where-Object {$_.FolderPath -like '*Gmail*'} | Select Name,FolderPath,FolderSize,FolderAndSubfolderSize,ItemsInFolderAndSubfolders


Get-MailboxFolderStatistics -identity adupont@s10peru.com | Where-Object {$_.FolderPath -like '*Bandeja*'} | Select Name,FolderPath,FolderSize,FolderAndSubfolderSize,ItemsInFolderAndSubfolders

Get-MailboxFolderStatistics -identity adupont@s10peru.com | 
Where-Object {
($_.FolderAndSubfolderSize -ne '0 B (0 bytes)') -and 
($_.FolderPath -notlike '*Calendario*') -and
($_.FolderPath -notlike '*Principio*')} | 
Select Name,FolderPath,FolderSize,FolderAndSubfolderSize,ItemsInFolderAndSubfolders | 
Export-Csv -Path C:\Users\Rodolfo\Desktop\Storage.csv -Delimiter ";"