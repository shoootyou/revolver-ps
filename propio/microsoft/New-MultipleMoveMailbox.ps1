$DB_MOV = Import-Csv "C:\Users\rcastelo\Downloads\migration.csv" | ? { $_.'Mailbox Class Deseada' -eq "CLASS C"} | Select -First 30
foreach($MBX in $DB_MOV){
        New-MoveRequest -Identity $MBX.'Primary Email Address' -TargetDatabase "CNVEXMDBCC01" -BatchName $MBX.'Primary Email Address' -ArchiveTargetDatabase "CNVEXMDBCC01" -SuspendWhenReadyToComplete -AllowLargeItems -BadItemLimit 100 -Priority Highest 
}