$TotalBytes = ($blobCopy | Get-AzureStorageBlobCopyState).TotalBytes
 
cls
 
while(($blobCopy | Get-AzureStorageBlobCopyState).Status -eq "Pending")
 
{
 
Start-Sleep 1
 
$BytesCopied = ($blobCopy | Get-AzureStorageBlobCopyState).BytesCopied
 
$PercentCopied = [math]::Round($BytesCopied/$TotalBytes * 100,2)
 
Write-Progress -Activity "Blob Copy in Progress" -Status "$PercentCopied% Complete:" -PercentComplete $PercentCopied
 
}