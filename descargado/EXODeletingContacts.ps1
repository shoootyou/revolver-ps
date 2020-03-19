<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#>


#requries -Version 2.0

<#
 	.SYNOPSIS
        This script is used to delete all contacts of a mailbox in Exchange Online.
    .DESCRIPTION
        This script is used to delete all contacts of a mailbox in Exchange Online.
    .PARAMETER  Credential
        This parameter indicates the credential to use for connecting to Office 365 Service. 
    .PARAMETER  SmtpAddress
        This parameter spcified which mailbox will be used to delete all  contacts.       
    .EXAMPLE
        EXODeletingContacts.ps1-Credential $Credential -smtpaddress “tom@tenant.com”
        Delete all the contacts of mailbox tom@tenant.com 
#>


Param
(
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory=$true)]
    [String] $SmtpAddress
)

Begin
{
    Add-Type -AssemblyName System.Core

    #Confirm the installation of EWS API
    $webSvcInstallDirRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0" -PSProperty "Install Directory" -ErrorAction:SilentlyContinue
    if ($webSvcInstallDirRegKey -ne $null) 
    {
        $moduleFilePath = $webSvcInstallDirRegKey.'Install Directory' + 'Microsoft.Exchange.WebServices.dll'
        Import-Module $moduleFilePath
    } 
    else 
    {
        $errorMsg = "Please install Exchange Web Service Managed API 2.0"
        throw $errorMsg
    }
	
    #Establish the connection to Exchange Web Service

    $verboseMsg = $Messages.EstablishConnection
    $PSCmdlet.WriteVerbose($verboseMsg)
    $exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
				    [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
			
    #Set network credential
    $userName = $Credential.UserName
    $exService.Credentials = $Credential.GetNetworkCredential()
            

    Try
    {
	    #Set the URL by using Autodiscover
	    $exService.AutodiscoverUrl($userName,{$true})
	    $verboseMsg = $Messages.SaveExWebSvcVariable
	    $PSCmdlet.WriteVerbose($verboseMsg)
	    Set-Variable -Name exService -Value $exService -Scope Global -Force
    }
    Catch [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRemoteException]
    {
	    $PSCmdlet.ThrowTerminatingError($_)
    }
    Catch
    {
	    $PSCmdlet.ThrowTerminatingError($_)
    }
}

Process
{
    $exService.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$smtpAddress);
    $ExportCollection = @()
    Write-Host "Process Contacts"
    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$SmtpAddress)   
    $Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$folderid)
    $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
 
    #Define ItemView to retrive just 1000 Items    
    $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    $deletemode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete     
    $fiItems = $null    
    do{    
        $fiItems = $exservice.FindItems($Contacts.Id,$ivItemView) 
        [Void]$exservice.LoadPropertiesForItems($fiItems,$psPropset)  
        foreach($Item in $fiItems.Items)
        {     
		    if($Item -is [Microsoft.Exchange.WebServices.Data.Contact])
            {
			      $Item.Delete($deletemode)
            }
        }    
        $ivItemView.Offset += $fiItems.Items.Count    
    }while($fiItems.MoreAvailable -eq $true) 
}

End
{
     
}