# 
# Remove-DuplicateItems.ps1 
# 
# By David Barrett, Microsoft Ltd. 2015. Use at your own risk.  No warranties are given. 
# 
#  DISCLAIMER: 
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND. 
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR 
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL 
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, 
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE 
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION 
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU. 
 
param ( 
    [Parameter(Position=0,Mandatory=$False,HelpMessage="Specifies the mailbox to be accessed")] 
    [ValidateNotNullOrEmpty()] 
    [string]$Mailbox, 
 
    [Parameter(Mandatory=$False,HelpMessage="When specified, the archive mailbox will be accessed (instead of the main mailbox)")] 
    [switch]$Archive, 
         
    [Parameter(Mandatory=$False,HelpMessage="Folder to search - if omitted, the mailbox message root folder is assumed")] 
    [string]$FolderPath, 
 
    [Parameter(Mandatory=$False,HelpMessage="When specified, any subfolders will be processed also")] 
    [switch]$RecurseFolders, 
 
    [Parameter(Mandatory=$False,HelpMessage="When specified, duplicates will be matched anywhere within the mailbox (instead of just within the current folder)")] 
    [switch]$MatchEntireMailbox, 
 
    [Parameter(Mandatory=$False,HelpMessage="If this switch is present, folder path is required and the path points to a public folder")] 
    [switch]$PublicFolders, 
 
    [Parameter(Mandatory=$False,HelpMessage="Means that the items will be hard deleted (normally they are only soft deleted)")] 
    [switch]$HardDelete, 
 
    [Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with EWS")] 
    [System.Management.Automation.PSCredential]$Credentials, 
                 
    [Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")] 
    [string]$Username, 
     
    [Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")] 
    [string]$Password, 
     
    [Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")] 
    [string]$Domain, 
     
    [Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")] 
    [switch]$Impersonate, 
     
    [Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]     
    [string]$EwsUrl, 
     
    [Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]     
    [string]$EWSManagedApiPath = "", 
     
    [Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]     
    [switch]$IgnoreSSLCertificate, 
     
    [Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]     
    [switch]$AllowInsecureRedirection, 
     
    [Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]     
    [string]$LogFile = "", 
     
    [Parameter(Mandatory=$False,HelpMessage="Do not apply any changes, just report what would be updated")]     
    [switch]$WhatIf, 
 
    [Parameter(Mandatory=$False,HelpMessage="If specified, EWS request and responses will be dumped to the standard output")]     
    [switch]$Trace     
) 
 
 
# Define our functions 
 
Function Log([string]$Details, [ConsoleColor]$Colour) 
{ 
    if ($Colour -eq $null) 
    { 
        $Colour = [ConsoleColor]::White 
    } 
    Write-Host $Details -ForegroundColor $Colour 
    if ( $LogFile -eq "" ) { return    } 
    $Details | Out-File $LogFile -Append 
} 
 
Function LoadEWSManagedAPI() 
{ 
    # Find and load the managed API 
     
    if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) ) 
    { 
        if ( Test-Path $EWSManagedApiPath ) 
        { 
            Add-Type -Path $EWSManagedApiPath 
            return $true 
        } 
        Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) -ForegroundColor Yellow 
    } 
     
    $a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) } 
    if (!$a) 
    { 
        $a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) } 
    } 
     
    if ($a)     
    { 
        # Load EWS Managed API 
        Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray 
        Add-Type -Path $a.VersionInfo.FileName 
        return $true 
    } 
    return $false 
} 
 
Function CurrentUserPrimarySmtpAddress() 
{ 
    # Attempt to retrieve the current user's primary SMTP address 
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)" 
    $result = $searcher.FindOne() 
 
    if ($result -ne $null) 
    { 
        $mail = $result.Properties["mail"] 
        return $mail 
    } 
    return $null 
} 
 
Function TrustAllCerts() { 
    <# 
    .SYNOPSIS 
    Set certificate trust policy to trust self-signed certificates (for test servers). 
    #> 
 
    ## Code From http://poshcode.org/624 
    ## Create a compilation environment 
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider 
    $Compiler=$Provider.CreateCompiler() 
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters 
    $Params.GenerateExecutable=$False 
    $Params.GenerateInMemory=$True 
    $Params.IncludeDebugInformation=$False 
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null 
 
    $TASource=@' 
        namespace Local.ToolkitExtensions.Net.CertificatePolicy { 
        public class TrustAll : System.Net.ICertificatePolicy { 
            public TrustAll() 
            {  
            } 
            public bool CheckValidationResult(System.Net.ServicePoint sp, 
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert,  
                                                System.Net.WebRequest req, int problem) 
            { 
                return true; 
            } 
        } 
        } 
'@  
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource) 
    $TAAssembly=$TAResults.CompiledAssembly 
 
    ## We now create an instance of the TrustAll and attach it to the ServicePointManager 
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll") 
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll 
} 
 
function CreateService($targetMailbox) 
{ 
    # Creates and returns an ExchangeService object to be used to access mailboxes 
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1) 
 
    # Set credentials if specified, or use logged on user. 
    if ($Credentials -ne $Null) 
    { 
        Write-Verbose "Applying given credentials" 
        $exchangeService.Credentials = $Credentials.GetNetworkCredential() 
    } 
    elseif ($Username -and $Password) 
    { 
        Write-Verbose "Applying given credentials for $Username" 
        if ($Domain) 
        { 
            $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain) 
        } else { 
            $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password) 
        } 
    } 
    else 
    { 
        Write-Verbose "Using default credentials" 
        $exchangeService.UseDefaultCredentials = $true 
    } 
 
    # Set EWS URL if specified, or use autodiscover if no URL specified. 
    if ($EwsUrl) 
    { 
        $exchangeService.URL = New-Object Uri($EwsUrl) 
    } 
    else 
    { 
        try 
        { 
            Write-Verbose "Performing autodiscover for $targetMailbox" 
            if ( $AllowInsecureRedirection ) 
            { 
                $exchangeService.AutodiscoverUrl($targetMailbox, {$True}) 
            } 
            else 
            { 
                $exchangeService.AutodiscoverUrl($targetMailbox) 
            } 
            if ([string]::IsNullOrEmpty($exchangeService.Url)) 
            { 
                Log "$targetMailbox : autodiscover failed" Red 
                return $Null 
            } 
            Write-Verbose "EWS Url found: $($exchangeService.Url)" 
        } 
        catch 
        { 
            Write-Host "Autodiscover failed: $($Error[0])" -ForegroundColor Red 
            Write-Host "Invalid credentials can cause failure here, even if the error looks generic (e.g$c. service could not be located)" -ForegroundColor Gray 
            exit 
        } 
    } 
  
    if ($Impersonate) 
    { 
        $exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox) 
    } 
 
    if ($Trace) 
    { 
        $exchangeService.TraceEnabled = $True 
    } 
 
    return $exchangeService 
} 
 
function GetFolderPath($Folder) 
{ 
    # Return the full path for the given folder 
 
    # We cache our folder lookups for this script 
    if (!$script:folderCache) 
    { 
        # Note that we can't use a PowerShell hash table to build a list of folder Ids, as the hash table is case-insensitive 
        # We use a .Net Dictionary object instead 
        $script:folderCache = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
    } 
 
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId) 
    $parentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($global:service, $Folder.Id, $propset) 
    $folderPath = $Folder.DisplayName 
    $parentFolderId = $Folder.Id 
    while ($parentFolder.ParentFolderId -ne $parentFolderId) 
    { 
        if ($script:folderCache.ContainsKey($parentFolder.ParentFolderId.UniqueId)) 
        { 
            $parentFolder = $script:folderCache[$parentFolder.ParentFolderId.UniqueId] 
        } 
        else 
        { 
            $parentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($global:service, $parentFolder.ParentFolderId, $propset) 
            $script:FolderCache.Add($parentFolder.Id.UniqueId, $parentFolder) 
        } 
        $folderPath = $parentFolder.DisplayName + "\" + $folderPath 
        $parentFolderId = $parentFolder.Id 
    } 
    return $folderPath 
} 
 
function IsDuplicateAppointment($item) 
{ 
    # Test for duplicate appointment 
    $isDupe = $False 
    if ($script:icaluids.ContainsKey($item.ICalUid)) 
    { 
        # Duplicate ICalUid exists 
        return $True 
    } 
    else 
    { 
        $script:icaluids.Add($item.ICalUid, $item.Id.UniqueId) 
 
        $subject_cmp = $item.Subject 
        if ([String]::IsNullOrEmpty($subject_cmp)) 
        { 
            $subject_cmp = "[No Subject]" # If the subject is blank, we need to give it an arbitrary value to prevent checks failing 
        } 
        if ($script:calsubjects.ContainsKey($subject_cmp)) 
        { 
            # Duplicate subject exists, so we now check the start and end date to confirm if this is a duplicate 
            $dupSubjects = $script:calsubjects[$subject_cmp] 
            foreach ($dupSubject in $dupSubjects) 
            { 
                if (($dupSubject.Start -eq $item.Start) -and ($dupSubject.End -eq $item.End)) 
                { 
                    # Same subject, start, and end date, so this is a duplicate 
                    return $true 
                } 
            } 
            # Add this item to the list of items with the same subject (as it is not a duplicate) 
            $script:calsubjects[$subject_cmp] += $item 
        } 
        else 
        { 
            # Add this to our subject list 
            $script:calsubjects.Add($subject_cmp, @($item)) 
        } 
    } 
    return $false 
} 
 
function IsDuplicateContact($item) 
{ 
    # Test for duplicate contact 
    $item.Load([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
 
    if (![String]::IsNullOrEmpty(($item.DisplayName))) 
    { 
        if (!$script:displayNames.ContainsKey($item.DisplayName)) 
        { 
            # No duplicate contact display name found, so we add this one to our list 
            $script:displayNames.Add($item.DisplayName, @($item)) 
            return $false 
        } 
    } 
    else 
    { 
        # If display name is empty, we do not count this as a duplicate (we ignore it) 
        return $false 
    } 
 
    # We have another contact with same display name, so we now need to check other fields to confirm match 
 
    $possibleMatches = $script:displayNames[$item.DisplayName] 
 
    foreach ($possibleMatch in $possibleMatches) 
    { 
        $match = $true 
        if ($item.EmailAddress1 -ne $possibleMatch.EmailAddress1) { $match = $false } 
        if ($item.EmailAddress2 -ne $possibleMatch.EmailAddress2) { $match = $false } 
        if ($item.EmailAddress3 -ne $possibleMatch.EmailAddress3) { $match = $false } 
        if ($item.ImAddress1 -ne $possibleMatch.ImAddress1) { $match = $false } 
        if ($item.ImAddress2 -ne $possibleMatch.ImAddress2) { $match = $false } 
        if ($item.ImAddress3 -ne $possibleMatch.ImAddress3) { $match = $false } 
        if ($item.BusinessPhone -ne $possibleMatch.BusinessPhone) { $match = $false } 
        if ($item.BusinessPhone2 -ne $possibleMatch.BusinessPhone2) { $match = $false } 
        if ($item.CompanyName -ne $possibleMatch.CompanyName) { $match = $false } 
        if ($item.HomePhone -ne $possibleMatch.HomePhone) { $match = $false } 
        if ($item.HomePhone2 -ne $possibleMatch.HomePhone2) { $match = $false } 
        if ($item.MobilePhone -ne $possibleMatch.MobilePhone) { $match = $false } 
        if ($item.Birthday -ne $possibleMatch.Birthday) { $match = $false } 
 
        if ($match) 
        { 
            Write-Verbose "Duplicate found: $($item.DisplayName)" 
            return $true 
        } 
    } 
 
    # This isn't a duplicate, so we want to add it to our list of possible duplicates with the same display name 
    $script:displayNames[$item.DisplayName] += $item 
    return $false 
} 
 
function IsDuplicateEmail($item) 
{ 
    # Test for duplicate email 
    $isDupe = $False 
 
    if (![String]::IsNullOrEmpty(($item.InternetMessageId))) 
    { 
        if ($script:imsgids.ContainsKey($item.InternetMessageId)) 
        { 
            # Duplicate Internet Message Id exists 
            return $True 
        } 
        $script:imsgids.Add($item.InternetMessageId, $item.Id.UniqueId) 
    } 
 
    $subject_cmp = $item.Subject 
    if ([String]::IsNullOrEmpty($subject_cmp)) 
    { 
        $subject_cmp = "[No Subject]" # If the subject is blank, we need to give it an arbitrary value to prevent checks failing 
    } 
    if ($script:msgsubjects.ContainsKey($subject_cmp)) 
    { 
        # Duplicate subject exists, so we now check the start and end date to confirm if this is a duplicate 
        $dupSubjects = $script:msgsubjects[$subject_cmp] 
        foreach ($dupSubject in $dupSubjects) 
        { 
            if ($item.IsFromMe) 
            { 
                # This is a sent item 
                if (($dupSubject.DateTimeSent -eq $item.DateTimeSent) -and ($dupSubject.IsFromMe)) 
                { 
                    # Same subject and sent date, so this is a duplicate 
                    return $true 
                } 
            } 
            else 
            { 
                # This is a received item 
                if (($dupSubject.DateTimeReceived -eq $item.DateTimeReceived) -and (!$dupSubject.IsFromMe)) 
                { 
                    # Same subject and received date, so this is a duplicate 
                    return $true 
                } 
            } 
        } 
        # Add this item to the list of items with the same subject (as it is not a duplicate) 
        $script:msgsubjects[$subject_cmp] += $item 
    } 
    else 
    { 
        # Add this to our subject list 
        $script:msgsubjects.Add($subject_cmp, @($item)) 
    } 
    return $false 
} 
 
function IsDuplicate($item) 
{ 
    # Test if item is duplicate (the check we do depends upon the item type) 
 
    if ($item.ItemClass.StartsWith("IPM.Note")) 
    { 
        return IsDuplicateEmail($item) 
    } 
    if ($item.ItemClass.Equals("IPM.Appointment")) 
    { 
        return IsDuplicateAppointment($item) 
    } 
    if ($item.ItemClass.Equals("IPM.Contact")) 
    { 
        return IsDuplicateContact($item) 
    } 
    Write-Verbose "Unsupported item type being ignored: $($item.ItemClass)" 
    return $false 
} 
 
function SearchForDuplicates($folder) 
{ 
    # Search the folder for duplicate appointments 
    # We read all the items in the folder, and build a list of all the duplicates 
 
    # First of all, check if we are recursing and process any subfolders first 
    if ($RecurseFolders) 
    { 
        if ($folder.ChildFolderCount -gt 0) 
        { 
            # Deal with any subfolders first 
            Write-Verbose "Processing subfolders of $($folder.DisplayName)" 
            $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000) 
            $FolderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,[Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount) 
            $findFolderResults = $folder.FindFolders($FolderView) 
            ForEach ($subfolder in $findFolderResults.Folders) 
            { 
                SearchForDuplicates $subfolder 
            } 
        } 
    } 
 
    if (!$MatchEntireMailbox -or ($script:calsubjects -eq $Null)) 
    { 
        # Clear/initialise the duplicate tracking lists (we are only checking for duplicates within a folder) 
        $script:calsubjects = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
        $script:msgsubjects = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
        $script:icaluids = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
        $script:imsgids = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
        $script:displayNames = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]' 
    } 
    $dupeCount = 0 
 
    $offset = 0 
    $moreItems = $true 
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(500, 0) 
    $propset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ICalUid) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsFromMe) 
    $propset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass) 
 
    $view.PropertySet = $propset 
 
    while ($moreItems) 
    { 
        $results = $Folder.FindItems($view) 
        $moreItems = $results.MoreAvailable 
        $view.Offset = $results.NextPageOffset 
        foreach ($item in $results) 
        { 
            Write-Verbose "Processing: $($item.Subject)" 
            If (IsDuplicate($item)) 
            { 
                $script:duplicateItems += $item 
                $dupeCount++ 
            } 
        } 
    } 
 
    if ($dupeCount -eq 0) 
    { 
        Log "No duplicate items found in folder $($folder.Displayname)" Green 
        return 
    } 
    Log "$dupeCount duplicates found in folder $($folder.Displayname)" Green 
 
} 
 
Function BatchDeleteDuplicates() 
{ 
    # We now have a list of duplicate items, so we can process them 
 
    Log "$($script:duplicateItems.Count) duplicate items have been found" Green 
    if ( $WhatIf ) 
    { 
        ForEach ($dupe in $script:duplicateItems) 
        { 
            if ([String]::IsNullOrEmpty($dupe.Subject)) 
            { 
                Log "Would delete: [No Subject]" Gray 
            } 
            else 
            { 
                Log ([string]::Format("Would delete: {0}", $dupe.Subject)) Gray 
            } 
        } 
        return 
    } 
 
    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete 
    if ($HardDelete) 
    { 
        $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete 
    } 
 
    # Delete the items (we will do this in batches of 500) 
    $itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx") 
    $itemIdType = [Type] $itemId.GetType() 
    $baseList = [System.Collections.Generic.List``1] 
    $genericItemIdList = $baseList.MakeGenericType(@($itemIdType)) 
    $deleteIds = [Activator]::CreateInstance($genericItemIdList) 
    ForEach ($dupe in $script:duplicateItems) 
    { 
        Log ([string]::Format("Deleting: {0}", $dupe.Subject)) Gray 
        $deleteIds.Add($dupe.Id) 
        if ($deleteIds.Count -ge 500) 
        { 
            # Send the delete request 
            [void]$global:service.DeleteItems( $deleteIds, $deleteMode, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null ) 
            Write-Verbose ([string]::Format("{0} items deleted", $deleteIds.Count)) 
            $deleteIds = [Activator]::CreateInstance($genericItemIdList) 
        } 
    } 
    if ($deleteIds.Count -gt 0) 
    { 
        [void]$global:service.DeleteItems( $deleteIds, $deleteMode, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null ) 
    } 
    Write-Verbose ([string]::Format("{0} items deleted", $deleteIds.Count)) 
} 
 
function GetFolder($FolderPath) 
{ 
    # Return a reference to a folder specified by path 
     
    try 
    { 
        if ($PublicFolders) 
        { 
            $mbx = "" 
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($global:service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot) 
        } 
        else 
        { 
            $mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox ) 
            if ($Archive) 
            { 
                $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $mbx ) 
            } 
            else 
            { 
                $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mbx ) 
            } 
            $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($global:service, $folderId) 
        } 
    } 
    catch 
    { 
        Write-Host "Failed to bind to root folder: $($Error[0])" -ForegroundColor Red 
        Write-Host "This could be due to lack of permissions to the mailbox, or invalid credentials." -ForegroundColor Gray 
        exit 
    } 
 
    if ($FolderPath -ne '\') 
    { 
        $PathElements = $FolderPath -split '\\' 
        For ($i=0; $i -lt $PathElements.Count; $i++) 
        { 
            if ($PathElements[$i]) 
            { 
                $View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0) 
                $View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly 
                         
                $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i]) 
                 
                $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                if ($FolderResults.TotalCount -ne 1) 
                { 
                    # We have either none or more than one folder returned... Either way, we can't continue 
                    $Folder = $null 
                    Write-Host "Failed to find $($PathElements[$i]), path requested was $FolderPath" -ForegroundColor Red 
                    break 
                } 
                 
                if (![String]::IsNullOrEmpty(($mbx))) 
                { 
                    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($FolderResults.Folders[0].Id ) 
                    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId) 
                } 
                else 
                { 
                    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderResults.Folders[0].Id) 
                } 
            } 
        } 
    } 
     
    return $Folder 
} 
 
function ProcessMailbox() 
{ 
    # Process the mailbox 
    Write-Host ([string]::Format("Processing mailbox {0}", $Mailbox)) -ForegroundColor Gray 
    $global:service = CreateService($Mailbox) 
    if ($global:service -eq $Null) 
    { 
        Write-Host "Failed to create ExchangeService" -ForegroundColor Red 
    } 
     
    $Folder = $Null 
    if ([String]::IsNullOrEmpty($FolderPath)) 
    { 
        $FolderPath = "\" 
    } 
    $Folder = GetFolder($FolderPath) 
    if (!$Folder) 
    { 
        Write-Host "Failed to find folder $FolderPath" -ForegroundColor Red 
        return 
    } 
 
    $script:duplicateItems = @() 
    SearchForDuplicates $Folder 
    BatchDeleteDuplicates 
} 
 
 
# The following is the main script 
 
if ( [string]::IsNullOrEmpty($Mailbox) ) 
{ 
    $Mailbox = CurrentUserPrimarySmtpAddress 
    if ( [string]::IsNullOrEmpty($Mailbox) ) 
    { 
        Write-Host "Mailbox not specified.  Failed to determine current user's SMTP address." -ForegroundColor Red 
        Exit 
    } 
    else 
    { 
        Write-Host ([string]::Format("Current user's SMTP address is {0}", $Mailbox)) -ForegroundColor Green 
    } 
} 
 
# Check if we need to ignore any certificate errors 
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!) 
if ($IgnoreSSLCertificate) 
{ 
    Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow 
    TrustAllCerts 
} 
  
# Load EWS Managed API 
if (!(LoadEWSManagedAPI)) 
{ 
    Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red 
    Exit 
} 
 
# Check we have valid credentials 
if ($Credentials -ne $Null) 
{ 
    If ($Username -or $Password) 
    { 
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red 
        Exit 
    } 
} 
 
   
 
Write-Host "" 
 
# Check whether we have a CSV file as input... 
$FileExists = Test-Path $Mailbox 
If ( $FileExists ) 
{ 
    # We have a CSV to process 
    $csv = Import-CSV $Mailbox 
    foreach ($entry in $csv) 
    { 
        $Mailbox = $entry.PrimarySmtpAddress 
        if ( [string]::IsNullOrEmpty($Mailbox) -eq $False ) 
        { 
            ProcessMailbox 
        } 
    } 
} 
Else 
{ 
    # Process as single mailbox 
    ProcessMailbox 
}