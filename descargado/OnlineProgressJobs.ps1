$verbose = $false
#  Author:        Wilson Souza
#  Date Created:  03/18/2016
#  Date Modified: 03/24/2016
#

Remove-Variable * -ErrorAction SilentlyContinue
$scriptversion = '2.41'

#
#  Change Log:
#  
#  2.41 - Added information about WaitForFileCatalog registry entry
#  2.4  - Added Support for Microsoft Azure Recovery Agent running on a DPM Server
#  2.3  - Fixed logic problem when a job is written to a different job file
#  2.2  - Now it works on a DPM server (as long as the MABS Powershell module is copied over.
#  2.1  - Took the 2 seconds delay out of the script as it looks like we were able to manage the memory consumption.
#  2.0  - Fixed duplicated output entry when the entry should be different from the previous one.
#  1.2  - Query job by assgined it to a variable. This is to avoid object creation flooring which causes memory leak
#         Data capture is no longer real time. There is a 2 seconds hold interval before checking if something has changed.
#
#

$ErrorActionPreference='Continue'
cls
 
function header
{
    if ($DPMInstallation)
    {    
        'Last Updated              Elapsed Time   State                  Data Transfer (MB)   File changes    Files Total     Status                 Datasource                                 Production Server                               Protection Group                 JobID/TaskID                           Error'
        '-----------------------   ------------   --------------------   ------------------   -------------   -------------   --------------------   ----------------------------------------   ---------------------------------------------   ------------------------------   ------------------------------------   --------------------'
    }
    else
    {   
        'Last Updated              Elapsed Time   State                  Data Transfer (MB)   File changes    Files Total     Status                 Datasource             JobID/TaskID                           Index   Error'
        '-----------------------   ------------   --------------------   ------------------   -------------   -------------   --------------------   --------------------   ------------------------------------   -----   --------------------'
    }
}

function body ([int]$Idx1, [int]$idx2,[string] $ColorStatus, [string] $JobPrint, [string] $DLS)
{ 
        $HOST.UI.RawUI.ForegroundColor = $ColorStatus
        $datebody   = get-date
        $dateoutput = ("{0:00}" -f $datebody.Month).ToString() +'/' + ("{0:00}" -f $datebody.Day).ToString() + '/' + ("{0:00}" -f $datebody.Year).ToString() + ' ' +  ("{0:00}" -f $datebody.Hour).ToString()  + ':' + ("{0:00}" -f  $datebody.Minute).ToString() + ':' + ("{0:00}" -f $datebody.Second).ToString() + ':' + ("{0:000}" -f  $datebody.Millisecond).ToString()
        if ($DPMInstallation)
        {
            ('{0,23}   {1,-12}   {2,-20}   {3,18:N2}   {4,13:N0}   {5,13:N0}   {6,-20}   {7,-40}   {8,-45}   {9,-30}   {10}   {11}' -f $dateoutput, $elapsedtime, $jobstate[$Idx1], ($byteProgress[$Idx1][$idx2]/1024/1024), $changed[$Idx1][$idx2], $total[$Idx1][$idx2], $jobstatus[$Idx1][$idx2], $DatasourceName[$Idx1][$idx2], $ProductionServer[$Idx1],$ProtectionGroup[$Idx1], $JobPrint, $DLS) 
        }
        else
        {
            ('{0,23}   {1,-12}   {2,-20}   {3,18:N2}   {4,13:N0}   {5,13:N0}   {6,-20}   {7,-20}   {8}   {9,3}   {10}' -f $dateoutput, $elapsedtime, $jobstate[$Idx1], ($byteProgress[$Idx1][$idx2]/1024/1024), $changed[$Idx1][$idx2], $total[$Idx1][$idx2], $jobstatus[$Idx1][$idx2], $DatasourceName[$Idx1][$idx2], $JobPrint, $idx2, $DLS) 
        }
}

function bodytime ([int]$Idx1, [int]$idx2,[string] $ColorStatus)
{
        $HOST.UI.RawUI.ForegroundColor = $ColorStatus
        [console]::SetCursorPosition(26,$Print[$Idx1][$idx2])
        write-host ('{0,-12}' -f $elapsedtime) -ForegroundColor $ColorStatus
}

# Check if we are running from POwerShell console
if ($HOST.name -ne 'ConsoleHost' -and !$verbose)
{
    write-host "Please run this script from PowerShell Console instead of ISE. Exiting..."
    exit
}


# Get current foreground color and setting Window Size
if (!$verbose)
{
    $Foreground   = $HOST.UI.RawUI.ForegroundColor
    $bufferOrig   = $host.UI.RawUI.BufferSize
    $buffer       = $host.UI.RawUI.BufferSize
    $Buffer.width = 350
    $Sizeorig     = $host.UI.RawUI.windowsize
    $size         = $host.UI.RawUI.windowsize
    $size.width   = 120
    $host.UI.RawUI.BufferSize = $buffer
    $host.UI.RawUI.windowsize = $size
}

#Azure/DPM Installation path
if (test-path 'HKLM:\SOFTWARE\Microsoft\Windows Azure Backup\Setup')           { $AzureInstallPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Azure Backup\Setup').InstallPath }
if (test-path 'HKLM:\SOFTWARE\Microsoft\Microsoft Data Protection Manager\DB') { $DPMInstallation = $true } else { $DPMInstallation = $false }

if ($DPMInstallation) { Connect-DPMServer (&hostname) -WarningAction SilentlyContinue | Out-Null }

if (!$AzureInstallPath)
{
    write-host 'Microsoft Azure Recovery Service agent is not installed on this machine. Exiting..' 
    exit
}

#Loading Azure PowerShell Module
& ($AzureInstallPath+'bin\WABModuleInitScript.ps1')

if (!$? -and $DPMInstallation)
{
        Write-Host "As this is a DPM Server, Windows Azure Backup client didn't install MSOnlineBackup PowerShell Module" 
        write-host "Please copy folder " -NoNewline
        write-host "modules " -ForegroundColor Yellow -NoNewline
        write-host "from a Windows Azure Backup server where DPM isn't present to " -NoNewline
        write-host ($AzureInstallPath + 'bin') -ForegroundColor Yellow
        Write-Host "Also copy file " -NoNewline
        write-host "WABModuleInitScript.ps1 " -ForegroundColor Yellow -NoNewline
        Write-Host "(this file is on the BIN folder) to folder " -NoNewline
        Write-Host ($AzureInstallPath + 'bin') -ForegroundColor Yellow
        Write-Host "`nExiting script as this step is required" -ForegroundColor Yellow
        exit   
}

# Get the Current PowerShell process information
$PSProcessInfo = [System.Diagnostics.Process]::GetCurrentProcess() 

#Create folder with the start time of the script (folder will only be created after the first online job shows up)

$date                  = get-date -Format 'MM-dd-yyy_HH-mm'
$ScriptPath            = $AzureInstallPath + 'temp\Online_jobs_in_progress_' + $date + '\' 
$file                  = $ScriptPath + (&hostname) + '_onlinebackup_' 
$job                   = @()
$jobs                  = @()
$jobid                 = @()
$jobstate              = @() 
$ProtectionGroup       = @()
$ProductionServer      = @() 
$AgentVersion          = "Microsoft Azure Recovery Services Agent Version.: " + (dir ($AzureInstallPath  + 'bin\cbengine.exe')).VersionInfo.FileVersion
$MachineID             = "Microsoft Azure MachineID.......................: " + (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Azure Backup\Config').MachineId
$DisableThreadTimeout  = (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows Azure Backup\DbgSettings').DisableThreadTimeout 
$WaitForFileCatalog    = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Azure Backup\Config\CloudBackupProvider').WaitForFileCatalog
$ResourceID            = "Microsoft Azure ResourceID......................: " + (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Azure Backup\Config').ResourceId
$ThreadTimeOut         = "Microsoft Azure DisableThreadTimeout value......: " 
$WaitForFileCatalogmsg = "Microsoft Azure WaitForFileCatalog value........: " 
if ($DPMInstallation) { $DPMVersion = "DPM Version.....................................: " + (dir ((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft Data Protection Manager\Setup').InstallPath + 'bin\msdpm.exe')).VersionInfo.FileVersion }

If ($DisableThreadTimeout -eq $null) { $ThreadTimeOut      = $ThreadTimeOut      + "Null" } else { $ThreadTimeOut = $ThreadTimeOut + $DisableThreadTimeout }
If ($WaitForFileCatalog   -eq $null) { $WaitForFileCatalogmsg = $WaitForFileCatalogmsg + "Null" } else { $WaitForFileCatalogmsg = $WaitForFileCatalogmsg + $WaitForFileCatalog }

write-host "Script Version $scriptversion`n" -f yellow
write-host "Script results will be saved at " $ScriptPath
write-host ("`nInitial Working set (MB): {0,-16:n2}" -f  ($PSProcessInfo.WS/1024/1024))
write-host "Current Working set (MB):"

# Loop waiting for first online job to show up
$joblist = @(Get-OBJob)
[console]::SetCursorPosition(0,7)
Write-Host "Waiting for online job to start. Press CTRL + C to exit" 
while (!$joblist)
{
    [System.GC]::Collect()
    $joblist = @(get-objob)
}
[console]::SetCursorPosition(0,7)
Write-Host "                                                         " 

# Create Folder to hold files
md $ScriptPath | Out-Null

[console]::setcursorposition(0,8)
header
$line = 10

# Main loop
while ($joblist)
{ 
    $joblist = @(get-objob)
    if ($verbose)
    {
        Write-Host "Line 104 - joblist" -ForegroundColor Yellow
        $joblist.jobid.guid
    }

    [console]::setcursorposition(26,5)
    $PSProcessInfo  = [System.Diagnostics.Process]::GetCurrentProcess() 
    Write-Host ("{0,-16:n2}" -f ($PSProcessInfo.WS/1024/1024))
    [System.GC]::Collect()

    foreach ($job in $joblist)
    {
        if ($jobs -notcontains $job)
        {
            $jobs = $jobs + $job
            $a    = [array]::IndexOf($jobs.jobid.guid,$job.jobid.guid) 

            if ($verbose)
            {
                Write-Host 'a: ' $a -ForegroundColor Yellow
            }

            $JobsArray         += ,@()
            $jobstatus         += ,@()    
            $byteProgress      += ,@()
            $changed           += ,@()      
            $total             += ,@() 
            $DatasourceName    += ,@()
            $Print             += ,@() 
            $jobstate          += $job.JobStatus.JobState 
            $count=0  

            if ($verbose)
            {
                write-host 'Line 143 - New Job shows up' $job.jobid.guid
            }


            foreach ($JobSteps in $job.JobStatus.DatasourceStatus)
            {
                [System.GC]::Collect()
                $jobid              += $Job.jobid.guid
                $jobstatus[$a]      += $JobSteps.jobstate
                if ($JobSteps.byteprogress.Progress) { $byteProgress[$a]   += $JobSteps.byteprogress.Progress } else { $byteProgress[$a]   +=0 } 
                $changed[$a]        += $JobSteps.fileprogress.Changed
                $total[$a]          += $JobSteps.fileprogress.total
                if ($DPMInstallation)
                {
                    $DPMOnlineJob        = (Get-DPMJob -Status "InProgress") | ? { $_.tasks.taskid.guid -eq $Job.jobid.guid }
                    $DatasourceName[$a] += $DPMOnlineJob.DataSources
                    $ProtectionGroup    += $DPMOnlineJob.ProtectionGroupName
                    $ProductionServer   += $DPMOnlineJob.tasks.productionservername
                }
                else
                {                    
                    $DatasourceName[$a] += $JobSteps.Datasource.DataSourceName
                }
                $PSProcessInfo       = [System.Diagnostics.Process]::GetCurrentProcess() 
                $print[$a]          += $line

                "Script Version..................................: " + $scriptversion                                                | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                $AgentVersion                                                                                                        | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                $MachineID                                                                                                           | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                $WaitForFileCatalogmsg                                                                                               | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                $ThreadTimeOut                                                                                                       | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                $ResourceID                                                                                                          | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                "Online JobId (DPM TaskID).......................: " + $job.jobid.guid                                               | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                "DatasourceID....................................: " + $job.JobStatus.DatasourceStatus.datasource.DataSourceid       | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                "Azure DataSource Name...........................: " + $JobSteps.Datasource.DataSourceName                           | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                if ($DPMInstallation)
                {
                    $DPMVersion                                                                                                      | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                    "DPM DataSource Name.............................: " + $DatasourceName[$a][$count]                               | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                }
                "PoweShell Process ID............................: " + $PSProcessInfo.Id                                             | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                ""                                                                                                                   | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                header | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append

                $date = get-date
                $elapsedtime = ("{0:00}" -f ($date - $job.JobStatus.StartTime.ToLocalTime()).Days).ToString() + 'd' + ("{0:00}" -f ($date - $job.JobStatus.StartTime.ToLocalTime()).Hours).ToString() + 'h' + ( "{0:00}" -f ($date - $job.JobStatus.StartTime.ToLocalTime()).Minutes).ToString() + 'm' + ( "{0:00}"-f ($date - $job.JobStatus.StartTime.ToLocalTime()).seconds).ToString() + 's'
                [console]::SetCursorPosition(0,$print[$a][$count])
                if ($jobstatus[$a] -eq 'Aborted') { $color = 'Yellow'} else { $color = 'White' }
                if ($verbose)
                {
                    write-host "line 172 - Output new entry to a file" -ForegroundColor Yellow
                }


                body $a $count $color $job.JobId.Guid 
                body $a $count $color $job.JobId.Guid | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append

                $count++
                $line++
            }
        }


        $a = [array]::IndexOf($jobs.jobid.guid,$job.jobid.guid) 
        $count=0

        if ($verbose)
        {
            Write-Host "Line 194 - Start Jobsteps loop"  $job.JobStatus.DatasourceStatus.Count -ForegroundColor Yellow
            write-host "A:" $a -ForegroundColor Yellow
            Write-Host "JobID:"  $job.jobid.guid -ForegroundColor Yellow
        }

        foreach ($JobSteps in $job.JobStatus.DatasourceStatus)
        {
            [System.GC]::Collect()
            $date = get-date
            $elapsedtime = ("{0:00}" -f ($date - $job.JobStatus.StartTime.ToLocalTime()).Days).ToString() + 'd' + ("{0:00}" -f ($date - $job.JobStatus.StartTime.ToLocalTime()).Hours).ToString() + 'h' + ( "{0:00}" -f ($date - $job.JobStatus.StartTime.ToLocalTime()).Minutes).ToString() + 'm' + ( "{0:00}"-f ($date - $job.JobStatus.StartTime.ToLocalTime()).seconds).ToString() + 's'
            [console]::SetCursorPosition(0,$print[$a][$count])

            if (
                    $Job.JobStatus.JobState         -ne $jobstate[$a]             -or 
                    $JobSteps.byteprogress.Progress -ne $byteProgress[$a][$count] -or 
                    $JobSteps.fileprogress.Changed  -ne $Changed[$a][$count]      -or 
                    $JobSteps.jobstate              -ne $jobstatus[$a][$count]    -or 
                    $JobSteps.fileprogress.total    -ne $total[$a][$count]  
               )
            {
                if ($verbose)
                {
                    write-host "Line 215 - Found change... will write the change to a file"        
                }
                $jobstate[$a]             = $job.JobStatus.JobState        
                $byteProgress[$a][$count] = $JobSteps.byteprogress.Progress  
                $changed[$a][$count]      = $JobSteps.fileprogress.Changed        
                $jobstatus[$a][$count]    = $JobSteps.jobstate      
                $total[$a][$count]        = $JobSteps.fileprogress.total                

                if ($jobstatus[$a] -eq 'Aborted') { $color = 'Yellow'} else { $color = 'White' }
                body $a $count $color $job.JobId.Guid 
                body $a $count $color $job.JobId.Guid | Out-File ($file + $job.jobid + '_' + $count + '.txt') -Append
                break
            }

            if ($jobstatus[$a] -eq 'Aborted') { $color = 'Yellow'} else { $color = 'White' }
            if ($verbose) {  write-host "Line 231 - Exit loop" -ForegroundColor Yellow }
            bodyTime $a $count $color
            $HOST.UI.RawUI.ForegroundColor = 'white'
            $count++
        }
    }
    foreach ($jobupdate in $jobs)
     {
         [System.GC]::Collect()
         if ($joblist.jobid.guid -notcontains $jobupdate.jobid.guid)
         {
             $jobresult = $Null
             $count = 0
             while (!$jobresult)
             {
                 $Jobresult = Get-OBJob -Previous 100000000 | ? { $_.jobid.guid -eq $jobupdate.jobid.guid }
                 $count++
                 if ($count -gt 100) { [console]::SetCursorPosition(0,$line++); write-host "Can't find job. This happens if CBENGINE was stopped or crashed"; exit }

             }
             if ($verbose) { write-host "Line 245 - job result:" $Jobresult.JobId.guid -ForegroundColor Yellow }

             $countB = 0
             Remove-Variable b -ErrorAction SilentlyContinue
             $b = [array]::IndexOf($jobs.jobid.guid,$jobresult.jobid.guid)
             if ($verbose) { if ($b) { "has value" } else { "doesn't have value" } }
             foreach ($JobSteps in $Jobresult.JobStatus.DatasourceStatus)
             {
                 [System.GC]::Collect()
                 if ($DatasourceName[$b][$countB] -ne 'BackupCompleted')
                 {                    
                    if ($verbose) { write-host "datasource name: "  ($DatasourceName[$b][$countB]) -ForegroundColor Yellow }
                    $jobstate[$b]                = $Jobresult.JobStatus.JobState        
                    $byteProgress[$b][$countB]   = $JobSteps.byteprogress.Progress  
                    $changed[$b][$countB]        = $JobSteps.fileprogress.Changed        
                    $jobstatus[$b][$countB]      = $JobSteps.jobstate      
                    $total[$b][$countB]          = $JobSteps.fileprogress.total  
                    if (!$DPMInstallation) { $DatasourceName[$b][$countB] = $JobSteps.Datasource.DataSourceName }

                    $date = get-date                    
                    $elapsedtime = ("{0:00}" -f ($date - $Jobresult.JobStatus.StartTime.ToLocalTime()).Days).ToString() + 'd' + ("{0:00}" -f ($date - $Jobresult.JobStatus.StartTime.ToLocalTime()).Hours).ToString() + 'h' + ( "{0:00}" -f ($date - $Jobresult.JobStatus.StartTime.ToLocalTime()).Minutes).ToString() + 'm' + ( "{0:00}"-f ($date - $Jobresult.JobStatus.StartTime.ToLocalTime()).seconds).ToString() + 's'

                    [console]::SetCursorPosition(0,$print[$b][$countB])
                    if ($jobstate[$b] -eq 'Aborted') { $color = 'Red'} else { $color = 'Green' }
                    if ($verbose) { write-host 'jobupdate loop'  -ForegroundColor Yellow }
                    $DLSError = $JobSteps.errorinfo.ErrorParamList.name + ' ' + $JobSteps.errorinfo.ErrorParamList.value
                    body $b $countb $color $jobresult.jobid.guid $DLSError
                    body $b $countb $color $jobresult.jobid.guid $DLSError | Out-File ($file + $jobresult.jobid.guid  + '_' + $countb + '.txt') -Append
                    $DatasourceName[$b][$countB] = 'BackupCompleted'
                }
                $countB++
             }
        }
    }
}
# Resetting values
[System.GC]::Collect()
[console]::SetCursorPosition(0,$line++)
$HOST.UI.RawUI.ForegroundColor = $Foreground 
#$host.UI.RawUI.BufferSize      = $bufferOrig
$host.UI.RawUI.windowsize      = $Sizeorig     
