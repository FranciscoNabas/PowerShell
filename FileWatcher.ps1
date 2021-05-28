
<#PSScriptInfo

.VERSION 1.0

.GUID 74bfbe4f-e53b-41b3-9819-360a585b68ae

.AUTHOR francisconabas@outlook.com

.COMPANYNAME

.COPYRIGHT

.TAGS File Copy File Watcher

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
FileWatcher.ps1 is composed of the following features:
Logging compatible with CMTrace.
Invoke-FileCopy function. with MD5 hash verification, copying only unique files.
New-FileWatcher function. Creates a File Watcher and register it for monitoring only or with file copy.
Invoke-LogCleaner function. Manages the logs generated by the solution.

.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Solution uses FileSystemWatch class to monitor a directory and copy files when detects a change, if desired. 

#> 
[CmdletBinding(DefaultParameterSetName = 'Monitor')]
param (

    [Parameter  (Mandatory = $false, ParameterSetName = 'Monitor', HelpMessage = 'Log file path location. Without log file name.')]
    [Parameter  (Mandatory = $false, ParameterSetName = 'InvokeCopy', HelpMessage = 'Log file path location. Without log file name.')]
    [ValidateNotNullOrEmpty()]
    [string]    $LogFilePath = "$Env:windir\Logs",

    [Parameter  (Mandatory = $false, ParameterSetName = 'Monitor', HelpMessage = 'File Watcher events to monitor.')]
    [Parameter  (Mandatory = $false, ParameterSetName = 'InvokeCopy', HelpMessage = 'File Watcher events to monitor.')]
    [ValidateNotNullOrEmpty()]
    [array]     $WatcherEvents = @('Changed','Created','Deleted','Disposed','Error','Renamed'),

    [Parameter  (Mandatory = $true, ParameterSetName = 'Monitor', HelpMessage = 'Path to monitor.')]
    [Parameter  (Mandatory = $true, ParameterSetName = 'InvokeCopy', HelpMessage = 'Path to monitor.')]
    [ValidateNotNullOrEmpty()]
    [string]    $Path,

    [Parameter  (Mandatory = $false, ParameterSetName = 'Monitor', HelpMessage = 'File Watcher filter. Default is *')]
    [Parameter  (Mandatory = $false, ParameterSetName = 'InvokeCopy', HelpMessage = 'File Watcher filter. Default is *')]
    [ValidateNotNullOrEmpty()]
    [string]    $Filter,

    [Parameter  (Mandatory = $false, ParameterSetName = 'Monitor', HelpMessage = 'Destination for the file copy. Designed to be used with -Copy switch.')]
    [Parameter  (Mandatory = $true, ParameterSetName = 'InvokeCopy', HelpMessage = 'Destination for the file copy. Designed to be used with -Copy switch.')]
    [ValidateNotNullOrEmpty()]
    [string]    $CopyDestination,

    [Parameter  (Mandatory = $false, ParameterSetName = 'InvokeCopy', HelpMessage = 'Triggers file copy when any of the File Watcher events occour.')]
    [switch]    $Copy

)

#region Functions
Function Global:Add-Log {

    param (
        
        [Parameter  (Mandatory = $true)]
        [string]    $LogValue,
        
        [Parameter  (Mandatory = $true)]
        [ValidateSet("Info", "Warning", "Error")]
        [string]    $Type,
        
        [Parameter  (Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]    $Component

    )
    switch ($Type) {
        "Info" {
            [int]$Type = 1
        }
        "Warning" {
            [int]$Type = 2
        }
        "Error" {
            [int]$Type = 3
        }
    }    
    $Source = $MyInvocation.MyCommand.Name
    $Content =  "<![LOG[$LogValue]LOG]!>" +`
                "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " +`
                "date=`"$(Get-Date -Format "M-d-yyyy")`" " +`
                "component=`"$Component`" " +`
                "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
                "type=`"$Type`" " +`
                "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
                "file=`"$Source`">"
    if (!(Test-Path $LogFilePath -ErrorAction SilentlyContinue)) {
        mkdir $LogFilePath | Out-Null
    }
    try {
        Add-Content -Path "$LogFilePath\CustomFileWatcher.log" -Value $Content -Force -ErrorAction Stop
    }
    catch {
        Start-Sleep -Milliseconds 700
        Add-Content -Path "$LogFilePath\CustomFileWatcher.log" -Value $Content -Force
    }

}

function Global:Invoke-FileCopy {

    [CmdletBinding()]
    param (

        [Parameter  (Mandatory = $false)]
        [ValidateRange([int]30, [int]::MaxValue)]
        [int]       $DaysToDelete = 30,
        
        [Parameter  (Mandatory = $true, Position = 0)]
        [string]    $Source,

        [Parameter  (Mandatory = $true, Position = 1)]
        [string]    $Destination

    )   

    #region DirectoryCheck
    Write-Verbose "File copy triggered. Source: $Source. Dest: $Destination. $(Get-Date)."
    Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "File copy triggered. Source: $Source. Dest: $Destination."
    if (!(Test-Path -Path $Destination)) {
        mkdir $Destination -Force | Out-Null
    }
    try {
        $Files = Get-ChildItem $Source -Filter *.* -Recurse -Force -ErrorAction Stop
    }
    catch {
        Add-Log -Type 'Error' -Component 'FileCopy' -LogValue "Error fetching files. $($_.Exception.Message)"
        return Write-Error "Error fetching files. $($_.Exception.Message)"
    }
    
    if ($Files) {
        $ExtMgt = $Files | Group-Object -Property Extension -NoElement
        Write-Verbose "Directory Found. $($Files.Count) Files found."
        Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "Directory Found. $($Files.Count) Files found."
        foreach ($Ext in $ExtMgt) {
            Write-Verbose "$($Ext | Select-Object -ExpandProperty Count) Files with extension $($Ext.Name)."
            Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "$($Ext.Count) Files with extension $($Ext.Name)."
        }
    }
    else {
        Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "No files found on the given directory."
        return Write-Warning "No files found on the given directory. $(Get-Date)."
    }
    #endregion  
    #region CleanOldFiles
    Write-Verbose "Cleaning files older than $DaysToDelete days. $(Get-Date)."
    Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "Cleaning files older than $DaysToDelete days."
    $OCopiedFiles = Get-ChildItem $Destination -Filter *.* -Recurse -Force -ErrorAction SilentlyContinue
    if ($OCopiedFiles) {
        foreach ($File in $OCopiedFiles) {
            if (((Get-Date) - $File.CreationTime).Days -ge $DaysToDelete) {
                try {
                    Remove-Item $File.FullName -Force -ErrorAction Stop
                }
                catch {
                    Write-Warning "Error removing file $($File.Name). $($_.Exception.Message) $(Get-Date)."
                    Add-Log -Type 'Warning' -Component 'FileCopy' -LogValue "Error removing file $($File.Name). $($_.Exception.Message)"
                }
            }
        }
    }
    #endregion  
    #region CopyFiles
    foreach ($File in $Files) {
        if ($OCopiedFiles) {
            if ($OCopiedFiles.Name -contains $File.Name) {
                Write-Verbose "Found file with name $($File.Name) on destination. Checking hash. $(Get-Date)."
                Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "Found file with name $($File.Name) on destination. Checking hash."
                $SameName = $OCopiedFiles | Where-Object {$_.Name -eq $File.Name}
                if ((Get-FileHash $File.FullName).Hash -eq (Get-FileHash $SameName.FullName).Hash) {
                    Write-Verbose "Both files with same MD5 hash. Skipping copy. $(Get-Date)."
                    Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "Both files with same MD5 hash. Skipping copy."
                }
                else {
                    Write-Verbose "Files with different MD5 hash. Copying with new name. $(Get-Date)."
                    Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "Files with different MD5 hash. Copying with new name."
                    try {
                        Copy-Item -Path $File.FullName -Destination "$Destination\$($File.Name)_$(Get-Date -Format 'MM-dd-yyyy_hh.mm.ss').$($File.Extension)" -Force -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Error copying file $($File.Name). $($_.Exception.Message) $(Get-Date)."
                        Add-Log -Type 'Warning' -Component 'FileCopy' -LogValue "Error copying file $($File.Name). $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-Verbose "File $($File.Name) not found on destination. Copying. $(Get-Date)."
                Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "File $($File.Name) not found on destination. Copying."
                try {
                    Copy-Item -Path $File.FullName -Destination $Destination -Force -ErrorAction Stop
                }
                catch {
                    Write-Warning "Error copying file $($File.Name). $($_.Exception.Message) $(Get-Date)."
                    Add-Log -Type 'Warning' -Component 'FileCopy' -LogValue "Error copying file $($File.Name). $($_.Exception.Message)"
                }
            }
        }
        else {
            Write-Verbose "No files on destination. Copying. $(Get-Date)."
            Add-Log -Type 'Info' -Component 'FileCopy' -LogValue "No files on destination. Copying."
            try {
                Copy-Item -Path $File.FullName -Destination $Destination -Force -ErrorAction Stop
            }
            catch {
                Write-Warning "Error copying file $($File.Name). $($_.Exception.Message) $(Get-Date)."
                Add-Log -Type 'Warning' -Component 'FileCopy' -LogValue "Error copying file $($File.Name). $($_.Exception.Message)"
            }
        }
    }
    #endregion
}

function New-FileWatcher {
    
    [cmdletbinding()]
    param (

        [Parameter  (Mandatory = $true)]
        [string]    $MonitorPath,

        [Parameter  (Mandatory = $false)]
        [string]    $MonitorFilter

    )

    try {
        $FullPathName = (Get-Item $MonitorPath -Force -ErrorAction Stop).FullName
    }
    catch {
        Add-Log -Type 'Error' -Component 'New-FileWatcher' -LogValue "Error trying to find path specified. $($_.Exception.Message)"
        return Write-Error "Error trying to find path specified. $($_.Exception.Message) $(Get-Date).)"
    }
    try {
        $FileWatcher = New-Object System.IO.FileSystemWatcher
        $FileWatcher.Path = $FullPathName
        if ($Filter) {
            $FileWatcher.Filter = $MonitorFilter
        }
        $FileWatcher.IncludeSubdirectories = $true
        $FileWatcher.EnableRaisingEvents = $true
        Add-Log -Type 'Info' -Component 'New-FileWatcher' -LogValue "File Watcher created for pass: $FullPathName."
        Write-Verbose "File Watcher created for path: $FullPathName. $(Get-Date)."
        
        if ($Copy) {
            $Action = {
                $Path = $Event.SourceEventArgs.FullPath
                $Type = $Event.SourceEventArgs.ChangeType
                $SourcePathName = $Event.MessageData.FullPathName
                $DestinationPathName = $Event.MessageData.CopyDestination
                Add-Log -Type 'Info' -Component 'FileWatch' -LogValue "$Type detected on $Path. Triggering file copy."
                Write-Host "$Type detected on $Path. Triggering file copy. $(Get-Date)." -ForegroundColor DarkCyan
                Invoke-FileCopy -Source $SourcePathName -Destination $DestinationPathName -Verbose
            }    
        }
        else {
            $Action = {
                $Path = $Event.SourceEventArgs.FullPath
                $Type = $Event.SourceEventArgs.ChangeType
                Add-Log -Type 'Info' -Component 'FileWatch' -LogValue "$Type detected on $Path. 'Copy' switch not called. Monitoring only."
                Write-Host "$Type detected on $Path. 'Copy' switch not called. Monitoring only. $(Get-Date)." -ForegroundColor DarkCyan
            }
        }
        
        $MessageObject = New-Object PsObject -Property @{LogFilePath = $LogFilePath; FullPathName = $FullPathName; CopyDestination = $CopyDestination}
        foreach ($WEvent in $WatcherEvents) {
            Register-ObjectEvent $FileWatcher "$WEvent" -Action $Action -MessageData $MessageObject | Out-Null
            Add-Log -Type 'Info' -Component 'New-FileWatcher' -LogValue "Object event registered for '$WEvent'."
            Write-Verbose "Object event registered for '$WEvent'. $(Get-Date)."
        }
    }
    catch {
        Add-Log -Type 'Error' -Component 'New-FileWatcher' -LogValue "Unable to set File Watcher. $($_.Exception.Message)"
        return Write-Error "Unable to set File Watcher. $($_.Exception.Message) $(Get-Date)."
    }
}

function Invoke-LogCleaner {
    
    param (

        [Parameter  (Mandatory = $true)]
        [string]    $CLPath

    )

    Write-Verbose 'Cleaning logfiles older than 7 days.'
    Add-Log -Type 'Info' -Component 'CleaningLogs' -LogValue 'Cleaning logfiles older than 7 days.'
    $LogCTime = (Get-ChildItem $CLPath -ErrorAction SilentlyContinue).CreationTime
    if ($LogCTime) {
        $LogFCreated = Get-Date($LogCTime)
        $DateTime = Get-Date
        if (($DateTime - $LogFCreated).Days -ge 7) {
            try {
                Remove-Item $CLPath -Force -ErrorAction Stop
            }
            catch {
                Write-Warning 'Error removing old logfile. Manual intervention needed.'
                Add-Log -Type "Warning' -Component 'CleaningLogs' -LogValue 'Error removing old logfile. Manual intervention needed. $($_.Exception.Message)."
            }
        }
        else {
            Write-Verbose 'Logfile newer than 7 days.'
            Add-Log -Type 'Info' -Component 'CleaningLogs' -LogValue 'Logfile newer than 7 days.'
        }
    }
    else {
        Write-Verbose 'Logfile not found.'
        Add-Log -Type 'Warning' -Component 'CleaningLogs' -LogValue 'Logfile not found.'
    }
    
}
#endregion

if (!$Filter) {
    New-FileWatcher -MonitorPath $Path -Verbose
}
else {
    New-FileWatcher -MonitorPath $Path -MonitorFilter $Filter -Verbose
}
Invoke-LogCleaner -CLPath "$LogFilePath\CustomFileWatcher.log"
Invoke-FileCopy -Source (Get-Item $Path -Force).FullName -Destination $CopyDestination -Verbose
while ($true) {
    Start-Sleep -Seconds 5
}