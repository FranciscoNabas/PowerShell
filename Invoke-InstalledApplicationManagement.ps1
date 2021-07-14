
<#PSScriptInfo

.VERSION 1.1.0

.GUID b7544f87-8485-445a-9d05-d50cb36c3e67

.AUTHOR Francisco Nabas

.COMPANYNAME 

.COPYRIGHT (c) 2021. All rights reserved.

.TAGS 

.LICENSEURI https://github.com/FranciscoNabas/PowerShell/blob/main/LICENSE

.PROJECTURI https://github.com/FranciscoNabas/PowerShell/blob/main/Invoke-InstalledApplicationManagement.ps1

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#> 



<# 

.DESCRIPTION 
- Check for apllications installed on the machine with the input Name.
- If a match is found on the registry, creates an object with its Name, Version and UninstallString.
- If no match is found on the registry, the Win32_Product CIM/WMI class is queried and an object is created with the app Name, Version and CimInstance.
- If the Uninstall switch is called, check if the installed version is less than input version and uninstall the application(s).
- If the ForceUninstall switch is called, uninstall the application(s) without checking the version.

#> 

Param()
function Invoke-InstalledApplicationManagement {

    <# 

    .SYNOPSIS

        This solution was designed to identify applications installed on the machine and remove it if required.

    .PARAMETER Name

        Name of the application to manage.
        The input will be set between wildcards.

    .PARAMETER Version

        Application current version. Versions older than this will be consider superseded.

    .PARAMETER Uninstall

        If called will uninstall superseded versions found.

    .PARAMETER ForceUninstall

        Caled with the Uninstall switch. Skips the version check and uninstall the application(s) regardless of the version.

    .PARAMETER MsiParameters

        MSI parameters and switches to be used on the uninstallation.
        Used ONLY when the object is found on the registry AND the UninstallString uses MsiExec.exe. In any other case it's ignored.

    .EXAMPLE

        Invoke-InstalledApplicationManagement -Name 'ApplicationName' -Version '1.1.0'
        Invoke-InstalledApplicationManagement 'ApplicationName' '1.1.0'
        Invoke-InstalledApplicationManagement 'ApplicationName' '1.1.0' -Uninstall -MsiParameters '/qn /norestart'
        Invoke-InstalledApplicationManagement 'ApplicationName' -Uninstall -MsiParameters '/qn /norestart' -ForceInstall

    #>

    #Requires -RunAsAdministrator

    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'UseVersion')]
    param (

        [Parameter  (Mandatory, Position = 0, ParameterSetName = 'UseVersion')]
        [Parameter  (Mandatory, Position = 0, ParameterSetName = 'ForceUninstall')]
        [Parameter  (Mandatory, Position = 0, ParameterSetName = 'UninstallWVersion')]
        [ValidateNotNullOrEmpty()]
        [string]    $Name,

        [Parameter  (Mandatory, Position = 1, ParameterSetName = 'UseVersion')]
        [Parameter  (Mandatory, Position = 1, ParameterSetName = 'UninstallWVersion')]
        [ValidateNotNullOrEmpty()]
        [string]    $Version,

        [Parameter  (Mandatory, ParameterSetName = 'UninstallWVersion')]
        [Parameter  (Mandatory, ParameterSetName = 'ForceUninstall')]
        [switch]    $Uninstall,

        [Parameter  (Mandatory, ParameterSetName = 'ForceUninstall')]
        [switch]    $ForceUninstall,

        [Parameter  (ParameterSetName = 'ForceUninstall')]
        [Parameter  (ParameterSetName = 'UninstallWVersion')]
        [ValidateNotNullOrEmpty()]
        [string]    $MsiParameters

    )
    
    begin {

        #region Functions
        function Add-Log {
        
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

            if (!(Test-Path "$Env:windir\Logs\InstalledApplicationManagement" -PathType Container)) {
                $null = mkdir "$Env:windir\Logs\InstalledApplicationManagement"
            }
            try {
                Add-Content -Path "$Env:windir\Logs\InstalledApplicationManagement\InstalledApplicationManagement-$Name-$Env:COMPUTERNAME.log" -Value $Content -Force -ErrorAction Stop
            }
            catch {
                Start-Sleep -Milliseconds 700
                Add-Content -Path "$Env:windir\Logs\InstalledApplicationManagement\InstalledApplicationManagement-$Name-$Env:COMPUTERNAME.log" -Value $Content -Force
            }
        }

        function Get-InstalledApplicationManagement {
            
            [CmdletBinding()]
            param (

                [Parameter (Mandatory)]
                [string]   $AppName,

                [Parameter  ()]
                [string]    $uninstallParameters

            )

            #region InitialVariablePayload
            Write-Verbose "Searching for installed applications on the machine with name: '$AppName'. $(Get-Date)."
            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Searching for installed applications on the machine with name: $AppName."
            $RegObjects = $null
            $AppProperties = @()
            $RegPath = @(
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
            )
            $RegObjects = $RegPath | Get-ChildItem | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"} ## Looking for the application on the registry.
            #endregion

            #region UsingRegistry
            if ($RegObjects) {
                foreach ($Object in $RegObjects) {
                    if ($Object.UninstallString -like 'MsiExec*') { ## If UninstallString uses 'MsiExec.exe', we parse it and add the MsiParameters
                        $parameters = "$uninstallParameters /l*vx+! ""%windir%\Logs\InstalledApplicationManagement\[MSI]$($Object.DisplayName)-$($Object.DisplayVersion)-Uninstall.log"""
                        $UninstallString = ($Object.UninstallString).Replace('/I', '/X')
                        $UninstallString = $UninstallString -replace '$', " $parameters"
                        Write-Verbose "Adding application '$($Object.DisplayName)' version '$($Object.DisplayVersion)' to the control. $(Get-Date)"
                        Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application '$($Object.DisplayName)' version '$($Object.DisplayVersion)' to the control."
                        $AppProperties += [PSCustomObject]@{
                            Name = $Object.DisplayName
                            Version = $Object.DisplayVersion
                            UninstallString = $UninstallString
                            UninstallMethod = 'Registry'
                        }
                    }
                    else { ## If UninstallString don't use MsiExec.exe (probably uses the app .exe), we just Invoke it,
                        Write-Verbose "Adding application '$($Object.DisplayName)' version '$($Object.DisplayVersion)' to the control. $(Get-Date)"
                        Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application '$($Object.DisplayName)' version '$($Object.DisplayVersion)' to the control."
                        $AppProperties += [PSCustomObject]@{
                            Name = $Object.DisplayName
                            Version = $Object.DisplayVersion
                            UninstallString = $Object.UninstallString
                            UninstallMethod = 'Registry'
                        }
                    }
                }
            }
            #endregion
    
            #region UsginCimInstance
            else { ## If no registry keys were found with the provided Name we query Win32_Product CIM\WMI class.
                if ($Verbose) { $params = @{ Verbose = $true; Query = "Select * From Win32_Product Where Name Like '%$AppName%'" } } ## Preparing the parameters for splatting on Invoke-CimInstance.
                else { $params = @{ Query = "Select * From Win32_Product Where Name Like '%$AppName%'" } }
                $cimInstance = Get-CimInstance @params
                if ($cimInstance) {
                    foreach ($product in $cimInstance) {
                        Write-Verbose "Adding application '$($Object.DisplayName)' version '$($Object.DisplayVersion)' to the control. $(Get-Date)"
                        Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application '$($Object.DisplayName)' version '$($Object.DisplayVersion)' to the control."
                        $AppProperties += [PSCustomObject]@{
                            Name = $product.Name
                            Version = $product.Version
                            UninstallMethod = 'CimMethod'
                            CimInstance = $product
                        }
                    }
                }
                else {
                    Write-Verbose "Not found applications with name: $Name on the machine. $(Get-Date)."
                    Add-Log -Type "Warning" -Component "ApplicationSearch" -LogValue "Not found applications with name: $Name on the machine."
                    continue
                }
            }
            #endregion

            return $AppProperties
            
        }

        function Remove-InstalledApplicationManagement {
            
            [CmdletBinding(SupportsShouldProcess)]
            param (
                
                [Parameter  (Mandatory = $true)]
                [PSCustomObject]    $appObject,

                [Parameter  (Mandatory = $false)]
                [System.Version]    $appVersion
            )

            if ($appObject) {
                foreach ($Application in $appObject) {
                    $uninstallSwitch = $false
                    if ($appVersion -and ([System.Version]::Parse($Application.version) -lt $appVersion)) {
                        $uninstallSwitch = $true
                    }
                    elseif ($appVersion -and ([System.Version]::Parse($Application.version) -ge $appVersion)) {
                        if ($PSCmdlet.ShouldProcess(("Application {0} with version equal or greater than current {1}. Skipping uninstallation." -f $Application.Name, $Application.Version), $null, $null)) {
                            Write-Verbose "Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation. $(Get-Date)."
                            Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation."    
                        }
                    }
                    elseif (!$appVersion) {
                        $uninstallSwitch = $true
                    }
                    if ($uninstallSwitch) {
                        switch ($Application.UninstallMethod) {
                            ## Uninstalling using the UninstallString found on HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (x64 AND x86) ##
                            'Registry' {
                                if ($PSCmdlet.ShouldProcess(("[Registry] - Uninstalling application '{0}'." -f $Application.Name), ("Are you sure you want to uninstall the application '{0}'?" -f $Application.Name), $null)) {
                                    Write-Verbose "Uninstalling application '$($Application.Name)' version '$($Application.Version)'. $(Get-Date)."
                                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstalling application '$($Application.Name)' version '$($Application.Version)'."
                                    Write-Verbose "Uninstall string: $($Application.UninstallString). $(Get-Date)."
                                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstall string: $($Application.UninstallString)."
                                    try {
                                        Start-Process -FilePath cmd -ArgumentList '/c', $Application.UninstallString -NoNewWindow -Wait
                                        Write-Verbose "Uninstallation completed for application $($Application.Name). $(Get-Date)."
                                        Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstallation completed for application $($Application.Name)."
                                    }
                                    catch {
                                        Write-Verbose "Unable to uninstall application $($Application.Name). $($_.Exception.Message) $(Get-Date)."
                                        Add-Log -Type "Error" -Component "UninstallApplications" -LogValue "Unable to uninstall application $($Application.Name). $($_.Exception.Message)"
                                    }
                                }
                            }
                            ## Application not found on registry. Uninstalling Calling the CIM Method 'Uninstall' ##
                            'CimMethod' {
                                if ($PSCmdlet.ShouldProcess(("[cimInstance] - Uninstalling application '{0}'." -f $Application.Name), ("Are you sure you want to uninstall the application '{0}'?" -f $Application.Name), $null)) {
                                    Write-Verbose "Uninstalling application '$($Application.Name)' version '$($Application.Version)'. $(Get-Date)."
                                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstalling application '$($Application.Name)' version '$($Application.Version)'."
                                    try {
                                        $process = $Application.CimInstance | Invoke-CimMethod -MethodName 'Uninstall'
                                        if ($process.ReturnValue -eq 0) {
                                            Write-Verbose "Uninstallation completed for application $($Application.Name). $(Get-Date)."
                                            Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstallation completed for application $($Application.Name)."
                                        }
                                        else {
                                            Write-Verbose "Unable to uninstall application $($Application.Name). $($_.Exception.Message) $(Get-Date)."
                                            Add-Log -Type "Error" -Component "UninstallApplications" -LogValue "Unable to uninstall application $($Application.Name). $($_.Exception.Message)"
                                        }
                                    }
                                    catch {
                                        Write-Verbose "Unable to uninstall application $($Application.Name). $($_.Exception.Message) $(Get-Date)."
                                        Add-Log -Type "Error" -Component "UninstallApplications" -LogValue "Unable to uninstall application $($Application.Name). $($_.Exception.Message)"
                                    }    
                                }
                            }
                        }
                    }
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess("No applications with name $Name found to uninstall. Finishing execution.", $null, $null)) {
                    Write-Verbose "No applications with name $Name found to uninstall. Finishing execution. $(Get-Date)."
                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "No applications with name $Name found to uninstall. Finishing execution."
                }
            }
            
        }

        #endregion

        Write-Verbose "Parameter set name: '$($PSCmdlet.ParameterSetName)'"
        #region StartingNewExecution
        Add-Log -Type 'Info' -Component 'StartingNewExecution' -LogValue "################################"
        Add-Log -Type 'Info' -Component 'StartingNewExecution' -LogValue "###### Starting New Execution. #######"
        Add-Log -Type 'Info' -Component 'StartingNewExecution' -LogValue "################################"
        #endregion

    }
    
    process {

        switch ($PSCmdlet.ParameterSetName) {

            'UseVersion' {
                if ($PSCmdlet.ShouldProcess("Searching for applications with name '$Name' on the machine.", $null, $null)) {
                    if ($MsiParameters) { $queryParams = @{ AppName = $Name; uninstallParameters = $MsiParameters } }
                    else { $queryParams = @{ AppName = $Name } }
                    $appProperties = Get-InstalledApplicationManagement @queryParams
                }
            }

            'UninstallWVersion' {
                if ($MsiParameters) { $queryParams = @{ AppName = $Name; uninstallParameters = $MsiParameters } }
                else { $queryParams = @{ AppName = $Name } }
                $appProperties = Get-InstalledApplicationManagement @queryParams
                
                if ($WhatIf) { $removeParams = @{ appObject = $appProperties; appVersion = $Version; WhatIf = $true } }
                elseif ($Verbose) { $removeParams = @{ appObject = $appProperties; appVersion = $Version; Verbose = $true } }
                elseif ($WhatIf -and $Verbose) { $removeParams = @{ appObject = $appProperties; appVersion = $Version; WhatIf = $true; Verbose = $true } }
                else { $removeParams = @{ appObject = $appProperties; appVersion = $Version } }
                Remove-InstalledApplicationManagement @removeParams
            }

            'ForceUninstall' {
                if ($MsiParameters) { $queryParams = @{ AppName = $Name; uninstallParameters = $MsiParameters } }
                else { $queryParams = @{ AppName = $Name } }
                $appProperties = Get-InstalledApplicationManagement @queryParams
                
                if ($WhatIf) { $removeParams = @{ appObject = $appProperties; WhatIf = $true } }
                elseif ($Verbose) { $removeParams = @{ appObject = $appProperties; Verbose = $true } }
                elseif ($WhatIf -and $Verbose) { $removeParams = @{ appObject = $appProperties; WhatIf = $true; Verbose = $true } }
                else { $removeParams = @{ appObject = $appProperties } }
                Remove-InstalledApplicationManagement @removeParams
            }

        }

    }
           
    end {
        
    }

}
