function Invoke-InstalledApplicationManagement {

    <#PSScriptInfo

    .VERSION 1.0

    .GUID e36a6e01-018b-4329-a42d-099c8db2e3f2

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

    .SYNOPSIS

        This solution was designed to identify applications installed on the machine and remove it if required.

    .DESCRIPTION

        - Check for apllications installed on the machine with the input Name.
        - If a match is found on the registry, creates an object with its Name, Version and UninstallString.
        - If no match is found on the registry, the Win32_Product CIM/WMI class is queried and an object is created with the app Name, Version and CimInstance.
        - If the Uninstall switch is called, check if the installed version is less than input version and uninstall the application(s).
        - If the ForceUninstall switch is called, uninstall the application(s) without checking the version.

    .PARAMETER Name

        Name of the application to manage.
        The input will be set between wildcards.

    .PARAMETERVersion

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
            try {
                Add-Content -Path "$Env:windir\Logs\ManageSApplications-$Name-$Env:COMPUTERNAME.log" -Value $Content -Force -ErrorAction Stop
            }
            catch {
                Start-Sleep -Milliseconds 700
                Add-Content -Path "$Env:windir\Logs\ManageSApplications-$Name-$Env:COMPUTERNAME.log" -Value $Content -Force
            }
        }
        #endregion

        Write-Verbose "Parameter set name: '$($PSCmdlet.ParameterSetName)'"
        #region StartingNewExecution
        Add-Log -Type 'Info' -Component 'StartingNewExecution' -LogValue "################################"
        Add-Log -Type 'Info' -Component 'StartingNewExecution' -LogValue "###### Starting New Execution. #######"
        Add-Log -Type 'Info' -Component 'StartingNewExecution' -LogValue "################################"
        #endregion

        #region InitialVariablePayload
        Write-Verbose "Searching for installed applications on the machine with name: '$Name'. $(Get-Date)."
        Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Searching for installed applications on the machine with name: $Name."
        $RegObjects = $null
        $AppProperties = @()
        $RegPath = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        )
        $RegObjects = $RegPath | Get-ChildItem | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$Name*"} ## Looking for the application on the registry.
        #endregion
    
    }
    
    process {

        switch ($PSCmdlet.ParameterSetName) {
            
            'UseVersion' {

                if ($RegObjects) {
                    foreach ($Object in $RegObjects) {
                        if ($PSCmdlet.ShouldProcess(("[Registry] - Found application with name '{0}', version '{1}'" -f $Object.DisplayName, $Object.DisplayVersion), $null, $null)) {
                            if ($Object.UninstallString -like 'MsiExec*') { ## If UninstallString uses 'MsiExec.exe', we parse it and add the MsiParameters
                                $MSIParameters += " /l*vx+! ""%windir%\Logs\[MSI]$($Object.DisplayName)-$($Object.DisplayVersion)-Uninstall.log"""
                                $UninstallString = ($Object.UninstallString).Replace('/I', '/X')
                                $UninstallString = $UninstallString -replace '$', " $MSIParameters"
                                Write-Verbose "Adding application $($Object.DisplayName) to the control. $(Get-Date)"
                                Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($Object.DisplayName) to the control."
                                $AppProperties += [PSCustomObject]@{
                                    Name = $Object.DisplayName
                                    Version = $Object.DisplayVersion
                                    UninstallString = $UninstallString
                                    UninstallMethod = 'Registry'
                                }
                            }
                            else { ## If UninstallString don't use MsiExec.exe (probably uses the app .exe), we just Invoke it,
                                Write-Verbose "Adding application $($Object.DisplayName) to the control. $(Get-Date)"
                                Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($Object.DisplayName) to the control."
                                $AppProperties += [PSCustomObject]@{
                                    Name = $Object.DisplayName
                                    Version = $Object.DisplayVersion
                                    UninstallString = $Object.UninstallString
                                    UninstallMethod = 'Registry'
                                }
                            }
                        }
                    }
                }
                #endregion
        
                #region UsginCimInstance
                else { ## If no registry keys were found with the provided Name we query Win32_Product CIM\WMI class.
                    if ($Verbose) { $params = @{ Verbose = $true; Query = "Select * From Win32_Product Where Name Like '%$Name%'" } } ## Preparing the parameters for splatting on Invoke-CimInstance.
                    else { $params = @{ Query = "Select * From Win32_Product Where Name Like '%$Name%'" } }
                    $cimInstance = Get-CimInstance @params
                    if ($cimInstance) {
                        foreach ($product in $cimInstance) {
                            if ($PSCmdlet.ShouldProcess(("[cimInstance] - Found application with name '{0}', version '{1}'" -f $product.Name, $product.Version), $null, $null)) {
                                Write-Verbose "Adding application $($product.Name) to the control. $(Get-Date)"
                                Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($product.Name) to the control."
                                $AppProperties += [PSCustomObject]@{
                                    Name = $product.Name
                                    Version = $product.Version
                                    UninstallMethod = 'CimMethod'
                                    CimInstance = $product
                                }
                            }
                        }
                    }
                    else {
                        if ($PSCmdlet.ShouldProcess("No products found with name '$Name' installed on the machine.", $null, $null)) {
                            Write-Verbose "Not found applications with name: $Name on the machine. $(Get-Date)."
                            Add-Log -Type "Warning" -Component "ApplicationSearch" -LogValue "Not found applications with name: $Name on the machine."
                            return
                        }
                    }
                }
                #endregion
                if ($PSCmdlet.ShouldProcess("Uninstall switch not called. Finishing execution.", $null, $null)) {
                    Write-Verbose "Uninstall switch not called. Finishing execution. $(Get-Date)."
                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstall switch not called. Finishing execution."
                }
                
            }

            'UninstallWVersion' {

                if ($RegObjects) {
                    foreach ($Object in $RegObjects) {
                        if ($Object.UninstallString -like 'MsiExec*') { ## If UninstallString uses 'MsiExec.exe', we parse it and add the MsiParameters
                            $MSIParameters += " /l*vx+! ""%windir%\Logs\[MSI]$($Object.DisplayName)-$($Object.DisplayVersion)-Uninstall.log"""
                            $UninstallString = ($Object.UninstallString).Replace('/I', '/X')
                            $UninstallString = $UninstallString -replace '$', " $MSIParameters"
                            Write-Verbose "Adding application $($Object.DisplayName) to the control. $(Get-Date)"
                            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($Object.DisplayName) to the control."
                            $AppProperties += [PSCustomObject]@{
                                Name = $Object.DisplayName
                                Version = $Object.DisplayVersion
                                UninstallString = $UninstallString
                                UninstallMethod = 'Registry'
                            }
                        }
                        else { ## If UninstallString don't use MsiExec.exe (probably uses the app .exe), we just Invoke it,
                            Write-Verbose "Adding application $($Object.DisplayName) to the control. $(Get-Date)"
                            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($Object.DisplayName) to the control."
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
                    if ($Verbose) { $params = @{ Verbose = $true; Query = "Select * From Win32_Product Where Name Like '%$Name%'" } } ## Preparing the parameters for splatting on Invoke-CimInstance.
                    else { $params = @{ Query = "Select * From Win32_Product Where Name Like '%$Name%'" } }
                    $cimInstance = Get-CimInstance @params
                    if ($cimInstance) {
                        foreach ($product in $cimInstance) {
                            Write-Verbose "Adding application $($product.Name) to the control. $(Get-Date)"
                            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($product.Name) to the control."
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
                        return
                    }
                }
                #endregion

                #region Uninstall
                $Version = [System.Version]::Parse($Version)
                if ($AppProperties) {
                    foreach ($Application in $AppProperties) {
                        $CheckVersion = [System.Version]::Parse($Application.Version)
                        if ($CheckVersion -lt $Version) {
                            switch ($Application.UninstallMethod) {
                                ## Uninstalling using the UninstallString found on HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (x64 AND x86) ##
                                'Registry' {
                                    if ($PSCmdlet.ShouldProcess(("[Registry] - Uninstalling application '{0}'." -f $Application.Name), ("Are you sure you want to uninstall the application '{0}'?" -f $Application.Name), $null)) {
                                        Write-Verbose "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall. $(Get-Date)."
                                        Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall."
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
                                        Write-Verbose "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall. $(Get-Date)."
                                        Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall."
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
                        else {
                            if ($PSCmdlet.ShouldProcess(("Application {0} with version equal or greater than current {1}. Skipping uninstallation." -f $Application.Name, $Application.Version), $null, $null)) {
                                Write-Verbose "Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation. $(Get-Date)."
                                Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation."    
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
                #endregion

            }

            'ForceUninstall' {

                if ($RegObjects) {
                    foreach ($Object in $RegObjects) {
                        if ($Object.UninstallString -like 'MsiExec*') { ## If UninstallString uses 'MsiExec.exe', we parse it and add the MsiParameters
                            $MSIParameters += " /l*vx+! ""%windir%\Logs\[MSI]$($Object.DisplayName)-$($Object.DisplayVersion)-Uninstall.log"""
                            $UninstallString = ($Object.UninstallString).Replace('/I', '/X')
                            $UninstallString = $UninstallString -replace '$', " $MSIParameters"
                            Write-Verbose "Adding application $($Object.DisplayName) to the control. $(Get-Date)"
                            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($Object.DisplayName) to the control."
                            $AppProperties += [PSCustomObject]@{
                                Name = $Object.DisplayName
                                Version = $Object.DisplayVersion
                                UninstallString = $UninstallString
                                UninstallMethod = 'Registry'
                            }
                        }
                        else { ## If UninstallString don't use MsiExec.exe (probably uses the app .exe), we just Invoke it,
                            Write-Verbose "Adding application $($Object.DisplayName) to the control. $(Get-Date)"
                            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($Object.DisplayName) to the control."
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
                    if ($Verbose) { $params = @{ Verbose = $true; Query = "Select * From Win32_Product Where Name Like '%$Name%'" } } ## Preparing the parameters for splatting on Invoke-CimInstance.
                    else { $params = @{ Query = "Select * From Win32_Product Where Name Like '%$Name%'" } }
                    $cimInstance = Get-CimInstance @params
                    if ($cimInstance) {
                        foreach ($product in $cimInstance) {
                            Write-Verbose "Adding application $($product.Name) to the control. $(Get-Date)"
                            Add-Log -Type "Info" -Component "ApplicationSearch" -LogValue "Adding application $($product.Name) to the control."
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
                        return
                    }
                }
                #endregion

                #region Uninstall
                if ($AppProperties) {
                    foreach ($Application in $AppProperties) {
                        switch ($Application.UninstallMethod) {
                            ## Uninstalling using the UninstallString found on HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (x64 AND x86) ##
                            'Registry' {
                                if ($PSCmdlet.ShouldProcess(("[Registry] - Uninstalling application '{0}'." -f $Application.Name), ("Are you sure you want to uninstall the application '{0}'?" -f $Application.Name), $null)) {
                                    Write-Verbose "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall. $(Get-Date)."
                                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall."
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
                                    Write-Verbose "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall. $(Get-Date)."
                                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with outdated version $($Application.Version). Calling uninstall."
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
                else {
                    if ($PSCmdlet.ShouldProcess("No applications with name $Name found to uninstall. Finishing execution.", $null, $null)) {
                        Write-Verbose "No applications with name $Name found to uninstall. Finishing execution. $(Get-Date)."
                        Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "No applications with name $Name found to uninstall. Finishing execution."
                    }
                }
            }
            #endregion
        }
    }
           
    end {
        
    }

}
