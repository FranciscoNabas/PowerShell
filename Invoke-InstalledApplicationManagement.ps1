function Invoke-InstalledApplicationManagement {

    #Requires -RunAsAdministrator

    [CmdletBinding(SupportsShouldProcess)]
    param (

        [Parameter  (Mandatory, Position = 0, ParameterSetName = 'UseVersion')]
        [Parameter  (Mandatory, ParameterSetName = 'ForceUninstall')]
        [Parameter  (Mandatory, ParameterSetName = 'Uninstall')]
        [ValidateNotNullOrEmpty()]
        [string]    $Name,

        [Parameter  (Mandatory, Position = 1, ParameterSetName = 'UseVersion')]
        [Parameter  (Mandatory, ParameterSetName = 'Uninstall')]
        [ValidateNotNullOrEmpty()]
        [string]    $CurrentVersion,

        [Parameter  (ParameterSetName = 'UseVersion')]
        [Parameter  (Mandatory, ParameterSetName = 'Uninstall')]
        [Parameter  (Mandatory, ParameterSetName = 'ForceUninstall')]
        [switch]    $Uninstall,

        [Parameter  (Mandatory, ParameterSetName = 'ForceUninstall')]
        [Parameter  (ParameterSetName = 'Uninstall')]
        [switch]    $ForceUninstall,

        [Parameter  (ParameterSetName = 'ForceUninstall')]
        [Parameter  (ParameterSetName = 'Uninstall')]
        [Parameter  (ParameterSetName = 'UseVersion')]
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

        #region UsingRegistry
        if ($RegObjects) {
            foreach ($Object in $RegObjects) {
                if ($PSCmdlet.ShouldProcess("[Registry] - Found file with name '$($Object.DisplayName)', version '$($Object.DisplayVersion)'", $Object, 'Query')) {
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
                    if ($PSCmdlet.ShouldProcess("[cimInstance] - Found file with name '$($product.Name)', version '$($product.Version)'", $product, 'Query')) {
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
                if ($PSCmdlet.ShouldProcess("No products found with name '$Name' installed on the machine.", $Name, 'Query')) {
                    Write-Verbose "Not found applications with name: $Name on the machine. $(Get-Date)."
                    Add-Log -Type "Warning" -Component "ApplicationSearch" -LogValue "Not found applications with name: $Name on the machine."
                    return
                }
            }
        }
        #endregion
        
        #region UninstallingApplications
        if ($Uninstall) {
            ## ForceInstall not called. Checking version prior to Uninstallation ##
            if (!$ForceUninstall) {
                $CurrentVersion = [System.Version]::Parse($CurrentVersion)
                if ($AppProperties) {
                    foreach ($Application in $AppProperties) {
                        $CheckVersion = [System.Version]::Parse($Application.Version)
                        if ($CheckVersion -lt $CurrentVersion) {
                            switch ($Application.UninstallMethod) {
                                ## Uninstalling using the UninstallString found on HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (x64 AND x86) ##
                                'Registry' {
                                    if ($PSCmdlet.ShouldProcess("[Registry] - Uninstalling application '$($Application.Name).'", $Application, 'Uninstall')) {
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
                                    if ($PSCmdlet.ShouldProcess("[cimInstance] - Uninstalling application '$($Application.Name).'", $Application, 'Uninstall')) {
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
                            if ($PSCmdlet.ShouldProcess("Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation.", $Application, 'Uninstall')) {
                                Write-Verbose "Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation. $(Get-Date)."
                                Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Application $($Application.Name) with version equal or greater than current $($Application.Version). Skipping uninstallation."    
                            }
                        }
                    }
                }
                else {
                    Write-Verbose "No applications with name $Name found to uninstall. Finishing execution. $(Get-Date)."
                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "No applications with name $Name found to uninstall. Finishing execution."    
                }
            }
            ## ForceInstall called. Skipping version check prior to uninstallation. ##
            else {
                if ($AppProperties) {
                    foreach ($Application in $AppProperties) {
                        switch ($Application.UninstallMethod) {
                            ## Uninstalling using the UninstallString found on HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (x64 AND x86) ##
                            'Registry' {
                                if ($PSCmdlet.ShouldProcess("[Registry] - Uninstalling application '$($Application.Name).'", $Application, 'Uninstall')) {
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
                                if ($PSCmdlet.ShouldProcess("[cimInstance] - Uninstalling application '$($Application.Name).'", $Application, 'Uninstall')) {
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
                    Write-Verbose "No applications with name $Name found to uninstall. Finishing execution. $(Get-Date)."
                    Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "No applications with name $Name found to uninstall. Finishing execution."    
                }
            }
        }
        #endregion
        
        else {
            if ($PSCmdlet.ShouldProcess("Uninstall switch not called. Finishing execution.", $Name, 'Uninstall')) {
                Write-Verbose "Uninstall switch not called. Finishing execution. $(Get-Date)."
                Add-Log -Type "Info" -Component "UninstallApplications" -LogValue "Uninstall switch not called. Finishing execution."    
            }
        }

    }
    
    end {
        
    }
}