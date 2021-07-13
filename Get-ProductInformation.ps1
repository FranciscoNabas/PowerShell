function Get-ProductInformation {

    [CmdletBinding()]
    param (
        [Parameter  (Mandatory = $true, Position = 0)]
        [string]    $Name
    )

    $systemProfileSID = (Get-CimInstance -Query "Select SID From Win32_UserProfile Where LocalPath Like '%systemprofile'").SID
    $regInstallerPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$systemProfileSID\Products"
    $regUninstallPath = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $productProperties = @()
    $productItems = Get-ChildItem -Path $regUninstallPath | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$Name*"}
    if ($productItems) {
        foreach ($product in $productItems) {
            $localPackage = (Get-ChildItem -Path $regInstallerPath -Recurse | Get-ItemProperty | Where-Object {($_.DisplayName -eq $product.DisplayName) -and ($_.UninstallString -eq $product.UninstallString)}).LocalPackage
            $productProperties += [PSCustomObject]@{
                Name = $product.DisplayName
                Version = $product.DisplayVersion
                Location = $product.InstallLocation
                InstallSource = $product.InstallSource
                IdentifyingNumber = $product.PSChildName
                InstallerCache = $localPackage
            }
        }    
    }
    else {
        if ($Verbose) { $params = @{ Verbose = $true; Query = "Select * From Win32_Product Where Name Like '%$Name%'"; ErrorAction = 'SilentlyContinue' } }
        else { $params = @{ Query = "Select * From Win32_Product Where Name Like '%$Name%'"; ErrorAction = 'SilentlyContinue' } }
        $cimInstance = Get-CimInstance @params
        if ($cimInstance) {
            foreach ($product in $cimInstance) {
                $productProperties += [PSCustomObject]@{
                    Name = $product.Name
                    Version = $product.Version
                    Location = $product.InstallLocation
                    InstallSource = $product.InstallSource
                    IdentifyingNumber = $product.IdentifyingNumber
                    InstallerCache = $product.LocalPackage
                }
            }
        }
        else {
            throw "Not found product with name '$Name'."
        }
    }
    
    return $productProperties

}