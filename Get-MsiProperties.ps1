function Get-MsiProperties {
    
    [CmdletBinding()]
    param (
        
        [Parameter  (Mandatory = $true, Position = 0)]
        [string]    $Path
    )
    
    begin {
        
        try {
            $uncPath = & {
                try {
                    $itemPath = Get-Item -Path $Path -ErrorAction Stop
                    if (($itemPath.Attributes -ne 'Archive') -or ($itemPath.Extension -ne '.msi')) {
                        throw "Invalid path. Input a MSI file."
                    }
                    return $itemPath.FullName
                }
                catch {
                    throw $_.Exception.Message
                }
            }
            $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
            $database = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $windowsInstaller, @($uncPath, 0))
            $query = "SELECT Property, Value FROM Property"
            $propView = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $database, ($query))
            $propView.GetType().InvokeMember("Execute", "InvokeMethod", $null, $propView, $null) | Out-Null
            $propRecord = $propView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $propView, $null)    
        }
        catch {
            throw $_.Exception.Message
        }
        
    }
    
    process {
        
        $output = while  ($null -ne $propRecord)
        {
        	$col1 = $propRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $propRecord, 1)
        	$col2 = $propRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $propRecord, 2)
        
        	@{$col1 = $col2}
        
        	#fetch the next record
        	$propRecord = $propView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $propView, $null)	
        }

    }
    
    end {
        
        $propView.GetType().InvokeMember("Close", "InvokeMethod", $null, $propView, $null) | Out-Null          
        $propView = $null 
        $propRecord = $null
        $database = $null
        return $output

    }
}