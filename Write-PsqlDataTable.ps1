
<#PSScriptInfo

.VERSION 1.1.0

.GUID 4901f60c-7d08-462d-8351-27b057ea549c

.AUTHOR francisco.nabas

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES PostgreSQL ODBC x64 Driver 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Script to "Bulk" copy data into PostgreSQL database using the COPY statement and a csv file. Input is a PSCustomObject. 

#> 

Param()



<#PSScriptInfo

.VERSION 1.1.0

.GUID 62de916e-e033-4946-95f9-e353debcd349

.AUTHOR Francisco Nabas

.COMPANYNAME 

.COPYRIGHT 

.TAGS PostgreSQL

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES PostgreSQL ODBC x64 Driver

.RELEASENOTES


#>

<# 

.DESCRIPTION 
 Script to "Bulk" copy data into PostgreSQL database using the COPY statement and a csv file. Input is a PSCustomObject
 
#> 
Function Write-PsqlDataTable
{

    [CmdletBinding()] 
    param(

        [Parameter  (Position = 0, Mandatory = $true)]
        [string]    $ServerInstance,

        [Parameter  (Mandatory = $false)]
        [int]       $Port=5432,

        [Parameter  (Position = 1, Mandatory = $true)]
        [string]    $Database,

        [Parameter  (Position = 2, Mandatory = $true)]
        [string]    $TableName,

        [Parameter  (Position = 3, Mandatory = $true)]
        [System.Data.DataTable] $Data,
                    
        [Parameter  (Position = 4, Mandatory = $false)] 
        [string]    $Username,

        [Parameter  (Position = 5, Mandatory = $false)]
        [Security.SecureString] $Password,

        [Parameter  (Mandatory = $false)]
        [string]    $Docker

    ) 
     
    
    $DBConn = New-Object System.Data.Odbc.OdbcConnection
    $BTSR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PTP = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BTSR)
    if ($Username) {
        $DBConnectionString = "Driver={PostgreSQL UNICODE(x64)};Server=$ServerInstance;Port=$Port;Database=$Database;Uid=$Username;Pwd=$PTP;"
    } 
    else {
        $DBConnectionString = "Driver={PostgreSQL UNICODE(x64)};Server=$ServerInstance;Port=$Port;Database=$Database;"
    } 
 
    $DBConn.ConnectionString = $DBConnectionString
         
    try 
    {
        $Columns = $Data.Columns.ColumnName
        $Columns = [System.String]::Join(',',$Columns)
        $Data | Export-Csv $Env:TEMP\TempPsAd.csv -Delimiter ';' -NoTypeInformation
        if ($Docker) {
            docker cp $Env:TEMP\TempPsAd.csv "$($Docker):/media/TempPsAd.csv"
            $DBConn.Open()
            $DBCmd = $DBConn.CreateCommand()
            $DBCmd.CommandText = @"
                COPY $TableName ($Columns)
                FROM '/media/TempPsAd.csv'
                DELIMITER ';'
                CSV HEADER
"@
            $DBCmd.ExecuteReader()
            $DBConn.Close()
            docker exec $Docker rm -rf /media/TempPsAd.csv
            Remove-Item $Env:TEMP\TempPsAd.csv -Force
        }
        else {
            $DBConn.Open()
            $DBCmd = $DBConn.CreateCommand()
            $DBCmd.CommandText = @"
                COPY $TableName ($Columns)
                FROM '$Env:TEMP\TempPsAd.csv'
                DELIMITER ';'
                CSV HEADER
"@
            $DBCmd.ExecuteReader()
            $DBConn.Close()
            Remove-Item $Env:TEMP\TempPsAd.csv -Force
        }
        
    } 
    catch 
    { 
        Write-Error "$($_.Exception.Message)"
        continue 
    } 
}
