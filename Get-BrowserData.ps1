[CmdletBinding()]
Param
(

    [Parameter  (Position = 0)]
    [ValidateSet('Chrome', 'IE', 'FireFox', 'Edge', 'All')]
    [String[]]  $Browser = 'All',

    [Parameter  (Position = 1)]
    [ValidateSet('History', 'Bookmarks', 'All')]
    [String[]]  $DataType = 'All',

    [Parameter  (Position = 2)]
    [String]    $UserName

)
function Get-ChromiumBasedHistory {

    [CmdletBinding()]
    param (
        [Parameter  ()]
        [ValidateSet('Chrome', 'Edge')]
        [String]    $BrowserBrand = 'Chrome'
    )

    switch ($BrowserBrand) {
        'Chrome' { $dataPath = 'Google\Chrome' }
        'Edge' { $dataPath = 'Microsoft\Edge' }
    }

    if (!(Test-Path -Path "$env:SystemDrive\Users\$UserName\AppData\Local\$dataPath\User Data\Default\History" -PathType Leaf -ErrorAction Ignore)) {
        continue "Could not find Chrome History for username: '$UserName'. $(Get-Date)."
    }
    else {
        $chromiumHistoryOutput = @()
        $regex = '(htt(p|s))://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?'
        $historyValue = Get-Content -Path "$env:SystemDrive\Users\$UserName\AppData\Local\$dataPath\User Data\Default\History" | Select-String -AllMatches $regex | ForEach-Object { $_.Matches.Value } | Select-Object -Unique
        foreach ($value in $historyValue) {
            $chromiumHistoryOutput += [PSCustomObject]@{
                User     = $UserName
                Browser  = $BrowserBrand
                DataType = 'History'
                Data     = $value
            }
        }
        return $chromiumHistoryOutput
    }
            
}
function Get-ChromiumBasedBookmarks {

    [CmdletBinding()]
    param (
        [Parameter  ()]
        [ValidateSet('Chrome', 'Edge')]
        [String]    $BrowserBrand = 'Chrome'
    )

    switch ($BrowserBrand) {
        'Chrome' { $dataPath = 'Google\Chrome' }
        'Edge' { $dataPath = 'Microsoft\Edge' }
    }

    if (!(Test-Path -Path "$env:SystemDrive\Users\$UserName\AppData\Local\$dataPath\User Data\Default\Bookmarks" -PathType Leaf -ErrorAction Ignore)) {
        continue "Could not find Chrome History for username: '$UserName'. $(Get-Date)."
    }
    else {
        $chromiumBookmarksOutput = @()
        $bookmarksUrls = (Get-Content "$env:SystemDrive\Users\$UserName\AppData\Local\$dataPath\User Data\Default\Bookmarks" | ConvertFrom-Json).roots.bookmark_bar.children.url | Select-Object -Unique
        foreach ($url in $bookmarksUrls) {
            $chromiumBookmarksOutput += [PSCustomObject]@{
                User = $UserName
                Browser = $BrowserBrand
                DataType = 'Bookmarks'
                Data = $url
            }
        }
        return $chromiumBookmarksOutput
    }
}
function Get-InternetExplorerHistory {
    
    $userProfileSid = ([System.Security.Principal.NTAccount] $UserName).Translate([System.Security.Principal.SecurityIdentifier]).Value
    [void](New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS)
    if (!(Test-Path -Path "HKU:\$userProfileSid\SOFTWARE\Microsoft\Internet Explorer\TypedURLs\" -ErrorAction Ignore)) {
        continue "Could not find IE History for user: '$UserName'. $(Get-Date)."
    }
    else {
        $ieHistoryOutput = @()
        $typedUrls = Get-Item -Path "HKU:\$userProfileSid\SOFTWARE\Microsoft\Internet Explorer\TypedURLs\" -ErrorAction Ignore
        foreach ($valueName in $typedUrls.GetValueNames()) {
            $ieHistoryOutput += [PSCustomObject]@{
                User = $UserName
                Browser = 'Internet Explorer'
                DataType = 'History'
                Data = $typedUrls.GetValue($valueName)
            }
        }
        return $ieHistoryOutput
    }
    
}
function Get-InternetExplorerBookmarks {

    $favoriteFiles = Get-ChildItem -Path C:\Users\$UserName\Favorites\ -Filter "*.url" -ErrorAction Ignore
    if ($favoriteFiles) {
        $ieBookmarksOutput = @()
        foreach ($file in $favoriteFiles) {
            $bookMarkName = $file.BaseName
            $bookMarkUrl = (Get-Content -Path $file.FullName | Select-String -Pattern 'URL=') -replace '^.*URL='
            $ieBookmarksOutput += [PSCustomObject]@{
                User = $UserName
                Browser = 'Internet Explorer'
                DataType = 'Bookmarks'
                Name = $bookMarkName
                Data = $bookMarkUrl
            }
        }
    }
    else {
        continue "Could not find IE Bookmarks for user: '$UserName'. $(Get-Date)."
    }
    return $ieBookmarksOutput
    
}
function Get-FireFoxHistory {
    $Path = "$Env:systemdrive\Users\$UserName\AppData\Roaming\Mozilla\Firefox\Profiles\"
    if (-not (Test-Path -Path $Path)) {
        Write-Verbose "[!] Could not find FireFox History for username: $UserName"
    }
    else {
        $Profiles = Get-ChildItem -Path "$Path\*.default\" -ErrorAction SilentlyContinue
        $Regex = '(htt(p|s))://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?'
        $Value = Get-Content $Profiles\places.sqlite | Select-String -Pattern $Regex -AllMatches | Select-Object -ExpandProperty Matches | Sort -Unique
        $Value.Value | ForEach-Object {
            if ($_ -match $Search) {
                ForEach-Object {
                    New-Object -TypeName PSObject -Property @{
                        User     = $UserName
                        Browser  = 'Firefox'
                        DataType = 'History'
                        Data     = $_
                    }    
                }
            }
        }
    }
}
if (!$UserName) {
    $UserName = "$ENV:USERNAME"
}
if (($Browser -Contains 'All') -or ($Browser -Contains 'Chrome')) {
    if (($DataType -Contains 'All') -or ($DataType -Contains 'History')) {
        Get-ChromeHistory
    }
    if (($DataType -Contains 'All') -or ($DataType -Contains 'Bookmarks')) {
        Get-ChromeBookmarks
    }
}
if (($Browser -Contains 'All') -or ($Browser -Contains 'IE')) {
    if (($DataType -Contains 'All') -or ($DataType -Contains 'History')) {
        Get-InternetExplorerHistory
    }
    if (($DataType -Contains 'All') -or ($DataType -Contains 'Bookmarks')) {
        Get-InternetExplorerBookmarks
    }
}
if (($Browser -Contains 'All') -or ($Browser -Contains 'FireFox')) {
    if (($DataType -Contains 'All') -or ($DataType -Contains 'History')) {
        Get-FireFoxHistory
    }
}
