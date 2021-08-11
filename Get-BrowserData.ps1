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
    [String]    $UserName,

    [Parameter  (Position = 3)]
    [String]    $Search

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
        continue "Could not find Chrome History for username: $UserName. $(Get-Date)."
    }
    else {
        $regex = '(htt(p|s))://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?'
        $historyValue = Get-Content -Path "$env:SystemDrive\Users\$UserName\AppData\Local\$dataPath\User Data\Default\History" | Select-String -AllMatches $regex | ForEach-Object { $_.Matches.Value } | Select-Object -Unique
        foreach ($value in $historyValue) {
            if ($value -match $Search) {
                return [PSCustomObject]@{
                    User     = $UserName
                    Browser  = $BrowserBrand
                    DataType = 'History'
                    Data     = $value
                }
            }
        }
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
        continue "Could not find Chrome History for username: $UserName. $(Get-Date)."
    }
    else {
        $bookmarksUrls = (Get-Content "$env:SystemDrive\Users\$UserName\AppData\Local\$dataPath\User Data\Default\Bookmarks" | ConvertFrom-Json).roots.bookmark_bar.children.url | Select-Object -Unique
        foreach ($url in $bookmarksUrls) {
            if ($url -match $Search) {
                return [PSCustomObject]@{
                    User     = $UserName
                    Browser  = $BrowserBrand
                    DataType = 'Bookmarks'
                    Data     = $url
                }
            }
        }
    }
}
function Get-InternetExplorerHistory {
    [void](New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS)
    $Paths = Get-ChildItem 'HKU:\' -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'S-1-5-21-[0-9]+-[0-9]+-[0-9]+-[0-9]+$' }
    ForEach ($Path in $Paths) {
        $User = ([System.Security.Principal.SecurityIdentifier] $Path.PSChildName).Translate( [System.Security.Principal.NTAccount]) | Select -ExpandProperty Value
        $Path = $Path | Select-Object -ExpandProperty PSPath
        $UserPath = "$Path\Software\Microsoft\Internet Explorer\TypedURLs"
        if (-not (Test-Path -Path $UserPath)) {
            Write-Verbose "[!] Could not find IE History for SID: $Path"
        }
        else {
            Get-Item -Path $UserPath -ErrorAction SilentlyContinue | ForEach-Object {
                $Key = $_
                $Key.GetValueNames() | ForEach-Object {
                    $Value = $Key.GetValue($_)
                    if ($Value -match $Search) {
                        New-Object -TypeName PSObject -Property @{
                            User     = $UserName
                            Browser  = 'IE'
                            DataType = 'History'
                            Data     = $Value
                        }
                    }
                }
            }
        }
    }
}
function Get-InternetExplorerBookmarks {
    $URLs = Get-ChildItem -Path "$Env:systemdrive\Users\" -Filter "*.url" -Recurse -ErrorAction SilentlyContinue
    ForEach ($URL in $URLs) {
        if ($URL.FullName -match 'Favorites') {
            $User = $URL.FullName.split('\')[2]
            Get-Content -Path $URL.FullName | ForEach-Object {
                try {
                    if ($_.StartsWith('URL')) {
                        # parse the .url body to extract the actual bookmark location
                        $URL = $_.Substring($_.IndexOf('=') + 1)
                        if ($URL -match $Search) {
                            New-Object -TypeName PSObject -Property @{
                                User     = $User
                                Browser  = 'IE'
                                DataType = 'Bookmark'
                                Data     = $URL
                            }
                        }
                    }
                }
                catch {
                    Write-Verbose "Error parsing url: $_"
                }
            }
        }
    }
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
