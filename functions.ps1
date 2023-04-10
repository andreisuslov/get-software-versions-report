$JsonReportPath = "./report.json"
$HtmlReportPath = "./report.html"
$AppsJsonFilePath = "./apps.json"

function Empty-File {
    param([string]$FilePath)
    $FileExists = Check-IfFileExists $FilePath
    if ($FileExists) {
        Set-Content $FilePath -Value ""
    }
    else {
        Write-Error "${FilePath} file not found."
    }
}

function Check-IfFileExists {
    param([string]$FilePath)
    if (Test-Path $FilePath) {
        return $true
    }
    else {
        return $false
    }
}

function Create-NewFile {
    param([string]$FilePath)
    $FileExists = Check-IfFileExists $FilePath
    if ($FileExists) { return }
    $File = New-Item -ItemType File -Path $FilePath
    return $File
}

function Write-ToFile {
    param([string]$Path, [string]$Value)
    $FileExists = Check-IfFileExists $Path
    if ($FileExists) {
        Out-File -FilePath $Path -InputObject $Value -Encoding utf8
    }
    else {
        Write-Error "${Path} file not found."
    }
}

function Get-FileContent {
    param([string]$FilePath)
    $FileExists = Check-IfFileExists $FilePath
    if ($FileExists) {
        $FileContent = Get-Content -Path $FilePath -Raw
        return $FileContent
    }
    else {
        Write-Error "${FilePath} file not found."
    }
}

function Get-ExecutableVersion {
    param([parameter(Mandatory = $true)][string]$ExecutablePath)
    if (Test-Path -Path $ExecutablePath) {
        $VersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExecutablePath)
        $ProductVersionRaw = $VersionInfo.ProductMajorPart, $VersionInfo.ProductMinorPart, $VersionInfo.ProductBuildPart, $VersionInfo.ProductPrivatePart -join "."
        return $ProductVersionRaw
    }
    else {
        return "N/A"
    }
}

function Get-AppsJsonFileContent {
    $AppsJsonFileContent = Get-FileContent $AppsJsonFilePath
    return $AppsJsonFileContent
}


function Get-Names {
    $AppsInfo = Get-AppsJsonFileContent
    $AppsObj = ConvertFrom-Json $AppsInfo
    return $AppsObj.psObject.properties.name
}

function Get-Paths {
    $AppsInfo = Get-AppsJsonFileContent
    $AppsObj = ConvertFrom-Json $AppsInfo
    return $AppsObj.psObject.properties.value
}

function Get-Versions {
    $AppPaths = Get-Paths
    $Versions = @()
    foreach ($AppPath in $AppPaths) {
        $Version = Get-ExecutableVersion -ExecutablePath $AppPath
        $Versions += $Version
    }
    return $Versions
}

function Get-Apps {
    $Apps = @()
    $Names = Get-Names
    $Versions = Get-Versions
    for ($i = 0; $i -lt $Names.Count; $i++) {
        $Apps += [PSCustomObject]@{
            Name    = $Names[$i]
            Version = $Versions[$i]
        }
    }
    return $Apps
}

function Get-NamesToVersionsJsonMap {
    $Map = [ordered]@{}
    $Apps = Get-Apps
    foreach ($Product in $Apps) {
        $Map[$Product.Name] = $Product.Version
    }
    $Map | ConvertTo-Json
}

function Generate-NodeSoftwareInventoryItem {
    $HostName = HostName
    $NamesToVersionsMap = (Get-NamesToVersionsJsonMap | ConvertFrom-Json)
    $Output = @{
        $HostName = $NamesToVersionsMap
    }
    return $Output
}

function Get-InventoryItemNodeName {
    param([hashtable]$Item)
    $NodeName = $Item.psObject.properties.name # get the key of the hashtable
    return $NodeName
}

function Add-ItemToJsonReport {
    param([hashtable]$Item)
    $Report = Get-Content -Path $JsonReportPath -Raw # remove | ConvertFrom-Json -AsHashtable
    $Hashtable = @{}
    (ConvertFrom-Json $Report).psobject.properties | ForEach { $Hashtable[$_.Name] = $_.Value } # convert the JSON report to a hashtable
    $Report = $Hashtable
    if (-not $Report) {
        $Report = @{}
    } # if the JSON report is empty, initialize it as an empty hashtable
    foreach ($Key in $Item.Keys) {
        # if the node already exists in the JSON report, don't add it again
        if (-not $Report.ContainsKey($Key)) {
            $Report[$Key] = $Item[$Key]
        }
    }
    $UpdatedJsonReport = ConvertTo-Json $Report -Depth 100
    Out-File -FilePath $JsonReportPath -InputObject $UpdatedJsonReport -Encoding utf8
}

function Get-JsonReportFileContent {
    $Content = Get-Content $JsonReportPath | ConvertFrom-Json
    $SortedContent = [ordered]@{}
    foreach ($NodeName in ($Content.psObject.properties.name | Sort-Object)) {
        $SortedContent[$NodeName] = $Content.$NodeName
    }
    $PsCustomObject = New-Object -TypeName PsCustomObject -Property $SortedContent
    return $PsCustomObject | ConvertTo-Json -Depth 100
}

function Clear-JsonReportFileContent {
    Empty-File $JsonReportPath
}

function Get-HtmlStyles {
    return @"
    table {
        border-collapse: collapse;
        width: auto;
        font-family: Arial, sans-serif;
    }

    th {
        border: 1px solid rgb(183, 182, 182);
        text-align: left;
        padding: 8px;
    }

    td {
        border: 1px solid rgb(224, 223, 223);
        text-align: left;
        padding: 8px;
    }

    th {
        background-color: #f2f2f2;
        font-weight: bold;
    }

    tr:nth-child(even) {
        background-color: #e7e7e7;
    }
"@
}

function Get-HostNames {
    param($Data)
    $HostNames = $Data.psObject.properties.name
    return $HostNames
}

function Generate-HtmlTableHeader {
    param($Data)
    $Header = ""
    $HostNames = Get-HostNames $Data
    foreach ($HostName in $HostNames) {
        $Header += "<th>$HostName</th>"
    }
    return $Header
}

function Get-UniqueAppNames {
    param($Data)
    $UniqueApps = @{}
    $HostNames = Get-HostNames $Data
    foreach ($HostName in $HostNames) {
        $AppNames = $Data.$HostName.psObject.properties.name
        foreach ($AppName in $AppNames) {
            $UniqueApps[$AppName] = $true # use a hashtable to get unique app names
        }
    }
    $UniqueAppNames = $UniqueApps.Keys
    $UniqueAppNamesSorted = $UniqueAppNames | Sort-Object
    return $UniqueAppNamesSorted
}

function Create-VersionCell {
    param($Data, $HostName, $AppName)
    $Version = $Data.$HostName.$AppName
    if ($null -eq $Version) {
        $Version = "N/A"
    }
    return "<td>$Version</td>"
}

function Create-AppRow {
    param($Data, $AppName)
    $Row = "<tr><th>$AppName</th>"
    $HostNames = Get-HostNames $Data
    foreach ($HostName in $HostNames) {
        $Row += (Create-VersionCell $Data $HostName $AppName)
    }
    $Row += "</tr>"
    return $Row
}

function Generate-HtmlTableRows {
    param($Data)
    $Rows = ""
    $AppNames = Get-UniqueAppNames $Data
    foreach ($AppName in $AppNames) {
        $Rows += (Create-AppRow $Data $AppName)
    }
    return $Rows
}

function Create-HtmlHead {
    return @"
<head>
    <title>"Software Inventory Report"</title>
    <style>
    $(Get-HtmlStyles)
    </style>
</head>
"@
}

function Create-HtmlTable {
    param($Data)
    return @"
<table>
    <tr>
        <th>Software Products / Host Names</th>
        $(Generate-HtmlTableHeader $Data)
    </tr>
    $(Generate-HtmlTableRows $Data)
</table>
"@
}

function Create-FullHtmlDocument {
    param($Data)
    $Html = @"
<!DOCTYPE html>
<html>
    $(Create-HtmlHead)
<body>
    $(Create-HtmlTable $Data)
</body>
</html>
"@
    return $Html
}

function Create-SoftwareInventoryHtmlReport {
    param([string]$JsonReportFileContent, [string]$HtmlReportPath = "./report.html")
    $Data = $JsonReportFileContent | ConvertFrom-Json
    $Html = Create-FullHtmlDocument $Data
    Write-ToFile $HtmlReportPath $Html
}

function Get-HtmlReportFileContent {
    $HtmlReportFileContent = Get-FileContent $HtmlReportPath
    return $HtmlReportFileContent
}

function Clear-HtmlReportFileContent {
    Empty-File $HtmlReportPath
}