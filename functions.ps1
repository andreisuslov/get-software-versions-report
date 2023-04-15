$jsonReportPath = "./report.json"
$htmlReportPath = "./report.html"
$appsJsonFilePath = "./apps.json"

function emptyFile($filePath) {
    $fileExists = checkIfFileExists $filePath
    if ($fileExists) {
        set-content $filePath -value ""
    }
    else {
        write-error "${filePath} file not found."
    }
}

function checkIfFileExists($filePath) {
    if (test-path $filePath) {
        return $true
    }
    else {
        return $false
    }
}

function writeToFile($path, $value) {
    $fileExists = checkIfFileExists $path
    if ($fileExists) {
        out-file -filePath $path -inputObject $value -encoding utf8
    }
    else {
        write-error "${path} file not found."
    }
}

function getFileContent($filePath) {
    $fileExists = checkIfFileExists $filePath
    if ($fileExists) {
        $fileContent = get-content -path $filePath -raw
        return $fileContent
    }
    else {
        write-error "${filePath} file not found."
    }
}

function getExecutableVersion {
    param([parameter(mandatory = $true)][string]$executablePath)
    if (Test-Path -Path $executablePath) {
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($executablePath)
        $productVersionRaw = $versionInfo.ProductMajorPart, $versionInfo.ProductMinorPart, $versionInfo.ProductBuildPart, $versionInfo.ProductPrivatePart -join "."
        return $productVersionRaw
    }
    else {
        return "N/A"
    }
}

function getAppsJsonFileContent() {
    $appsJsonFileContent = getFileContent $appsJsonFilePath
    return $appsJsonFileContent
}

function getNames() {
    $appsInfo = getAppsJsonFileContent
    $appsObj = convertFrom-json $appsInfo
    return $appsObj.psObject.properties.name
}

function getPaths() {
    $appsInfo = getAppsJsonFileContent
    $appsObj = convertFrom-json $appsInfo
    return $appsObj.psObject.properties.value
}

function getVersions() {
    $appPaths = getPaths
    $versions = @()
    foreach ($appPath in $appPaths) {
        $version = getExecutableVersion -executablePath $appPath
        $versions += $version
    }
    return $versions
}

function getApps() {
    $apps = @()
    $names = getNames
    $versions = getVersions
    for ($i = 0; $i -lt $names.count; $i++) {
        $apps += [PSCustomObject]@{
            name    = $names[$i]
            version = $versions[$i]
        }
    }
    return $apps
}

function getNamesToVersionsJsonMap() {
    $map = [ordered]@{}
    $apps = getApps
    foreach ($product in $apps) {
        $map[$product.name] = $product.version
    }
    $map | ConvertTo-Json
}


function createNewFile($filePath) {
    $fileExists = checkIfFileExists $filePath
    if ($fileExists) { return }
    $file = new-item -itemType file -path $filePath
    return $file
}
function generateNodeSoftwareInventoryItem() {
    $hostName = hostName
    $namesToVersionsMap = (getNamesToVersionsJsonMap | convertFrom-json)
    $output = @{
        $hostName = $namesToVersionsMap
    }
    return $output
}

function getInventoryItemNodeName([hashtable]$item) {
    $nodeName = $item.psObject.properties.name # get the key of the hashtable
    return $nodeName
}

function addItemToJsonReport([hashtable]$item) {
    $report = Get-Content -Path $jsonReportPath -Raw # remove | ConvertFrom-Json -AsHashtable
    $hashtable = @{}
    (ConvertFrom-Json $report).psobject.properties | Foreach { $hashtable[$_.Name] = $_.Value } # convert the JSON report to a hashtable
    $report = $hashtable 
    if (-not $report) {
        $report = @{}
    } # if the JSON report is empty, initialize it as an empty hashtable
    foreach ($key in $item.Keys) {
        # if the node already exists in the JSON report, don't add it again
        if (-not $report.ContainsKey($key)) {
            $report[$key] = $item[$key]
        }
    }
    $updatedJsonReport = ConvertTo-Json $report -Depth 100
    Out-File -FilePath $jsonReportPath -InputObject $updatedJsonReport -Encoding utf8
}

function getJsonReportFileContent {
    $content = Get-Content $jsonReportPath | ConvertFrom-Json
    $sortedContent = [ordered]@{}
    foreach ($nodeName in ($content.psObject.properties.name | sort-object)) {
        $sortedContent[$nodeName] = $content.$nodeName
    }
    $psCustomObject = new-object -typeName psCustomObject -property $sortedContent
    return $psCustomObject | ConvertTo-Json -Depth 100
}

function clearJsonReportFileContent() {
    emptyFile $jsonReportPath
}

function getHtmlStyles() {
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

function getHostNames($data) {
    $hostNames = $data.psObject.properties.name
    return $hostNames
}

function generateHtmlTableHeader($data) {
    $header = ""
    $hostNames = getHostNames $data
    foreach ($hostName in $hostNames) {
        $header += "<th>$hostName</th>"
    }
    return $header
}

function getUniqueAppNames ($data) {
    $uniqueApps = @{}
    $hostNames = getHostNames $data
    foreach ($hostName in $hostNames) {
        $appNames = $data.$hostName.psObject.properties.name
        foreach ($appName in $appNames) {
            $uniqueApps[$appName] = $true # use a hashtable to get unique app names
        }
    }
    $uniqueAppNames = $uniqueApps.keys
    $uniqueAppNamesSorted = $uniqueAppNames | sort-object
    return $uniqueAppNamesSorted
}

function createVersionCell ($data, $hostName, $appName) {
    $version = $data.$hostName.$appName
    if ($null -eq $version) {
        $version = "N/A"
    }
    return "<td>$version</td>"
}

function createAppRow ($data, $appName) {
    $row = "<tr><th>$appName</th>"
    $hostNames = getHostNames $data
    foreach ($hostName in $hostNames) {
        $row += (createVersionCell $data $hostName $appName)
    }
    $row += "</tr>"
    return $row
}

function generateHtmlTableRows ($data) {
    $rows = ""
    $appNames = getUniqueAppNames $data
    foreach ($appName in $appNames) {
        $rows += (createAppRow $data $appName)
    }
    return $rows
}

function createHtmlHead() {
    return @"
<head>
    <title>"Software Inventory Report"</title>
    <style>
    $(getHtmlStyles)
    </style>
</head>
"@
}

function createHtmlTable ($data) {
    return @"
<table>
    <tr>
        <th>Software Products / Host Names</th>
        $(generateHtmlTableHeader $data)
    </tr>
    $(generateHtmlTableRows $data)
</table>
"@
}

function createFullHtmlDocument($data) {
    $html = @"
<!DOCTYPE html>
<html>
    $(createHtmlHead)
<body>
    $(createHtmlTable $data)
</body>
</html>
"@
    return $html
}

function createSoftwareInventoryHtmlReport([string]$jsonReportFileContent, [string]$htmlReportPath = "./report.html") {
    $data = $jsonReportFileContent | ConvertFrom-Json
    $html = createFullHtmlDocument $data
    writeToFile $htmlReportPath $html
}

function getHtmlReportFileContent() {
    $htmlReportFileContent = getFileContent $htmlReportPath
    return $htmlReportFileContent
}

function clearHtmlReportFileContent {
    emptyFile $htmlReportPath
}
