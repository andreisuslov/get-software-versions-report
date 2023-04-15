clearJsonReportFileContent
clearHtmlReportFileContent

$a = @{
    "us18739asuslov" = @{
        "NEXIS Client Manager" = "22"
        "Interplay Transfer Client" = "4.0.1.23758"
    }
}

$b = @{
    "A51-WG7-AUTO1" = @{
        "NEXIS Client Manager" = "22.9.0.135"
        "Interplay Transfer Client" = "4.0.1.23758"
    }
}

$c = @{
    "A51-WG7-AUTO2" = @{
        "NEXIS Client Manager" = "22.9.0.135"
        "Interplay Transfer Client" = "4.0.1.23758"
    }
}

addInventoryItemToJsonReport $a
addInventoryItemToJsonReport $b
addInventoryItemToJsonReport $c

$jsonReportFileContent = getJsonReportFileContent
createSoftwareInventoryHtmlReport $jsonReportFileContent
getHtmlReportFileContent
