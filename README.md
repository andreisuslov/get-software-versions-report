# get-software-versions-report

When you execute the provided commands (file `commands-to-execute.ps1`) in the PowerShell Terminal, the following steps will be performed:

- The functions from the "functions.ps1" PowerShell script will be imported into the current session.
- The content of the JSON and HTML report files will be cleared using the Clear-JsonReportFileContent and Clear-HtmlReportFileContent functions.
- Three hashtables ($a, $b, and $c) will be created with sample data representing software inventory on different hosts.
- The Add-ItemToJsonReport function will be called three times to add each hashtable to the JSON report file.
- The Get-JsonReportFileContent function will be called to get the content of the JSON report file as a sorted JSON string.
- The Create-SoftwareInventoryHtmlReport function will be called to create an HTML report based on the JSON report file content. This function will generate an HTML table with rows for each unique software product and columns for each host, showing the versions of the software products installed on each host.
- Finally, the Get-HtmlReportFileContent function will be called to retrieve the content of the HTML report file.

After executing these commands, the HTML report file will be created (or updated) with a software inventory report in an HTML table format. The report will show the software product names, host names, and their respective installed software product versions. 

The code can be used as a starting point for creating your own software inventory reports.