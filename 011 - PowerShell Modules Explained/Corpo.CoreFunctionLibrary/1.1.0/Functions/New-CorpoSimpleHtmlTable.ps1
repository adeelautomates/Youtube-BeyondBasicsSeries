<#
.SYNOPSIS
Generates a simple HTML table from a PowerShell object.

.DESCRIPTION
The New-CorpoSimpleHtmlTable function takes one or more PSCustomObjects
and converts them into a formatted HTML <table> string.
Each property of the input object becomes a column, and each
object instance becomes a row.  
Useful for creating clean, readable HTML Tables

.PARAMETER Object
One or more PowerShell objects (PSCustomObject) whose properties
will be used as table columns.

.EXAMPLE
$data = Get-Process | Select-Object Name, CPU, Id -First 5
$html = New-CorpoSimpleHtmlTable -Object $data
$html | Out-File "C:\temp\process.html"

Generates an HTML table of the first five running processes.

.EXAMPLE
$resources = Get-AzResource | Select-Object Name, ResourceGroupName, Location
$table = New-CorpoSimpleHtmlTable -Object $resources
$body  = "<b>Azure Resources:</b><br><br>$table"

Creates an HTML fragment for embedding resource data in an email.

.OUTPUTS
[String]  
Returns a string containing the HTML markup for the table.

.NOTES
Author: Adeel Anwar  
Version: 1.0  
Intended for use in scripts or runbooks that send HTML reports or email notifications.
#>
function New-CorpoSimpleHtmlTable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [pscustomobject]$object # only accepts pscustomobjects to make the table
    )
    # Form column using the object's properties
    $columns = $object[0].PSObject.Properties.Name 
    $htmlTHData = "" 
    $htmlTHData = ($columns | ForEach-Object {"        <th>$_</th>"}) -join "`n"
$htmlTableColumns = @"
    <tr align="left" style="background-color: #e7e7e7; color: #000000; padding: 10px; font-size: 16px">
$htmlTHData
    </tr>
"@
    # Form rows using the properties data
    $htmlTableRows = ""
    foreach($item in $object){
        $htmlTableRow = "    <tr style='border-top: 1px solid #e7e7e7; border-bottom: 1px solid #e7e7e7; padding: 10px;'>`n"
        foreach($column in $columns){
            $value = $item.$column
            $htmlTableRow += "      <td>$value</td>`n"
        }
        $htmlTableRow += "    </tr>`n"
        $htmlTableRows += $htmlTableRow
    }
    $htmlTableRows = $htmlTableRows.TrimEnd("`n")
    # Put it all in a table
$htmlTable = @"
<table cellpadding="10" cellspacing="0" border="0" style="border-collapse: collapse; width:auto; display:inline-table; font-size: 14px; font-family: Calibri, Arial, sans-serif;">
$htmlTableColumns
$htmlTableRows
</table>
"@
    # Return the HTML table
    return $htmlTable
}