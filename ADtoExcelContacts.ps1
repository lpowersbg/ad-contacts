# ENSURE YOU CHANGE DPATH TO THE SHEET DIRECTORY, OR USE "./" TO SEND TO CURRENT FOLDER
# Need AD module for Powershell https://www.microsoft.com/en-us/download/details.aspx?id=45520
# MUST HAVE THE IMPORT-EXCEL MODULE INSTALLED FOR POWERSHELL
# https://github.com/dfinke/ImportExcel
# Install-Module ImportExcel or download manually and copy to docs/powershell/mods
# Documentation is similar to https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv?view=powershell-7 

$dpath = "$PSScriptRoot"
$tpath = "$dpath\Template_Station Extension_Email.xlsx"
$props = @("SurName", "GivenName", "Department", "Title", "EmailAddress", "telephoneNumber")
$adgroup = "WGME-Users"
# Exception case for Stephanie Elston, appends her info at the end of the sheet
$elston = @'
 SurName, GivenName, Department, Title, EmailAddress, telephoneNumber,
 Elston,Stephanie,WPFO TRAFFIC,Copy Coordinator,seelston@sbgtv.com,(207) 228-7671,
'@ | ConvertFrom-Csv
$UpDate = Get-Date -Format "MM/dd/yyyy"
$FDate = "$dpath\$(Get-Date -Format MMddyy)_Station Extension_Email.xlsx"

# Copies the template file into a file with the day's date
Copy-Item $tpath -Destination $FDate

# Grabbing of base dataset and appending of Elston 
Get-ADGroupMember $adgroup | Get-ADUser -Properties $props | Select-Object $props | Where-Object {$_.SurName -ne "Elston" -and $_.EmailAddress -ne $null -and $_.Department -ne "STG HUB NSM" -and $_.Department -ne "DIELECTRIC BUSINESS"} | Export-Excel -Path $FDate -TableName "Staff" -Worksheet SORT -AutoSize -ClearSheet
Export-Excel -Path $FDate -InputObject $elston -Worksheet SORT -Append

# Manual title adjustment, updating of 'last modified.'
$ExShe = New-Object -ComObject Excel.Application
$Book=$ExShe.Workbooks.Open($FDate)
$Sheet = $Book.Sheets.Item("SORT")
$PSheet = $Book.Sheets.Item("PRINT")
$Sheet.Columns.Replace("WGME NEWS","News")
$Sheet.Columns.Replace("WGME BUSINESS","Business")
$Sheet.Columns.Replace("WGME ON-AIR OPERATIONS","Operations")
$Sheet.Columns.Replace("WPFO ON-AIR OPERATIONS","Operations")
$Sheet.Columns.Replace("WGME ENGINEERING","Engineering")
$Sheet.Columns.Replace("WPFO ENGINEERING","Engineering")
$Sheet.Columns.Replace("WGME DIGITAL INTERACTIVE", "Sales")
$Sheet.Columns.Replace("WGME GENERAL SALES", "Sales")
$Sheet.Columns.Replace("WPFO GENERAL SALES", "Sales")
$Sheet.Columns.Replace("WGME PRODUCTION", "Production")
$Sheet.Columns.Replace("WGME PROMOTION", "Promotion")
$Sheet.Columns.Replace("WPFO PROMOTION", "Promotion")
$Sheet.Columns.Replace("WPFO TRAFFIC", "Traffic")
$Sheet.Columns.Replace("Anchor/Reporter I", "Anchor/Reporter")
$Sheet.Columns.Replace("Anchor/ReporterI", "Anchor/Reporter")
$Sheet.Columns.Replace("Anchor/Reporter II", "Anchor/Reporter")
$Sheet.Columns.Replace("Assistant Chief Engineer II", "Assistant Chief Engineer")
$Sheet.Columns.Replace("Assistant Director, News", "Assistant News Director")
$Sheet.Columns.Replace("Assistant, News", "News Assistant")
$Sheet.Columns.Replace("Assistant, Sales", "Sales Assistant")
$Sheet.Columns.Replace("Chief Photographer I", "Chief Photographer")
$Sheet.Columns.Replace("Coordinator I, Digital Sales", "Digital Sales Coordinator")
$Sheet.Columns.Replace("Coordinator I, Human Resources", "Human Resources Coordinator")
$Sheet.Columns.Replace("Coordinator, Projects", "Projects Coordinator")
$Sheet.Columns.Replace("Director, Engineering", "Engineering Director")
$Sheet.Columns.Replace("Director, Operations", "Operations Director")
$Sheet.Columns.Replace("Executive Producer II", "Executive Producer")
$Sheet.Columns.Replace("Manager II, Marketing", "Marketing Manager")
$Sheet.Columns.Replace("Manager, Business", "Business Manager")
$Sheet.Columns.Replace("Manager, Creative Services", "Creative Services Manager")
$Sheet.Columns.Replace("Manager, Multimedia", "Multimedia Manager")
$Sheet.Columns.Replace("Managing Editor I", "Managing Editor")
$Sheet.Columns.Replace("Marketing Consultant Sales", "Marketing Consultant")
$Sheet.Columns.Replace("Meteorologist I", "Meteorologist")
$Sheet.Columns.Replace("MeteorologistI", "Meteorologist")
$Sheet.Columns.Replace("Meteorologist II", "Meteorologist")
$Sheet.Columns.Replace("Multimedia Journalist I", "Multimedia Journalist")
$Sheet.Columns.Replace("News Anchor II", "News Anchor")
$Sheet.Columns.Replace("Newscast Director I", "Newscast Director")
$Sheet.Columns.Replace("Newscast Producer I", "Newscast Producer")
$Sheet.Columns.Replace("Photographer I", "Photographer")
$Sheet.Columns.Replace("Producer I, Creative Services", "Creative Services Producer")
$Sheet.Columns.Replace("Producer I, Digital Content", "Digital Content Producer")
$Sheet.Columns.Replace("Producer I, Production", "Production Producer")
$Sheet.Columns.Replace("Producer I, Promotion", "Promotion Producer")
$Sheet.Columns.Replace("Reporter I, General Assignment", "General Assignment Reporter")
$Sheet.Columns.Replace("Technician I, Engineering", "Engineering Technician")
$Sheet.Columns.Replace("Technician, Operations", "Operations Technician")
$Sheet.Columns.Replace("Video Editor I", "Video Editor")
$Sheet.Columns.Replace("VP, General Manager", "General Manager")
$Sheet.Columns.Replace("Editor I, Promotion", "Promotion Editor")
$PSheet.Cells.Item(19,6) = "Updated $UpDate"


$used = $Sheet.usedRange 
$lastCell = $used.SpecialCells(11) 
$row = $lastCell.row
for ($i = 1; $i -lt $row; $i++) {
if ($Sheet.cells.Item($i,5).Value() -eq "jbrogers@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Brad"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "anpoulin@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Nick"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "jbowne@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Jim"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "PYoung@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Pam"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "bbarton@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Bill"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "wagayer@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Ritt"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "ksehobbs@sbgtv.com") {$Sheet.Cells.Item($i,2) = "KiKi"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "sacross@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Gus"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "tlham@sbgtv.com") {$Sheet.Cells.Item($i,4) = "WPFO Station Manager"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "mmeyer@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Mal"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "krasor@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Kerry"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "jdchrisos@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Jon"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "kdegroot@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Kate"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "loconnor@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Lexie"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "jpeterson@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Jeff"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "loconnor@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Lexie"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "rcsamowitz@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Becky"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "loconnor@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Lexie"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "ksampson@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Katie"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "thomer@sbgtv.com") {$Sheet.Cells.Item($i,2) = "Ted"}
elseif ($Sheet.cells.Item($i,5).Value() -eq "scbuccheri@sbgtv.com") {
    $Sheet.Cells.Item($i,1) = ""
    $Sheet.Cells.Item($i,2) = ""
    $Sheet.Cells.Item($i,3) = ""
    $Sheet.Cells.Item($i,4) = ""
    $Sheet.Cells.Item($i,5) = ""
    $Sheet.Cells.Item($i,6) = ""
    }
}

[void]$Book.save()
[void]$Book.close()
[void]$ExShe.quit()
