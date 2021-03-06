$spath = "C:\Users\lpower\OneDrive - Sinclair Broadcast Group, Inc\!Profile\Documents\adtest\110920_Station Extension_Email - Testing.xlsx"
$props = @("SurName", "GivenName", "Department", "Title", "EmailAddress", "telephoneNumber")
$adgroup = "WGME-Users"
$elston = @'
 SurName, GivenName, Department, Title, EmailAddress, telephoneNumber,
 Elston,Stephanie,WPFO TRAFFIC,Copy Coordinator,seelston@sbgtv.com,(207) 228-7671,
'@ | ConvertFrom-Csv
#$users = Get-ADGroupMember $adgroup | Get-ADUser -Properties $props | Select-Object $props 
#$users | Where-Object ($_.SurName -eq "Elston") | Set-ItemProperty -Name "Title" -Value "Copy Coordinator"
#$users | Export-Excel -Path $spath -Worksheet TestSheet -AutoSize
#Write-Output $users

#$elston | Export-Excel -Path $spath -Worksheet TestSheet -TableName "Staff" -Append -AutoSize -ClearSheet 
Get-ADGroupMember $adgroup | Get-ADUser -Properties $props | Select-Object $props | Where-Object {$_.SurName -ne "Elston" -and $_.EmailAddress -ne $null -and $_.Department -ne "STG HUB NSM" -and $_.Department -ne "DIELECTRIC BUSINESS"} | Export-Excel -Path $spath -TableName "Staff" -Worksheet SORT -AutoSize -ClearSheet
Export-Excel -Path $spath -InputObject $elston -Worksheet SORT -Append

$ExShe = New-Object -ComObject Excel.Application
$Book=$ExShe.Workbooks.Open($spath)
$Sheet = $Book.Sheets.Item("SORT")
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
[void]$Book.save()
[void]$Book.close()
[void]$ExShe.quit()
#[Runtime.Interopservices.Marshal]::ReleaseComObject( $excel_obj ) >$null

#$test = Get-ADGroupMember $adgroup | Get-ADUser -Properties $props | Select-Object $props | Where-Object {$_.EmailAddress -ne $null -and $_.Department -ne "STG HUB NSM" -and $_.Department -ne "DIELECTRIC BUSINESS"}
#$test | Where-Object {$_.SurName -eq "Elston"} | $_.Title = "Copy Coordinator"
#$test | Where-Object {$_.SurName -eq "Elston"} | Write-Output