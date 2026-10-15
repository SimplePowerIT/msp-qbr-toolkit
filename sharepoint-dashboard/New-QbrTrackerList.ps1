<#
.SYNOPSIS
  Creates the "QBR Tracker" list that the QBR Dashboard web part reads and writes.

.DESCRIPTION
  Requires PnP.PowerShell (Install-Module PnP.PowerShell -Scope CurrentUser) and an Entra app
  registered for PnP interactive sign-in (Register-PnPEntraIDAppForInteractiveLogin).
  Safe to re-run: existing fields are skipped.

.EXAMPLE
  .\New-QbrTrackerList.ps1 -SiteUrl https://yourco.sharepoint.com/sites/Operations -ClientId <app id>
#>
param(
  [Parameter(Mandatory)] [string] $SiteUrl,
  [Parameter(Mandatory)] [string] $ClientId,
  [string] $ListTitle = "QBR Tracker"
)
$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

if (-not (Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue)) {
  New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch | Out-Null
  Write-Host "Created list $ListTitle"
}
Set-PnPField -List $ListTitle -Identity Title -Values @{ Title = "Client" } | Out-Null

function Add-Field($xml, $name) {
  if (Get-PnPField -List $ListTitle -Identity $name -ErrorAction SilentlyContinue) { Write-Host "  $name exists"; return }
  Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml | Out-Null
  Write-Host "  added $name"
}
function Choices($list) { ($list | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join "" }

Add-Field "<Field Type='Choice' Name='Priority' StaticName='Priority' DisplayName='Priority' Format='Dropdown'><Default>Primary</Default><CHOICES>$(Choices 'Primary','Secondary','As Needed')</CHOICES></Field>" Priority
Add-Field "<Field Type='Number' Name='Frequency' StaticName='Frequency' DisplayName='Every (years)' Decimals='0' Min='1' />" Frequency
Add-Field "<Field Type='Choice' Name='PreferredWindow' StaticName='PreferredWindow' DisplayName='Preferred window' Format='Dropdown'><CHOICES>$(Choices 'Q1','Q2','Q3','Q4')</CHOICES></Field>" PreferredWindow
Add-Field "<Field Type='Choice' Name='QBRStatus' StaticName='QBRStatus' DisplayName='Status' Format='Dropdown'><Default>Open</Default><CHOICES>$(Choices 'Open','Contacted','No response','Scheduled','Completed','Declined','As needed')</CHOICES></Field>" QBRStatus
foreach ($d in @(
    @('LastCompleted', 'Last completed'), @('PlannedDue', 'Planned due'), @('ScheduledDate', 'Scheduled date'),
    @('ManualDue', 'Manual due (override)'), @('LastContact', 'Last contact'))) {
  Add-Field "<Field Type='DateTime' Name='$($d[0])' StaticName='$($d[0])' DisplayName='$($d[1])' Format='DateOnly' />" $d[0]
}
Add-Field "<Field Type='Text' Name='Plan' StaticName='Plan' DisplayName='Plan' />" Plan
Add-Field "<Field Type='Note' Name='QBRNotes' StaticName='QBRNotes' DisplayName='Notes' NumLines='6' RichText='FALSE' />" QBRNotes
Add-Field "<Field Type='Text' Name='SPFolder' StaticName='SPFolder' DisplayName='SharePoint folder' />" SPFolder

# Next due: manual override, else scheduled date, else planned due rolled forward past the last completed review.
$formula = '=IF(ISBLANK([Manual due (override)]),IF(ISBLANK([Scheduled date]),IF(ISBLANK([Planned due]),IF(OR(ISBLANK([Last completed]),ISBLANK([Every (years)])),"",DATE(YEAR([Last completed])+[Every (years)],MONTH([Last completed]),DAY([Last completed]))),IF(OR(ISBLANK([Last completed]),[Planned due]&gt;[Last completed]),[Planned due],IF(ISBLANK([Every (years)]),"",DATE(YEAR([Planned due])+[Every (years)]*(INT((YEAR([Last completed])-YEAR([Planned due]))/[Every (years)])+1),MONTH([Planned due]),DAY([Planned due]))))),[Scheduled date]),[Manual due (override)])'
Add-Field "<Field Type='Calculated' Name='NextDue' StaticName='NextDue' DisplayName='Next due' ResultType='DateTime' Format='DateOnly'><Formula>$formula</Formula><FieldRefs><FieldRef Name='ManualDue'/><FieldRef Name='ScheduledDate'/><FieldRef Name='PlannedDue'/><FieldRef Name='LastCompleted'/><FieldRef Name='Frequency'/></FieldRefs></Field>" NextDue

Write-Host "Done. Upload qbr-dashboard.sppkg to the app catalog and add the QBR Dashboard web part to a page on $SiteUrl."
