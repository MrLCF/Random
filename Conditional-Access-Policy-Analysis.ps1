$report  = "C:\Users\lysander.fernandes\Desktop\initializer.html"
$date = Get-Date
$dateval = [string]$date.Day + '-' + [string]$date.Month + '-' + [string]$date.Year + '-' + [string]$date.Hour + '-' + [string]$date.Minute + '-' + [string]$date.Second
$path = $env:USERPROFILE + '\Desktop\Conditional_Access_policy_Analysis_' + $dateval + '.xlsx'
$XL = New-Object -comobject Excel.Application
 
$XL.Visible = $True
 
$WB = $XL.Workbooks.Add()
$WS = $WB.Worksheets.Item(1)

$WS.Cells.Item(1,1)  = "ID"
$WS.Cells.Item(1,2)  = "DisplayName"
$WS.Cells.item(1,3)  = "State"
$WS.Cells.item(1,4)  = "Included Apps"
$WS.Cells.item(1,5)  = "Excluded Apps"
$WS.Cells.item(1,6)  = "Include User Actions"
$WS.Cells.item(1,7)  = "Include Protection Levels"
$WS.Cells.item(1,8)  = "Include Users"
$WS.Cells.item(1,9)  = "Exclude Users"
$WS.Cells.item(1,10) = "Include Groups"
$WS.Cells.item(1,11) = "Exclude Groups"
$WS.Cells.item(1,12) = "Include Roles"
$WS.Cells.item(1,13) = "Exclude Roles"
$WS.Cells.item(1,14) = "Platforms"
$WS.Cells.item(1,15) = "Included Locations"
$WS.Cells.item(1,16) = "Excluded Locations"
$WS.Cells.item(1,17) = "SignInRiskLevels"
$WS.Cells.item(1,18) = "ClientAppTypes"
$WS.Cells.item(1,19) = "Grant Controls"
$WS.Cells.item(1,20) = "BuiltIn Controls"
$WS.Cells.item(1,21) = "Custom Authentication Factors"
$WS.Cells.item(1,22) = "Terms Of Use"
$WS.Cells.item(1,23) = "ApplicationEnforcedRestrictions"
$WS.Cells.item(1,24) = "CloudAppSecurity"
$WS.Cells.item(1,25) = "SignInFrequency"
$WS.Cells.item(1,26) = "PersistentBrowser"

Function Applications {
$incapps = $applications.IncludeApplications
$excapps = $applications.ExcludeApplications
$incusrs = $applications.IncludeUserActions
$prlevls = $applications.IncludeProtectionLevels

$apps = ''
foreach($incapp in $incapps){

if($incapp -ne 'All'){
try{
#$value = Get-AzureADServicePrincipal -ObjectId $incapp
$value = $appsf|Where-Object {$_.appid -eq $incapp}
$apps += $value.DisplayName + ';'
$apps += '
'
}
catch{
$apps += $incapp + ';'
$apps += '
'
}
}
else{
$apps = $incapp
}

}
if($apps -ne $null){
$WS.Cells.Item($counter,4) = $apps
}
else{
$WS.Cells.Item($counter,4) = 'No-Apps'
}
$apps = ''
if($excapps.count -ne 0){
foreach($excapp in $excapps){

try{
#$value = Get-AzureADServicePrincipal -ObjectId $excapp
$value = $appsf|Where-Object {$_.appid -eq $excapp}
$apps += $value.DisplayName + ';'
$apps += '
'
}
catch{
$apps += $excapp + ';'
$apps += '
'
}
}
}
else{
$apps = 'No-Apps'
}
$WS.Cells.Item($counter,5) = $apps
if($incusrs -ne $null){
$WS.Cells.Item($counter,6) = $incusrs
}
if($incusrs.count -eq 0){
$WS.Cells.Item($counter,6) = 'N/A'
}
if($prlevls -ne $null){
$WS.Cells.Item($counter,7) = $prlevls
}
if($prlevls -eq $null){
$WS.Cells.Item($counter,7) = 'N/A'
}
}

Function Users {

$incurs = $Users.IncludeUsers
$ingrps = $Users.IncludeGroups
$excurs = $Users.ExcludeUsers
$exgrps = $Users.ExcludeGroups
$incrls = $Users.IncludeRoles
$excrls = $Users.ExcludeRoles

$included_users = ''
if($incurs.count -ne 0){
foreach($incur in $incurs){
try{
$included_users += (Get-AzureADUser -ObjectId $incur).userprincipalname + ';'
$included_users += '
'
}
catch{
$included_users += $incur + ';'
$included_users += '
'
}
}
}
if($incurs -eq 'All'){
$included_users = 'All'
}
if($incurs.Count -eq 0){
$included_users = 'No-Users'
}

$included_groups = ''
if($ingrps.count -ne 0){
foreach($ingrp in $ingrps){
try{
$included_groups += (Get-AzureADGroup -ObjectId $ingrp).Displayname + ';'
$included_groups += '
'
}
catch{
$included_groups += $ingrp + ';'
$included_groups += '
'
}
}
}
if($ingrps.count -eq 0){
$included_groups = 'No-Groups'
}

$included_roles = ''
if($incrls.count -ne 0){
foreach($incrl in $incrls){
try{
$included_roles += (Get-AzureADDirectoryRole -ObjectId $incrl).Displayname + ';'
$included_roles +='
'
}
catch{
$included_roles += $incrl + ';'
$included_roles +='
'
}
}
}
if($incrls.count -eq 0)
{
$included_roles = 'No-Roles'
}
$excluded_users = ''
if($excurs.count -ne 0){
foreach($excur in $excurs){
try{
$excluded_users += (Get-AzureADUser -ObjectId $excur).DisplayName + ';'
$excluded_users += '
'
}
catch{
$excluded_users += $excur + ';'
$excluded_users += '
'
}
}
}
if($excurs.count -eq 0){
$excluded_users = 'No-Users'
}
$excluded_groups = ''
if($exgrps.count -ne 0){
try{
foreach($exgrp in $exgrps){
$excluded_groups += (Get-AzureADGroup -ObjectId $exgrp).DisplayName + ';'
$excluded_groups += '
'
}
}
catch{
$excluded_groups += $exgrp + ';'
$excluded_groups += '
'
}
}

if($exgrps.count -eq 0){
$excluded_groups = 'No-Groups'
}
$excluded_roles = ''
if($excrls -ne $null){
foreach($excrl in $excrls){
try{
$excluded_roles += (Get-AzureADDirectoryRole -ObjectId $excrl).DisplayName + ';'
$excluded_roles += '
'
}
catch{
$excluded_roles += $excrl + ';'
$excluded_roles += '
'
}
}
}
if($excrls.count -eq 0){
$excluded_roles = 'No-Roles'
}
Start-Sleep -Seconds 1
$WS.Cells.Item($counter,8)  =  $included_users
$WS.Cells.Item($counter,9)  =  $excluded_users
$WS.Cells.Item($counter,10) =  $included_groups
$WS.Cells.Item($counter,11) =  $excluded_groups
$WS.Cells.Item($counter,12) =  $included_roles
$WS.Cells.Item($counter,13) =  $excluded_roles
}


Function Locations {

$incls = $Locations.IncludeLocations
$excls = $Locations.ExcludeLocations

$excluded_locations = ''
if($excls.count -ne 0){
foreach($excl in $excls){
try{
$excluded_locations += (Get-AzureADMSNamedLocationPolicy -PolicyId $excl).DisplayName + ';'
$excluded_locations += '
'
}
catch{
$excluded_locations += $excl + ';'
$excluded_locations += '
'
}
}
}
if($excls.count -eq 0){
$excluded_locations = 'No-Exclusions'
}
$included_locations = ''
if($incls.count -ne 0){
foreach($incl in $incls){
try{
$included_locations += (Get-AzureADMSNamedLocationPolicy -PolicyId $incl).DisplayName + ';'
$included_locations += '
'
}
catch{
$included_locations += $incl + ';'
$included_locations += '
'
}
}
}
if($incls -eq 'All'){
$included_locations = 'All'
}
if($incls.count -eq 0){
$included_locations = 'No-Locations'
}
$WS.Cells.Item($counter,15) =  $included_locations
$WS.Cells.Item($counter,16) =  $excluded_locations
}


$gc_values = 'Operator Value: ' + $GrantControls._Operator




Connect-AzureAD

$appsf = Get-AzureADServicePrincipal -All $true
$policies = Get-AzureADMSConditionalAccessPolicy |Select-Object *
$counter = 1

foreach($policy in $policies){
$counter += 1
$applications = $policy.Conditions.Applications
$Users = $policy.Conditions.Users
$Platforms = $policy.Conditions.Platforms
$Locations = $policy.Conditions.Locations
$SignInRiskLevels = $policy.Conditions.SignInRiskLevels
$ClientAppTypes = $policy.Conditions.ClientAppTypes
$GrantControls = $policy.GrantControls
$SessionControls = $policy.SessionControls

$WS.Cells.Item($counter,1) = $policy.Id
$WS.Cells.Item($counter,2) = $policy.DisplayName
$WS.Cells.Item($counter,3) = $policy.State
Applications
Users
$platformsin = [string]$Platforms.IncludePlatforms
$platformsout = [string]$Platforms.ExcludePlatforms
if($platformsout -eq ''){
$platformsout = 'Null'
}
$platformsin += ';Excluded :' + $platformsout
if($platformsin -ne $null){
$WS.Cells.Item($counter,14) = $platformsin
}
if($platformsin -eq $null){
$WS.Cells.Item($counter,14) = 'Null'
}
Locations
$srl = [string]$SignInRiskLevels
if($srl -ne ''){
$WS.Cells.Item($counter,17) = [string]$SignInRiskLevels
}
if($srl -eq ''){
$WS.Cells.Item($counter,17) = 'Not-Configured'
}
$cat = [string]$ClientAppTypes
if($cat -ne ''){
$WS.Cells.Item($counter,18) = [string]$ClientAppTypes
}
if($cat -eq ''){
$WS.Cells.Item($counter,18) = 'Null'
}

$gc = $GrantControls._operator
if($gc -ne $null){
$WS.Cells.Item($counter,19) = $GrantControls._Operator
$WS.Cells.Item($counter,20) = [string]$GrantControls.BuiltInControls
}
else{
$WS.Cells.Item($counter,19) = 'N/A'
$WS.Cells.Item($counter,20) = 'N/A'
}
$gcf = $GrantControls.CustomAuthenticationFactors
if($gcf -ne $null){
$WS.Cells.Item($counter,21) = [string]$GrantControls.CustomAuthenticationFactors
}
else{
$WS.Cells.Item($counter,21) = 'N/A'
}
$term = $GrantControls.TermsOfUse
if($term -ne $null){
$WS.Cells.Item($counter,22) = [string]$GrantControls.TermsOfUse
}
else{
$WS.Cells.Item($counter,22) = 'N/A'
}
$appr = $SessionControls.ApplicationEnforcedRestrictions
if ($appr -ne $null){
$WS.Cells.Item($counter,23) = [string]$SessionControls.ApplicationEnforcedRestrictions.IsEnabled
}
else{
$WS.Cells.Item($counter,23) = 'False'
}
$cldsec = $SessionControls.CloudAppSecurity
if($cldsec -ne $null){
$WS.Cells.Item($counter,24) = [string]$SessionControls.CloudAppSecurity.CloudAppSecurityType + ';' + [string]$SessionControls.CloudAppSecurity.IsEnabled
}
else{
$WS.Cells.Item($counter,24) = 'N/A'
}
$sgnf = $SessionControls.SignInFrequency
if($signf -ne $null){
$WS.Cells.Item($counter,25) = [string]$SessionControls.SignInFrequency.Type + ';'+  [string]$SessionControls.SignInFrequency.Value + ';' + [string]$SessionControls.SignInFrequency.IsEnabled
}
else{
$WS.Cells.Item($counter,25) = 'N/A'
}
$pb = $SessionControls.PersistentBrowser
if($pb -ne $null){
$WS.Cells.Item($counter,26) = [string]$SessionControls.PersistentBrowser.Mode + ';' + [string]$SessionControls.PersistentBrowser.IsEnabled
}
else{
$WS.Cells.Item($counter,26) = 'N/A'
}
}


$wb.Sheets.Add()
$WS2 = $WB.Worksheets.Item(1)

$WS2.Cells.Item(1,1)  = "ID"
$WS2.Cells.Item(1,2)  = "DisplayName"
$WS2.Cells.item(1,3)  = "Ipranges"
$WS2.Cells.item(1,4)  = "IsTrusted"
$WS2.Cells.item(1,5)  = "CountriesAndRegios"
$WS2.Cells.item(1,6)  = "IncludeUnkownCountriesAndRegions"


$namedlocationsAll = (Get-AzureADMSNamedLocationPolicy |Select-Object *)
$counter = 1
foreach($namedlocation in $namedlocationsAll){
$counter += 1
$ipranges = ''
$WS2.Cells.Item($counter,1) = $namedlocation.Id
$WS2.Cells.Item($counter,2) = $namedlocation.DisplayName
$ips = ''
$ips = ([string]$namedlocation.IpRanges.Cidraddress).Replace(' ',';')



if($ips.Length -ne 0){
$WS2.Cells.Item($counter,3) = $ips
}
else{
$WS2.Cells.Item($counter,3) = 'N/A'
}
$WS2.Cells.Item($counter,4) = $namedlocation.IsTrusted

$countries = ([string]$namedlocation.CountriesAndregions).Replace(' ',';')
if($countries.Length -ne 0){
$WS2.Cells.Item($counter,5) = $countries
}
else{
$WS2.Cells.Item($counter,5) = 'N/A'
}
if($namedlocation.IncludeUnknownCountriesAndRegions -ne $null){
$WS.Cells.Item($counter,6) = $namedLocation.IncludeUnkownCountriesAndRegions
}
else{
$WS.Cells.Item($counter,6) = 'No'
}
}

$ws2.Name = 'Named Locations'
$ws = $wb.Worksheets.Item(2)
$ws.Name = 'Conditional Access Policies'


$wb.SaveAs($path)
