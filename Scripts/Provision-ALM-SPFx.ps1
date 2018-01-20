Write-Host "Process Started"
$sppkgPath = "C:\Users\joao1\Desktop\sp-dev-gdpr-activity-hub-master\sp-dev-gdpr-activity-hub-master\GDPRStarterKit\sharepoint\solution\gdpr-starter-kit.sppkg"
$appPackageName = $sppkgPath.Split("\")[-1]
$appName = "gdpr-starter-kit-client-side-solution"

$siteURL = "http://spfxgdpr.sharepoint.com/sites/apps/"

Write-Host "Connecting to: $siteURL"
Connect-PnPOnline -Url $siteURL 
Write-Host "Connected!"

Write-Host "Checking if app $appName exists on AppCatalog..."
$appAdded = Get-PnPApp | ? { $_.Title -eq $appName }
if ($appAdded -eq $null) {
    Write-Host "It does not exist. Adding $appName to AppCatalog..."
    Add-PnPApp -Path $sppkgPath
}else{
    Write-Host "It exists. Updating $appName..."
    Update-PnPApp -Identity $appAdded.id 
}

$appAdded = Get-PnPApp | ? { $_.Title -eq $appName }
if ($appAdded -ne $null) {		
    Write-Host "App added/updated with success. Publishing it to all site collections..."
    Publish-PnPApp -Identity $appAdded.id -SkipFeatureDeployment
    Write-Host "App published with success. It is now available to use on any site."
}

Write-Host "Process completed, press any key to close this window ..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")