try
{

    # **********************************************
    # Provision SharePoint Online artifacts
    # **********************************************

    Write-Host "Creating artifacts on target site" $GroupSiteUrl


    # Connect to the target site
    Connect-PnPOnline -url "https://spfxgdpr.sharepoint.com/sites/gdpr/" -UseWebLogin

    # Provision taxonomy items, fields, content types, and lists
    Apply-PnPProvisioningTemplate -Path "C:\Users\joao1\Desktop\sp-dev-gdpr-activity-hub-master\sp-dev-gdpr-activity-hub-master\Scripts\GDPR-Activity-Hub-Information-Architecture-Full.xml" -Handlers Fields,ContentTypes,Lists,TermGroups

    # Provision workflows
    Apply-PnPProvisioningTemplate -Path "C:\Users\joao1\Desktop\sp-dev-gdpr-activity-hub-master\sp-dev-gdpr-activity-hub-master\Scripts\GDPR-Activity-Hub-Workflows.xml" -Handlers Workflows
    
}
catch 
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}
