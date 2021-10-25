#Example to hide a content type using PnP PowerShell
#Installing the PnP PowerShell 
#Install-Module -Name SharePointPnPPowerShellOnline
$SiteURL = "https://yourtenant.sharepoint.com/sites/yoursite"
$ListName = "Contacts" # your list name
$ContentTypeName ="Task" # your content type name that you need to change the visibility
 
#Connect to Pnp Online
Connect-PnPOnline -Url $SiteURL -ClientId "your client id" -ClientSecret "your client secret"  #-UseWebLogin
  
#Get the Context
$Context = Get-PnPContext
  
#Get the content type from List
$ContentType = Get-PnPContentType -Identity $ContentTypeName -List $ListName
  
#Set content type to hidden
$ContentType.Hidden = $true
$ContentType.Update($False)
$Context.ExecuteQuery()