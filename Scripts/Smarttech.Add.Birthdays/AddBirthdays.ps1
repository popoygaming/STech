$siteUrl = "https://sitechwa2.sharepoint.com/sites/STA-IntranetPortal"

# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -ClientId "8a4183a4-23fb-4de2-9191-9d7ca963f208" -Interactive

# Get web details including title
$web = Get-PnPWeb -Includes Title, Url

Write-Host "Site URL: $($web.Url)"
Write-Host "Site Title: $($web.Title)"