#https://www.c-sharpcorner.com/article/sharepoint-framework-deploy-spfx-webparts-to-office-365-public-cdn/
Connect-SPOService -Url https://zaindev-admin.sharepoint.com 
Set-SPOTenantCdnEnabled -CdnType Public
Add-SPOTenantCdnOrigin -CdnType Public -OriginUrl sites/cdn/cdn
#"https://zaindev.sharepoint.com/sites/cdn/cdn/ExpenseClaimsWebPart"
https://publiccdn.sharepointonline.com/zaindev.sharepoint.com/sites/cdn/cdn/expenseclaimswebpart