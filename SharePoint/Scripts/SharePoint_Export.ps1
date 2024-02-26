Write-Host "Make sure you are using Powershell 7.2 or later" -foreground Yellow

# Install SharePoint PnP PowerShell module
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Continue

# Import SharePoint PnP PowerShell module
Import-Module PnP.PowerShell -DisableNameChecking

# Connect to SharePoint Online
Connect-PnPOnline -Url "https://wlfschool-admin.sharepoint.com/" -UseWebLogin

# Define output file path
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\KST\SharePointSites.csv"

# Get all SharePoint sites
$sites = Get-PnPTenantSite -Detailed

# Create an array to store SharePoint sites information
$sitesInfo = @()

foreach ($site in $sites) {
    $siteName = $site.Title
    $siteUrl = $site.Url
    $siteTemplate = $site.Template
    $siteLastActivity = $site.LastContentModifiedDate
    $siteStorageUsed = $site.StorageUsageCurrent

    $sitesInfo += [PSCustomObject]@{
        "Site Name" = $siteName
        "Site URL" = $siteUrl
        "Template" = $siteTemplate
        "Last Activity" = $siteLastActivity
        "Storage Used (MB)" = $siteStorageUsed
    }
}

# Export SharePoint sites information to CSV
$sitesInfo | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "SharePoint sites information exported to $outputFilePath" -ForegroundColor Green
