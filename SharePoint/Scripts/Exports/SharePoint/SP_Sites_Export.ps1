<#
Developed by: Rhys Saddul
#>

# Check for correct PowerShell version
if ($PSVersionTable.PSVersion -lt [System.Version]::new(7, 2)) {
    Write-Host "Script requires PowerShell 7.2 or later." -ForegroundColor Red
}

# Uninstall legacy PnP PowerShell module if present
if (Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable) {
    Write-Host "Uninstalling legacy SharePointPnPPowerShellOnline module..."
    Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force
}

# Install and import PnP.PowerShell module if not installed
if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
    Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser -AllowClobber
    Write-Host "PnP.PowerShell module installed and imported successfully." -ForegroundColor Green
}

# Add-Type to load the PresentationFramework assembly for this to work in PowerShell 7.2 or later
Add-Type -AssemblyName PresentationFramework

# Prompt to ensure variables have been set up correctly prior to running the script
$dialogResult = [System.Windows.MessageBox]::Show("Have you set the variable for SharePoint Online?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Varibles
Connect-PnPOnline -Url "https://stmarysceprimaryschool-admin.sharepoint.com/" -Interactive # Admin SharePoint Online URL
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\SharePoint\SP_Sites_Export.csv" # Define output file path

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
