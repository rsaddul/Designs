<#
Developed: by Rhys Saddul

This script is tailored to export the file and folder structure within a library on a SharePoint site.

#>

# Prompt to ensure varibles have been setup correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the varibles?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

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

# Varibles
$SiteURL = "https://eduthingazurelab.sharepoint.com/sites/MAT-HUB"
$ListName = "Policies"
$ReportOutput = "C:\$ListName.csv"

# Connect to SharePoint Online site
Connect-PnPOnline $SiteURL -Interactive

# Array to store results
$Results = @()

# Get all Items from the document library
$List  = Get-PnPList -Identity $ListName
$ListItems = Get-PnPListItem -List $ListName -PageSize 500
Write-host "Total Number of Items in the List:" $List.ItemCount -ForegroundColor Yellow

$ItemCounter = 0 
# Iterate through each item
Foreach ($Item in $ListItems)
{
    $Results += New-Object PSObject -Property ([ordered]@{
        Name              = $Item["FileLeafRef"]
        Type              = $Item.FileSystemObjectType
        FileType          = $Item["File_x0020_Type"]
        RelativeURL       = $Item["FileRef"]
        CreatedBy         = $Item["Author"].Email
        CreatedOn         = $Item["Created"]
        ModifiedBy        = $Item["Editor"].Email
        ModifiedOn        = $Item["Modified"]
    })

    $ItemCounter++
    Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Getting Metadata from Item '$($Item['FileLeafRef'])'"         
}

# Export the results to CSV
$Results | Select-Object Name, Type, FileType, RelativeURL, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn | Export-Csv -Path $ReportOutput -NoTypeInformation
Write-host "$ListName Exported to CSV Successfully!" -ForegroundColor Green
