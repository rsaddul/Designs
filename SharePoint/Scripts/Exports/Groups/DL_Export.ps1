<#
Developed by: Rhys Saddul
#>

# Install and import the ExchangeOnlineManagement module if not already installed
if (-not (Get-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement

# Prompt to ensure varibles have been setup correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the varible for the output file path?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Varible for the output file path
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\Groups\DL_Export.csv"

# Connect to Exchange Online with progress display
Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false

# Create an array to store distribution list members
$membersArray = @()

# Get all distribution lists
$distributionLists = Get-DistributionGroup -ResultSize Unlimited

foreach ($distributionList in $distributionLists) {
    $distributionListName = $distributionList.DisplayName
    $distributionListEmailAddress = $distributionList.PrimarySmtpAddress
    
    # Get members of the distribution list
    try {
        $members = Get-DistributionGroupMember -Identity $distributionListEmailAddress -ResultSize Unlimited -ErrorAction Stop
        
        foreach ($member in $members) {
            $membersArray += [PSCustomObject]@{
                "Distribution List Name" = $distributionListName
                "Distribution List Email Address" = $distributionListEmailAddress
                "Member Name" = $member.DisplayName
                "Member Email Address" = $member.PrimarySmtpAddress
            }
        }
    }
    catch {
        Write-Error "Error retrieving members for distribution list: $distributionListName. $_"
    }
}

# Export distribution list members to CSV
$membersArray | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Distribution list members exported to $outputFilePath"