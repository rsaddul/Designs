<#
Developed by: Rhys Saddul
#>

# Install and import the ExchangeOnlineManagement module if not already installed
if (-not (Get-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement

# Prompt to ensure variables have been set up correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the variable for the output file path?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Variable for the output file path
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\Groups\M365_Export.csv"

# Connect to Exchange Online with progress display
Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false

# Create an array to store group members
$membersArray = @()

# Get all Microsoft 365 groups
$groups = Get-UnifiedGroup -ResultSize Unlimited

foreach ($group in $groups) {
    $groupName = $group.DisplayName
    $groupEmailAddress = $group.PrimarySmtpAddress
    
    # Get members of the Microsoft 365 group
    try {
        $members = Get-UnifiedGroupLinks -Identity $groupEmailAddress -LinkType Members -ResultSize Unlimited -ErrorAction Stop
        
        foreach ($member in $members) {
            $membersArray += [PSCustomObject]@{
                "Group Name" = $groupName
                "Group Email Address" = $groupEmailAddress
                "Member Name" = $member.DisplayName
                "Member Email Address" = $member.PrimarySmtpAddress
            }
        }
    }
    catch {
        Write-Error "Error retrieving members for Microsoft 365 group: $groupName. $_"
    }
}

# Export Microsoft 365 group members to CSV
$membersArray | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Microsoft 365 group members exported to $outputFilePath"
