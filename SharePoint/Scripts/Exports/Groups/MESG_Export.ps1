<#
Developed by: Rhys Saddul
#>

# Install and import the AzureAD module if not already installed
if (-not (Get-Module AzureAD -ErrorAction SilentlyContinue)) {
    Install-Module -Name AzureAD -Force -AllowClobber
}
Import-Module AzureAD

# Prompt to ensure variables have been set up correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the variable for the output file path?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Variable for the output file path
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\Groups\MESG_Export.csv"

# Connect to Azure AD
Connect-AzureAD

# Create an array to store group members
$membersArray = @()

# Get all mail-enabled security groups
$mailEnabledGroups = Get-AzureADGroup -Filter "SecurityEnabled eq true and MailEnabled eq true"

foreach ($group in $mailEnabledGroups) {
    $groupName = $group.DisplayName
    $groupEmailAddress = $group.MailNickname + "@" + (Get-AzureADDomain | Where-Object {$_.IsDefault -eq $true}).Name
    
    # Get members of the mail-enabled security group
    try {
        $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true
        
        foreach ($member in $members) {
            $membersArray += [PSCustomObject]@{
                "Group Name" = $groupName
                "Group Email Address" = $groupEmailAddress
                "Member Name" = $member.DisplayName
                "Member Email Address" = $member.Mail
            }
        }
    }
    catch {
        Write-Error "Error retrieving members for mail-enabled security group: $groupName. $_"
    }
}

# Export mail-enabled security group members to CSV
$membersArray | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Mail-enabled security group members exported to $outputFilePath"

# Disconnect from Azure AD
Disconnect-AzureAD
