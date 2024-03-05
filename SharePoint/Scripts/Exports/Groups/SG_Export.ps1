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
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\BFLT\SG_Export.csv"

# Connect to Azure AD
Connect-AzureAD

# Get all security groups
$securityGroups = Get-AzureADGroup -Filter "SecurityEnabled eq true and MailEnabled eq false"

# Create an array to store group members
$membersArray = @()

# Iterate through each security group
foreach ($group in $securityGroups) {
    $groupName = $group.DisplayName
    
    # Get members of the security group
    try {
        $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true -ErrorAction Stop
        
        foreach ($member in $members) {
            $membersArray += [PSCustomObject]@{
                "Group Name" = $groupName
                "Member Name" = $member.DisplayName
                "Member Email Address" = $member.Mail
            }
        }
    }
    catch {
        Write-Error "Error retrieving members for security group: $groupName. $_"
    }
}

# Export security group members to CSV
$membersArray | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Security group members exported to $outputFilePath"

# Disconnect from Azure AD
Disconnect-AzureAD