# Install AzureAD module if not already installed
<#
if (-not (Get-Module -Name "AzureAD" -ListAvailable)) {
    Install-Module -Name "AzureAD" -Force -Scope CurrentUser
}
#>

# Connect to Azure Active Directory
#Connect-AzureAD

# Define the school code and its alias
$schoolCode = "EDU"
$schoolCodeAlias = "HWK"

# Function to check if a group exists
function GroupExists($displayName) {
    return (Get-AzureADGroup -Filter "DisplayName eq '$displayName'").Count -gt 0
}

# Function to check if a member exists in a group
function MemberExists($groupId, $memberId) {
    return (Get-AzureADGroupMember -ObjectId $groupId | Where-Object {$_.ObjectId -eq $memberId}).Count -gt 0
}

# Create Static Security Groups
$staticGroups = @(
    "[EDU v2 Security] - $schoolCodeAlias Office Devices",
    "[EDU v2 Security] - $schoolCodeAlias Teacher Devices",
    "[EDU v2 Security] - All Corporate $schoolCodeAlias Windows Devices",
    "[EDU v2 Security] - $schoolCodeAlias Assigned or Shared Windows Intune Devices"
)

foreach ($groupName in $staticGroups) {
    # Check if the group already exists
    if (GroupExists($groupName)) {
        Write-Warning "Group '$groupName' already exists. Skipping creation."
        continue
    }

    # Create Static Group
    $newGroup = New-AzureADGroup -DisplayName $groupName -SecurityEnabled $true -MailEnabled $false -MailNickName $false
    Write-Host "Static Group '$groupName' created successfully." -ForegroundColor Green
}

# Create Dynamic Security Groups
$dynamicGroups = @(
    "[EDU v2 Security] - $schoolCodeAlias Office Desktops",
    "[EDU v2 Security] - $schoolCodeAlias Office Laptops",
    "[EDU v2 Security] - $schoolCodeAlias Teacher Desktops",
    "[EDU v2 Security] - $schoolCodeAlias Teacher Laptops",
    "[EDU v2 Security] - All Personal Owned Non Corporate Devices"
)

foreach ($groupName in $dynamicGroups) {
    # Check if the group already exists
    if (GroupExists($groupName)) {
        Write-Warning "Group '$groupName' already exists. Skipping creation."
        continue
    }

    $membershipRule = $null

    # Define dynamic membership rule based on group name
    switch ($groupName) {
        "[EDU v2 Security] - $schoolCodeAlias Office Desktops" {
            $membershipRule = "(device.displayName -startsWith ""$schoolCodeAlias-OD-"")"
        }
        "[EDU v2 Security] - $schoolCodeAlias Office Laptops" {
            $membershipRule = "(device.displayName -startsWith ""$schoolCodeAlias-OL-"")"
        }
        "[EDU v2 Security] - $schoolCodeAlias Teacher Desktops" {
            $membershipRule = "(device.displayName -startsWith ""$schoolCodeAlias-TD-"")"
        }
        "[EDU v2 Security] - $schoolCodeAlias Teacher Laptops" {
            $membershipRule = "(device.displayName -startsWith ""$schoolCodeAlias-TL-"")"
        }
        "[EDU v2 Security] - All Personal Owned Non Corporate Devices" {
            $membershipRule = '(device.deviceOwnership -eq "Personal") -or (device.deviceOwnership -eq "Unknown")'
        }

        default {
            Write-Warning "Unsupported group name: $groupName"
            continue
        }
    }

    # Create Dynamic Group
    $newGroup = New-AzureADMSGroup -DisplayName $groupName -GroupTypes "DynamicMembership" `
                                   -SecurityEnabled $true -MailEnabled $false -MailNickname $false `
                                   -MembershipRule $membershipRule -MembershipRuleProcessingState "On"
    Write-Host "Dynamic Group '$groupName' created successfully with membership rule: $membershipRule" -ForegroundColor Green
}

# Get ObjectIds for the target groups
$allCorporateDevicesGroupId = (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - All Corporate $schoolCodeAlias Windows Devices'").ObjectId
$officeDevicesGroupId = (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - $schoolCodeAlias Office Devices'").ObjectId
$teacherDevicesGroupId = (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - $schoolCodeAlias Teacher Devices'").ObjectId

# Function to add a member to a group if they don't already exist
function AddMemberIfNotExists($groupId, $memberId) {
    if (MemberExists $groupId $memberId) {
        Write-Warning "Member already exists in the group. Skipping add"
    } else {
        Add-AzureADGroupMember -ObjectId $groupId -RefObjectId $memberId
        Write-Host "Added member to group." -ForegroundColor Yellow
    }
}

# Add members to specified groups
Write-Host "Adding members to specified groups..." -ForegroundColor Cyan

AddMemberIfNotExists $allCorporateDevicesGroupId $officeDevicesGroupId

AddMemberIfNotExists $allCorporateDevicesGroupId $teacherDevicesGroupId

AddMemberIfNotExists $officeDevicesGroupId (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - $schoolCodeAlias Office Desktops'").ObjectId

AddMemberIfNotExists $officeDevicesGroupId (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - $schoolCodeAlias Office Laptops'").ObjectId

AddMemberIfNotExists $teacherDevicesGroupId (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - $schoolCodeAlias Teacher Desktops'").ObjectId

AddMemberIfNotExists $teacherDevicesGroupId (Get-AzureADGroup -Filter "DisplayName eq '[EDU v2 Security] - $schoolCodeAlias Teacher Laptops'").ObjectId

# Disconnect from Azure AD
#Disconnect-AzureAD
