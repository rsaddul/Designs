<#
Developed by: Rhys Saddul

Overview:

1: Define Variables:
                    $matCodeAlias: A string representing a code alias (presumably related to a specific organization or department).
                    $siteCodes: An array of strings representing site codes.
                    $Type: A string representing a type (likely used for naming groups).
                    $Domain: A string representing a domain name.
                    $AdminAccount: A string representing an administrative account.

2: Define Group Names and Descriptions using Arrays: Lists are created for staff group names, student group names, and a group for all users. Descriptions for these groups are also defined.
3: Create or Check Existence of Staff Groups: Loops through each staff group name, checks if the group already exists, and creates it if it doesn't.
4: Create or Check Existence of Student Groups: Similar to the staff groups, this section creates or checks the existence of student groups.
5: Create or Check Existence of All Users Group: Checks if the group for all users exists and creates it if it doesn't.
6: Wait for Groups to Sync: Waits for a specified period before continuing execution. This is likely to ensure that any newly created groups are fully synchronised and available for membership changes.
7: Add Staff Group and Student Group as Members to All Users Group: Adds the staff group and student group as members to the group representing all users.

#>

<#
--------------------------------- IMPORTANT--------------------------

If you would like to intergrate a Student setup, please let me know
                                                        
--------------------------------- IMPORTANT--------------------------
#>



# Check if ExchangeOnlineManagement module is installed
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "ExchangeOnlineManagement module is not installed. Please install the module before proceeding." -ForegroundColor Red
    exit
}

# Prompt to ensure varibles have been setup correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the varibles?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# 1: Define Variables
$matCodeAlias = "SCPS"
$siteCodes = @("FHI", "FRJ", "JRI")
$Type = "MESG"
$Domain = "@southcamberley.surrey.sch.uk"
$AdminAccount = "eduadmin"

# 2: Define varibles for the groups hash table for matCodes staff with prefix
$matGroups = @{
    "MAT_All_Staff" = "${Type}_${matCodeAlias}_All_Staff"
    "MAT_Office_All_Staff" = "${Type}_${matCodeAlias}_Office_All_Staff"
    "MAT_Teachers_All_Staff" = "${Type}_${matCodeAlias}_Teachers_All_Staff"
    "MAT_Personnel_All_Staff" = "${Type}_${matCodeAlias}_Personnel_All_Staff"
    "MAT_SLT_All_Staff" = "${Type}_${matCodeAlias}_SLT_All_Staff"
    "MAT_Safeguarding_All_Staff" = "${Type}_${matCodeAlias}_Safeguarding_All_Staff"
    "MAT_Finance_All_Staff" = "${Type}_${matCodeAlias}_Finance_All_Staff"
}

# 3: Define varibles for the groups hash table for site staff with prefix
$siteGroups = @{}

foreach ($siteCode in $siteCodes) {
    $siteGroups[$siteCode] = @{
        "All_Staff" = "${Type}_${siteCode}_All_Staff"
        "Office_All_Staff" = "${Type}_${siteCode}_Office_All_Staff"
        "Teachers_All_Staff" = "${Type}_${siteCode}_Teachers_All_Staff"
        "Personnel_All_Staff" = "${Type}_${siteCode}_Personnel_All_Staff"
        "SLT_All_Staff" = "${Type}_${siteCode}_SLT_All_Staff"
        "Safeguarding_All_Staff" = "${Type}_${siteCode}_Safeguarding_All_Staff"
        "Finance_All_Staff" = "${Type}_${siteCode}_Finance_All_Staff"
    }
}

# Section 1: Loop through matGroups to create groups for matCodes staff
foreach ($matGroup in $matGroups.Keys) {  # Use $matGroups.Keys to access the keys of the hash table
    $matGroupName = $matGroups[$matGroup]  # Use $matGroups[$matGroup] to access the value (group name) corresponding to the key (group description)
    $matGroupDescription = "This group is used for Intune permissions for $matGroupName" # Define the description here
    if (-not (Get-DistributionGroup -Identity $matGroupName -ErrorAction SilentlyContinue)) {
        Write-Host "Creating mail-enabled security group: $matGroupName" -ForegroundColor Green
        New-DistributionGroup -Name $matGroupName -Alias $matGroupName -PrimarySmtpAddress "${matGroupName}$Domain" -Description $matGroupDescription -Type "Security" | Out-Null
    }
    else {
        Write-Host "Group already exists: $matGroupName" -ForegroundColor Red
    }
}

# Section 2: Loop through siteGroups to create groups for site staff
foreach ($siteCode in $siteGroups.Keys) {
    $siteGroupData = $siteGroups[$siteCode]  # Access the nested hash table for the current site code
    foreach ($staffCategory in $siteGroupData.Keys) {
        $siteGroupName = $siteGroupData[$staffCategory]  # Access the group name for the current staff category
        $siteGroupDescription = "This group is used for Intune permissions for $siteGroupName" # Define the description here
        if (-not (Get-DistributionGroup -Identity $siteGroupName -ErrorAction SilentlyContinue)) {
            Write-Host "Creating mail-enabled security group: $siteGroupName" -ForegroundColor Yellow
            New-DistributionGroup -Name $siteGroupName -Alias $siteGroupName -PrimarySmtpAddress "${siteGroupName}$Domain" -Description $siteGroupDescription -Type "Security" | Out-Null
        }
        else {
            Write-Host "Group already exists: $siteGroupName" -ForegroundColor Red
        }
    }
    Write-Host "-------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Waiting for 5 seconds before proceeding..." -ForegroundColor Cyan
    Write-Host "-------------------------------------------------------" -ForegroundColor Cyan
    Start-Sleep -Seconds 5  # Pause execution for 5 seconds before processing the next site code
}

# Section 3: Loop through siteGroups to add site staff groups to corresponding MAT groups
foreach ($siteCode in $siteGroups.Keys) {
    $siteGroupData = $siteGroups[$siteCode]  # Access the nested hash table for the current site code

    foreach ($staffCategory in $siteGroupData.Keys) {
        $siteGroup = $siteGroupData[$staffCategory]  # Get the group for the current staff category at the current site
        $matGroup = $matGroups["MAT_$staffCategory"]  # Get the corresponding MAT group for the current staff category

        # Add the site group to the corresponding MAT group
        if (-not (Get-DistributionGroupMember -Identity $matGroup -ErrorAction SilentlyContinue | Where-Object {$_.PrimarySmtpAddress -eq $siteGroup})) {
            Write-Host "Adding $siteGroup to $matGroup" -ForegroundColor Green
            Add-DistributionGroupMember -Identity $matGroup -Member $siteGroup | Out-Null
        }
        else {
            Write-Host "$siteGroup is already a member of $matGroup" -ForegroundColor Red
        }
    }   
}


<# USE THIS TO DELETE GROUPS FOR TESTING PURPOSES

# Check if ExchangeOnlineManagement module is installed
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "ExchangeOnlineManagement module is not installed. Please install the module before proceeding." -ForegroundColor Red
    exit
}

# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# Define Variables
$matCodeAlias = "SCPS"
$siteCodes = @("FHI", "FRJ", "JRI")
$Type = "MESG"
$Domain = "@southcamberley.surrey.sch.uk"
$AdminAccount = "eduadmin"

# Define varibles for the groups hash table for matCodes staff with prefix
$matGroups = @{
    "MAT_All_Staff" = "${Type}_${matCodeAlias}_All_Staff"
    "MAT_Office_All_Staff" = "${Type}_${matCodeAlias}_Office_All_Staff"
    "MAT_Teachers_All_Staff" = "${Type}_${matCodeAlias}_Teachers_All_Staff"
    "MAT_Personnel_All_Staff" = "${Type}_${matCodeAlias}_Personnel_All_Staff"
    "MAT_SLT_All_Staff" = "${Type}_${matCodeAlias}_SLT_All_Staff"
    "MAT_Safeguarding_All_Staff" = "${Type}_${matCodeAlias}_Safeguarding_All_Staff"
    "MAT_Finance_All_Staff" = "${Type}_${matCodeAlias}_Finance_All_Staff"
}

# Define varibles for the groups hash table for site staff with prefix
$siteGroups = @{}

foreach ($siteCode in $siteCodes) {
    $siteGroups[$siteCode] = @{
        "All_Staff" = "${Type}_${siteCode}_All_Staff"
        "Office_All_Staff" = "${Type}_${siteCode}_Office_All_Staff"
        "Teachers_All_Staff" = "${Type}_${siteCode}_Teachers_All_Staff"
        "Personnel_All_Staff" = "${Type}_${siteCode}_Personnel_All_Staff"
        "SLT_All_Staff" = "${Type}_${siteCode}_SLT_All_Staff"
        "Safeguarding_All_Staff" = "${Type}_${siteCode}_Safeguarding_All_Staff"
        "Finance_All_Staff" = "${Type}_${siteCode}_Finance_All_Staff"
    }
}

# Loop through matGroups to delete groups for matCodes staff
foreach ($matGroup in $matGroups.Keys) {
    $matGroupName = $matGroups[$matGroup]
    if (Get-DistributionGroup -Identity $matGroupName -ErrorAction SilentlyContinue) {
        Write-Host "Deleting mail-enabled security group: $matGroupName" -ForegroundColor Yellow
        Remove-DistributionGroup -Identity $matGroupName -Confirm:$false
    }
    else {
        Write-Host "Group does not exist: $matGroupName" -ForegroundColor Red
    }
}

# Loop through siteGroups to delete groups for site staff
foreach ($siteCode in $siteGroups.Keys) {
    $siteGroupData = $siteGroups[$siteCode]
    foreach ($staffCategory in $siteGroupData.Keys) {
        $siteGroupName = $siteGroupData[$staffCategory]
        if (Get-DistributionGroup -Identity $siteGroupName -ErrorAction SilentlyContinue) {
            Write-Host "Deleting mail-enabled security group: $siteGroupName" -ForegroundColor Yellow
            Remove-DistributionGroup -Identity $siteGroupName -Confirm:$false
        }
        else {
            Write-Host "Group does not exist: $siteGroupName" -ForegroundColor Red
        }
    }
}

#>