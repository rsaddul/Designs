<#
Developed by: Rhys Saddul

Overview:

1: Connect to Exchange Online: This connects the PowerShell session to Exchange Online without displaying the banner.
2: Define Variables:
                    $schoolCodeAlias: Represents the alias for the school.
                    $Type: Represents a type identifier.
                    $Domain: Represents the domain for the school's email addresses.
                    $AdminAccount: Specifies the admin account used for managing the groups.

3: Define Group Names and Descriptions using Arrays: Lists are created for staff group names, student group names, and a group for all users. Descriptions for these groups are also defined.
4: Create or Check Existence of Staff Groups: Loops through each staff group name, checks if the group already exists, and creates it if it doesn't.
5: Create or Check Existence of Student Groups: Similar to the staff groups, this section creates or checks the existence of student groups.
6: Create or Check Existence of All Users Group: Checks if the group for all users exists and creates it if it doesn't.
7: Wait for Groups to Sync: Waits for a specified period before continuing execution. This is likely to ensure that any newly created groups are fully synchronised and available for membership changes.
8: Add Staff Group and Student Group as Members to All Users Group: Adds the staff group and student group as members to the group representing all users.

#>

<#
--------------------------------- IMPORTANT----------------------------------

If you would like to intergrate a Student setup, un-hash the below
                            Varibles:  3
                            Sectoions: 2, 5
                            
--------------------------------- IMPORTANT---------------------------------- 
#>


# Check if ExchangeOnlineManagement module is installed
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "ExchangeOnlineManagement module is not installed. Please install the module before proceeding."
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

#  1: Define school code
$schoolCodeAlias = "HWK"
$Type = "MESG"
$Domain = "@hawkedale.surrey.sch.uk"
$AdminAccount = "eduadmin"  # Set the admin account here

#  2: Define staff group names and descriptions using arrays
$staffGroupNames = @("${Type}_${schoolCodeAlias}_All_Staff",
                    "${Type}_${schoolCodeAlias}_Office_All_Staff",
                    "${Type}_${schoolCodeAlias}_Personnel_All_Staff",
                    "${Type}_${schoolCodeAlias}_SLT_All_Staff",
                    "${Type}_${schoolCodeAlias}_Safeguarding_All_Staff"
                    "${Type}_${schoolCodeAlias}_Finance_All_Staff"
                    )

<# 3: Define student group names and descriptions using arrays

$studentGroupNames = @("${Type}_${schoolCodeAlias}_All_Students",
                      "${Type}_${schoolCodeAlias}_Year1_All-Students",
                      "${Type}_${schoolCodeAlias}_Year2_All-Students",
                      "${Type}_${schoolCodeAlias}_Year3_All-Students",
                      "${Type}_${schoolCodeAlias}_Year4_All-Students",
                      "${Type}_${schoolCodeAlias}_Year5_All-Students"
                      )

#>

# 4: Define all user groups
$allUsersGroupName = "${Type}_${schoolCodeAlias}_All_Users"

# 5: Define group name descriptions
$staffGroupDescription = "This group is used for Intune and SharePoint permissions for $schoolCodeAlias Staff"
$studentGroupDescription = "This group is used for Intune and SharePoint permissions for $schoolCodeAlias Students"
$allUsersGroupDescription = "This group is used for Intune and SharePoint permissions for all users at $schoolCodeAlias"

#  Section 1: Create or check the existence of staff groups
foreach ($groupName in $staffGroupNames) {
    $existingGroup = Get-DistributionGroup -Identity $groupName -ErrorAction SilentlyContinue

    if (-not $existingGroup) {
        Write-Host "Creating $groupName..." -ForegroundColor Green
        $newGroup = New-DistributionGroup -Name $groupName -Alias $groupName -Type Security -ManagedBy $AdminAccount -PrimarySmtpAddress "$groupName$Domain" -Description $staffGroupDescription
    } else {
        Write-Host "$groupName already exists." -ForegroundColor Yellow
    }
}

<# Section 2: Create or check the existence of student groups

foreach ($groupName in $studentGroupNames) {
    $existingGroup = Get-DistributionGroup -Identity $groupName -ErrorAction SilentlyContinue

    if (-not $existingGroup) {
        Write-Host "Creating $groupName..." -ForegroundColor Green
        $newGroup = New-DistributionGroup -Name $groupName -Alias $groupName -Type Security -ManagedBy $AdminAccount -PrimarySmtpAddress "$groupName$Domain" -Description $studentGroupDescription
    } else {
        Write-Host "$groupName already exists." -ForegroundColor Yellow
    }
}

#>

#  Section 3: Create or check the existence of all users group
$existingAllUsersGroup = Get-DistributionGroup -Identity $allUsersGroupName -ErrorAction SilentlyContinue

if (-not $existingAllUsersGroup) {
    Write-Host "Creating $allUsersGroupName..." -ForegroundColor Green
    $allUsersGroup = New-DistributionGroup -Name $allUsersGroupName -Alias $allUsersGroupName -Type Security -ManagedBy $AdminAccount -PrimarySmtpAddress "$allUsersGroupName$Domain" -Description $allUsersGroupDescription
} else {
    Write-Host "$allUsersGroupName already exists." -ForegroundColor Yellow
}

 for ($i = 15; $i -ge 0; $i--) { 
            Write-Host "Waiting for all groups to sync before adding nested members" -ForegroundColor Cyan
            Start-Sleep -Seconds 1
            }

#  Section 4: Add staff group as a member to all users group
$staffGroup = "${Type}_${schoolCodeAlias}_All_Staff"
Write-Host "Adding $staffGroup as a member to $allUsersGroupName..." -ForegroundColor Green
Add-DistributionGroupMember -Identity $allUsersGroupName -Member $staffGroup -ErrorAction SilentlyContinue

<# Section 5: Add student group as a member to all users group

$studentGroup = "${Type}_${schoolCodeAlias}_All_Students"
Write-Host "Adding $studentGroup as a member to $allUsersGroupName..." -ForegroundColor Green
Add-DistributionGroupMember -Identity $allUsersGroupName -Member $studentGroup -ErrorAction SilentlyContinue

#>

Disconnect-ExchangeOnline -Confirm:$false