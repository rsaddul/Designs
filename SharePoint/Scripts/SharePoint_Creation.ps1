<#
Developed by: Rhys Saddul

Overview:

1:  Variable Setup: Defines various variables such as SharePoint URLs, site codes, suffixes, site titles, etc.
2:  Child Sites Creation: Checks if child sites exist and creates them if not.
3:  Hub Site Creation: Checks if the hub site exists and creates it if not.
4:  Registering Hub Site: Registers the hub site if it's not already registered.
5:  Associating Child Sites with Hub Site: Associates child sites with the hub site.
6:  Custom Configuration for CloudMapper: Configures custom permissions and settings for CloudMapper, a tool used for visualising and assessing network security.
7:  Creating Libraries on Sites: Creates libraries on sites as defined in the script.
8:  Removing Default Visitors SharePoint Group: Removes the default Visitors SharePoint group from document libraries.
9:  Checking Mail-Enabled Security Groups: Checks if mail-enabled security groups exist.
10: Configuring Permissions for Default Visitors SharePoint Group: Configures permissions for the default Visitors SharePoint group.
11: Custom Permissions Configuration for Libraries: Configures custom permissions for libraries on sites.
12: Configuring Permissions for Home ASPX Page: Configures permissions for the Home ASPX page on each site.
13: Applying Permission Changes to Default Read Permission Level: Applies permission changes to the default Read permission level for each site.

#>

<#
-------------------------------- IMPORTANT-------------------------------- 

If you do not use CloudMapper, comment out the below Varibles and Sections.
                            Varibles: 4, 9, 10
                            Sections: 8, 14, 15

-------------------------------- IMPORTANT-------------------------------- 
#>

<#
--------------------------------- IMPORTANT----------------------------------

If you would like to intergrate a Student setup, un-hash this in the Varibles
                            Varibles: 1, 2, 3, 4, 5, 7, 9, 10
                            
--------------------------------- IMPORTANT---------------------------------- 
#>

<#
 1: Base variables: Used in all of other varibles
#>
$adminCenterURL = "https://hawkedaleinfantschool-admin.sharepoint.com" # SharePoint Admin URL (Used in Section 6)
$siteTemplate = "SITEPAGEPUBLISHING#0" # Communication site template (Used inSection 6)
$baseSharePointURL = "https://hawkedaleinfantschool.sharepoint.com/sites" # Used throughout the script to connect to different sites
$SiteOwner = "eduadmin@HawkedaleInfantSchool.onmicrosoft.com"  # User that will be made the owner for site creation (Used in Section 5 and Section 6)
$hubSiteTitle = "RSA School" # Should be the same as $sitecode variable (Used in Section 6)
$hubSiteDescription = "This is the hub site." # (Used in Section 6)
$siteCode = "RSA" # Used throughout the script for SharePoint site targeting 
$Suffix = "@hawkedale.surrey.sch.uk"  # Suffix for mail-enabled security groups (Used in Varible 3)
$staffSite = "Staff" # Used throughout the script for SharePoint site targeting
$operationSite = "Operations" # Used throughout the script for SharePoint site targeting
# $studentSite = "Student" # Used throughout the script for SharePoint site targeting

# 2: Define sites that will be created and associated with the hub site. (Used in Section 1 and 6) 
$childSites = @(  
    @{
        "Title" = "$siteCode $staffSite"
        "RelativeURL" = "$siteCode-$staffSite"
    },
    @{
        "Title" = "$siteCode $operationSite"
        "RelativeURL" = "$siteCode-$operationSite"
    }<#,
    @{
        "Title" = "$siteCode $studentSite"
        "RelativeURL" = "$siteCode-$studentSite"
    }#>
    # --------- DO NOT ADD THE HUB SITE ------ Add more sites as needed
)

# 3: Define mail-enabled security groups which will be used for site and library permissions. (Used Section 11, Varible 7 and 8)
# Site Section
$allUsers = "MESG_${siteCode}_All_Users" + $Suffix 

# Staff Groups
$allStaff = "MESG_${siteCode}_All_Staff" + $Suffix 
$safeGuarding = "MESG_${siteCode}_Safeguarding_All_Staff" + $Suffix 
$office = "MESG_${siteCode}_Office_All_Staff" + $Suffix 
#$finance = "MESG_${siteCode}_Finance_All_Staff" + $Suffix 
$personnel = "MESG_${siteCode}_Personnel_All_Staff" + $Suffix 
$slt = "MESG_${siteCode}_SLT_All_Staff" + $Suffix
# --------- IMPORTANT --------- Add more groups as needed (you will need to update Section 11)

# Student Groups
$allStudents = "MESG_${siteCode}_All_Students" + $Suffix 
#$year1 = "MESG_${siteCode}_Year1_All_Students" + $Suffix 
#$year2 = "MESG_${siteCode}_Year2_All_Students" + $Suffix
#$year3 = "MESG_${siteCode}_Year3_All_Students" + $Suffix
#$year4 = "MESG_${siteCode}_Year4_All_Students" + $Suffix
#$year5 = "MESG_${siteCode}_Year5_All_Students" + $Suffix 
# --------- IMPORTANT --------- Add more groups as needed (you will need to update Section 11)

# Create an array containing all the group variables
$groups = @($allUsers, $allStaff, $safeGuarding, $office, $finance, $personnel, $slt, $allStudents, $year1, $year2, $year3, $year4, $year5)

# Filter out undefined or empty variables from the array
$definedGroups = $groups | Where-Object { $_ -ne $null -and $_ -ne '' }

# 4: Bespoke CloudMapper Config: Define what sites will have the custom Home ASPX permission created on. (Used in Section 8)
$HomeASPX = @(    
    @{
        Url             = "$baseSharePointURL/$($siteCode)-$($staffSite)"
        CustomPermission= "Home ASPX"
        BasePermission  = "Read"
        Description     = "This is used for custom Home ASPX permissions"
        Permissions     = 'BrowseDirectories'  # Exclude 'AddListItems', 'EditListItems'
    },
    @{
        Url             = "$baseSharePointURL/$($siteCode)-$($operationSite)"
        CustomPermission= "Home ASPX"
        BasePermission  = "Read"
        Description     = "This is used for custom Home ASPX permissions"
        Permissions     = 'BrowseDirectories'  # Exclude 'AddListItems', 'EditListItems'
    }<#,
    @{
        Url             = "$baseSharePointURL/$($siteCode)-$($studentSite)"
        CustomPermission= "Home ASPX"
        BasePermission  = "Read"
        Description     = "This is used for custom Home ASPX permissions"
        Permissions     = 'BrowseDirectories'  # Exclude 'AddListItems', 'EditListItems'
    }#>
    # --------- DO NOT ADD THE HUB SITE ------ Add more sites as needed
)

# 5: Define sites and the libaries you would like created. # (Used in Section 13 and Section 14)
$sitesWithLibraries = @(  
    @{
        SiteUrl   = "$baseSharePointURL/$($siteCode)-$($staffSite)"
        Libraries = @("General", "Media", "Planning", "Safeguarding")
    },
    @{
        SiteUrl   = "$baseSharePointURL/$($siteCode)-$($operationSite)"
        Libraries = @("Office", "Finance", "Personnel", "SLT")
    }<#,
    @{
        SiteUrl   = "$baseSharePointURL/$($siteCode)-$($studentSite)"
        Libraries = @("General", "Year1", "Year2", "Year3", "Year4", "Year5")
    }#>
    # --------- DO NOT ADD THE HUB SITE ------ Add more sites as needed
)

# 6: Define variable for the site/s you dont want to create libraies on (Used in Section 9)
$siteWithoutLibraries = @{
    SiteUrl   = "$baseSharePointURL/${siteCode}" 
    Groups    = @($allUsers)
    # --------- TARGET THE HUB SITE ONLY ------# Add more sites as needed
}

# 7: Define sites and security groups that will be used for setting the site SharePoint Group Visitors permissions. (Used in Section 12)
$groupsForPermissionsAndAssociation = @(  
    @{
        SiteUrl   = "$baseSharePointURL/$($siteCode)-$($staffSite)"
        Groups    = @($allStaff) # (Referencing Varible 3)
    },
    @{
        SiteUrl   = "$baseSharePointURL/$($siteCode)-$($operationSite)"
        Groups    = @($allStaff) # (Referencing Varible 3)
    },
    @{
        SiteUrl   = "$baseSharePointURL/${siteCode}"
        Groups    = @($allStaff,$allStudents) # (Referencing Varible 3)
    }<#,
    @{
        SiteUrl   = "$baseSharePointURL/$($siteCode)-$($studentSite)"
        Groups    = @($allStaff,$allStudents) # (Referencing Varible 3)
    }#>
    # Add more groups as needed
)

# 8: Define site and specify which security groups will be applied to each library with edit or read permissions. (Used in Section 13 and References Variable 3)
$sitesLibrariesSecurityGroupsMapping = @{
    "$baseSharePointURL/$($siteCode)-$($staffSite)" = @{
        "General" = @{
            "Edit" = @("$allStaff")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        "Media" =  @{
            "Edit" = @("$allStaff")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        "Planning" =  @{
            "Edit" = @("$allStaff")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissionss
        }
        "Safeguarding" =  @{
            "Edit" = @("$safeguarding")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        # Add more libraries, and security groups as needed
    }
    "$baseSharePointURL/$($siteCode)-$($operationSite)" = @{
        "Office" = @{
            "Edit" = @("$office")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        "Finance" =  @{
            "Edit" = @("$finance")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        "SLT" =  @{
            "Edit" = @("$slt")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        "Personnel" =  @{
            "Edit" = @("$personnel")  # Group with edit permissions
           #"Read" = @("")       # No "Read" key specified for this library, indicating no read permissions
        }
        # Add more libraries, and security groups as needed
    } 
   <# "$baseSharePointURL/$($siteCode)-$($studentSite)" = @{
        "Year1" = @{
            "Edit" = @("$allStaff")  # Group with edit permissions
            "Read" = @("year1")  # Group with read permissions
        }
        "Year2" =  @{
            "Edit" = @("$allStaff")  # Group with edit permissions
            "Read" = @("Year2")       # Group with read permissions
        }
        "Year3" =  @{
            "Edit" = @("$allStaff")  # Group with edit permissions
            "Read" = @("Year3")       # Group with read permissions
        }
        "Year4" =  @{
            "Edit" = @("$allStaff")  # Group with edit permissions
            "Read" = @("$Year4")       # Group with read permissions
        }
        "Year5" =  @{
            "Edit" = @("$allStaff")  # Group with edit permissions
            "Read" = @("$Year5")       # Group with read permissions
        }
        # Add more libraries, and security groups as needed
    } #>
    # Add mappings for other sites as needed
}

# 9: Bespoke CloudMapper Config: Define the sites and specify the security groups that will be used for permissions on the Home ASPX page. (Used in Section 14)
$sitePagesAndRoles = @(
    @{
        SitePage  = "$baseSharePointURL/$($siteCode)-$($staffSite)/SitePages/Home.aspx"
        Groups    = @($allStaff)
    },
    @{
        SitePage  = "$baseSharePointURL/$($siteCode)-$($operationSite)/SitePages/Home.aspx"
        Groups    = @($allStaff)
    }<#,
    @{
        SitePage  = "$baseSharePointURL/$($siteCode)-$($studentSite)/SitePages/Home.aspx"
        Groups    = @("$allStaff", "$allStudents")
    }#>
    # --------- DO NOT ADD THE HUB SITE ------ Add more sites and groups as needed
)

# 10: Bespoke CloudMapper Config: Define the sites and specify the permissions you want to add or remove for the default Read permission level. (Used in Section 15)
$SitesPermissions = @(
    @{
        SiteURL = "$baseSharePointURL/$($siteCode)-$($staffSite)"
        ClearPermissions = @("ViewListItems", "OpenItems", "ViewVersions", "CreateAlerts", "ViewFormPages")
        AddPermissions = @("BrowseDirectories")
    },
    @{
        SiteURL = "$baseSharePointURL/$($siteCode)-$($operationSite)"
        ClearPermissions = @("ViewListItems", "OpenItems", "ViewVersions", "CreateAlerts", "ViewFormPages")
        AddPermissions = @("BrowseDirectories")
    }<#,
    @{
        SiteURL = "$baseSharePointURL/$($siteCode)-$($studentSite)"
        ClearPermissions = @("ViewListItems", "OpenItems", "ViewVersions", "CreateAlerts", "ViewFormPages")
        AddPermissions = @("BrowseDirectories")
    }#>
    # --------- DO NOT ADD THE HUB SITE ------ Add more sites as needed
)

# Prompt to ensure security groups have been setup prior to running the script
Add-Type -AssemblyName System.Windows.Forms

$dialogResult = [System.Windows.Forms.MessageBox]::Show("Did you run the SharePoint Group script to create security groups? If you press No we shall create them", "SharePoint Groups", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed security groups have been created. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed security groups have not been created. Please run User_Groups.ps1" -ForegroundColor Red
    return
}

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


# Section 1: Check if SharePoint sites exist. If not, create the sites listed in. (References Variable 2).
Write-Host "Checking if the child sites exist before creation" -ForegroundColor Yellow
Write-Host
foreach ($childSite in $childSites) {
    $siteUrl = "$baseSharePointURL/$($childSite["RelativeURL"])"

    Connect-PnPOnline -Url $siteUrl -Interactive

    $existingSite = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue
    if (-not $existingSite) {
        Connect-PnPOnline -Url $siteUrl -Interactive
        New-PnPTenantSite -Title $childSite["Title"] -Url $siteUrl -Owner $SiteOwner -Template $siteTemplate -TimeZone 4 -RemoveDeletedSite
        Write-Host "Creating $($childSite.Title)" -ForegroundColor Yellow
        Write-Host "$($childSite.Title) Created:" -ForegroundColor Green
        Write-Host
        Disconnect-PnPOnline
    } else {
        Write-Host "Site $($childSite.Title) already exists." -ForegroundColor Red
    }
}

# Section 2: Check if hub site exists. If not, create "$baseSharePointURL/$siteCode" listed in. (References Variables 1)
Write-Host "Checking if the hub site exists before creation" -ForegroundColor Yellow
Write-Host
$hubSiteUrl = "$baseSharePointURL/$siteCode"

    Connect-PnPOnline -Url $hubSiteUrl -Interactive

$existingHubSite = Get-PnPTenantSite -Url $hubSiteUrl -ErrorAction SilentlyContinue
if (-not $existingHubSite) {
    $hubSite = New-PnPTenantSite -Title $hubSiteTitle -Url $hubSiteUrl -Owner $SiteOwner -Template $siteTemplate -TimeZone 4 -RemoveDeletedSite
    Write-Host "Creating Hub site $hubSiteTitle" -ForegroundColor Yellow
    Write-Host "Hub site created: $hubSiteTitle" -ForegroundColor Green
    Write-Host
    Disconnect-PnPOnline

    # Section 3: Wait for sites to be created
    Write-Host "Be patient while the sites are syncing" -ForegroundColor Yellow
    for ($i = 90; $i -ge 0; $i--) {
        Write-Host "$i" -ForegroundColor Cyan
        Start-Sleep -Seconds 1
    }
    Write-Host "Site creation Countdown complete!" -ForegroundColor Green
    Write-Host
    Write-Host "Thank you for being patient" -ForegroundColor Cyan
    Write-Host
} else {
    Write-Host "Hub site $($existingHubSite.Title) already exists." -ForegroundColor Red
}

# Section 4: Check if site is registered as a hub site. If not, register $hubSiteUrl listed in. (Used in Section 2)
    Connect-PnPOnline -Url $hubSiteUrl -Interactive

# Check if the site is already registered as a HubSite
$existingHub = Get-PnPHubSite | Where-Object { $_.SiteUrl -eq $hubSiteUrl }

if ($existingHub) {
    Write-Host "Hub site $hubSiteUrl is already registered as a HubSite." -ForegroundColor Yellow
} else {
    Write-Host "Registering hub site $hubSiteUrl" -NoNewline
    try {
        Register-PnPHubSite -Site $hubSiteUrl
        for ($i = 15; $i -ge 0; $i--) { # Section 5: Wait for $hubSiteUrl registration (Used in Section 4)
            Write-Host "$i" -ForegroundColor Cyan
            Start-Sleep -Seconds 1
        }
        Write-Host "Hubsite registration countdown complete!" -ForegroundColor Green
        Write-Host
    } catch {
        Write-Warning "Failed to register hub site $hubSiteUrl as a HubSite. It might already be a HubSite or there was an error."
    }
}
    Disconnect-PnPOnline

# Section 6: Check if sites have an association. If not, associate sites with the $hubSiteUrl. (References Variable 2)
Write-Host "Associating child sites with the hub site" -ForegroundColor Yellow

# Define the hub site URL
$hubSiteUrl = "$baseSharePointURL/$siteCode"

foreach ($childSite in $childSites) {
    $siteUrl = "$baseSharePointURL/$($childSite["RelativeURL"])"
    Connect-PnPOnline -Url $siteUrl -Interactive
    
    # Get site properties to check HubSiteId
    $siteProperties = Get-PnPTenantSite -Identity $siteUrl | Select -ExpandProperty HubSiteId

    # Check if the site is already associated with a hub site
    if ($siteProperties -ne "00000000-0000-0000-0000-000000000000") {
        Write-Host "$($childSite.Title) is already associated with a hub site."
    } else {
        # Associate the site with the hub site
        Add-PnPHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl
        Write-Host "Associating $($childSite.Title) with $hubSiteUrl" -ForegroundColor Green

        # Wait for hub association
        for ($i = 15; $i -ge 0; $i--) {
            Write-Host "$i" -ForegroundColor Cyan
            Start-Sleep -Seconds 1
        }
        Write-Host "$($childSite.Title) association countdown complete!"
    }
    Disconnect-PnPOnline
}

# Section 8: Bespoke CloudMapper Config: Create Home ASPX permission level on sites listed in. (References Varible 4)
Write-Host
Write-Host "Custom configuration for CloudMapper" -ForegroundColor Magenta

# Loop through each site in the HomeASPX array (References Varible 4)
foreach ($site in $HomeASPX) {
    Connect-PnPOnline -Url $site.Url -Interactive
    # Get the base permission level
    $BasePermissionLevel = Get-PnPRoleDefinition -Identity $site.BasePermission

    # Check if the base permission level exists
    if ($BasePermissionLevel -ne $null) {
        # Set parameters for the new permission level
        $NewPermissionLevel = @{
            Include     = $site.Permissions
            Description = $site.Description
            RoleName    = $site.CustomPermission
            Clone       = $BasePermissionLevel
        }

        # Create custom permission level for the site
        Add-PnPRoleDefinition @NewPermissionLevel | Out-Null

        Write-Host "Custom CloudMapper permission level setup for site $($site.Url)" -ForegroundColor Green
    } else {
        Write-Host "Base permission level '$($site.BasePermission)' not found for $($site.Url)."
    }
    Disconnect-PnPOnline
}

# Section 9: Check if libraries exist. If not, create libraries on sites listed in. (References Variable 5)
Write-Host
Write-Host "Creating libraries on sites" -ForegroundColor Yellow
foreach ($site in $sitesWithLibraries) { 
    if ($site.SiteUrl -ne $siteWithoutLibraries.SiteUrl) { # Dont create libraries listed on site in (References Variable 6)
        Connect-PnPOnline -Url $site.SiteUrl -Interactive
        
        foreach ($library in $site.Libraries) {
            # Check if the library already exists
            $existingLibrary = Get-PnPList -Identity $library -ErrorAction SilentlyContinue
            if (-not $existingLibrary) {
                New-PnPList -Title $library -Template DocumentLibrary -ErrorAction SilentlyContinue -Connection $ctx | Out-Null
                Write-Host "Library '$library' created on $($site.SiteUrl)" -ForegroundColor Green
            } else {
                Write-Host "Library '$library' already exists on $($site.SiteUrl)" -ForegroundColor Yellow
            }
        }
        Disconnect-PnPOnline
    }
}

# Section 10: Check if the default Hub Visitors or default Visitors group exists on all default document libraries across all sites. If they exist, remove. (References Variable 5 and 6).
# Add the new site without libraries to the existing array (References Varible 5 and 6)
$sitesWithLibraries += $siteWithoutLibraries
Write-Host
Write-Host "Removing the default Visitors SharePoint Group from default document libraries" -ForegroundColor Yellow
foreach ($site in $sitesWithLibraries) {
    Connect-PnPOnline -Url $site.SiteUrl -Interactive
    $defaultDocumentLibrary = Get-PnPList -Identity "Documents"
    if ($defaultDocumentLibrary) {
        Set-PnPList -Identity "Documents" -BreakRoleInheritance -CopyRoleAssignments | Out-Null
        $allGroups = Get-PnPGroup
        # Remove the default Visitor group using its ID (4)
        $visitorsGroup = $allGroups | Where-Object { $_.Id -eq 4 }
        if ($visitorsGroup) {
            try {
                $defaultDocumentLibrary.RoleAssignments.GetByPrincipal($visitorsGroup).DeleteObject()
                Invoke-PnPQuery
                Write-Host "The default Visitors SharePoint group has been removed from the default document library in $($site.SiteUrl)" -ForegroundColor Green
            } catch {
                Write-Host "An error occurred while removing the default Visitors SharePoint group from the default document library in $($site.SiteUrl)." -ForegroundColor Red
            }
        } else {
            Write-Host
            Write-Host "The default Visitors SharePoint group (ID: 4) not found in $($site.SiteUrl). No action needed." -ForegroundColor Cyan
        }

        # Remove the default Hub Visitors group
        $hubVisitorsGroup = $allGroups | Where-Object { $_.Title -eq "Hub Visitors" }
        if ($hubVisitorsGroup) {
            try {
                $defaultDocumentLibrary.RoleAssignments.GetByPrincipal($hubVisitorsGroup).DeleteObject()
                Invoke-PnPQuery
                Write-Host "The Hub Visitors group has been removed from the default document library in $($site.SiteUrl)" -ForegroundColor Green
            } catch {
                Write-Host "An error occurred while removing the Hub Visitors group from the default document library in $($site.SiteUrl)." -ForegroundColor Red
            }
        } else {
            Write-Host "The Hub Visitors group not found in $($site.SiteUrl). No action needed." -ForegroundColor Cyan
        }
    } else {
        Write-Host "Default document library not found on $($site.SiteUrl). Skipping." -ForegroundColor Cyan
    }
    # Restore the array to original (References Variable 5 and 6)
    $sitesWithLibraries = $sitesWithLibraries | Where-Object { $_ -ne $siteWithoutLibraries } 
    Disconnect-PnPOnline
}

# Section 11: Check if mail-enabled security groups exist. If any do not, then stop the script. (References Variable 3)
Write-Host
Write-Host "Checking if mail-enabled security groups exist" -ForegroundColor Yellow
Connect-ExchangeOnline -ShowBanner:$false
foreach ($group in $definedGroups) { 
    if (-not (Get-DistributionGroup -Identity $group -ErrorAction SilentlyContinue)) {
        Write-Warning "'$group' does not exist."
        Disconnect-ExchangeOnline -Confirm:$false
        return
    }
    else {
        Write-Host "'$group' exists." -ForegroundColor Green
    }
}
Disconnect-ExchangeOnline -Confirm:$false

# Section 12: Check if Hub Site group exist and configure permissions for default SharePoint Group Visitors on sites listed in. (References Varible 7)

# Define the Visitors group ID
$visitorGroupId = 4
$hubVisitorsGroupId1 = 12
$hubVisitorsGroupId2 = 13
$hubVisitorsGroupId4 = 14
$hubVisitorsGroupId5 = 15

# Loop through each site in the $groupsForPermissionsAndAssociation variable
foreach ($site in $groupsForPermissionsAndAssociation) {
    # Extract site URL from the current site object
    $siteUrl = $site["SiteUrl"]
    Write-Host
    Write-Host "Configuring security groups on: $siteUrl" -ForegroundColor Yellow

    Connect-PnPOnline -Url $siteUrl -Interactive 
       
    # Check if "Hub Visitors" group with ID 12 exists on the site and remove if found
    $hubVisitorsGroup1 = Get-PnPGroup | Where-Object { $_.Id -eq $hubVisitorsGroupId1 }
    if ($hubVisitorsGroup1) {
        Remove-PnPGroup -Identity $hubVisitorsGroupId1 -Force -ErrorAction SilentlyContinue
        Write-Host "Removed group with ID '$hubVisitorsGroupId1' from $siteUrl" -ForegroundColor Yellow
    }

    # Check if "Hub Visitors" group with ID 13 exists on the site and remove if found
    $hubVisitorsGroup2 = Get-PnPGroup | Where-Object { $_.Id -eq $hubVisitorsGroupId2 }
    if ($hubVisitorsGroup2) {
        Remove-PnPGroup -Identity $hubVisitorsGroupId2 -Force -ErrorAction SilentlyContinue
        Write-Host "Removed group with ID '$hubVisitorsGroupId2' from $siteUrl" -ForegroundColor Yellow
    }

        # Check if "Hub Visitors" group with ID 14 exists on the site and remove if found
    $hubVisitorsGroup3 = Get-PnPGroup | Where-Object { $_.Id -eq $hubVisitorsGroupId3 }
    if ($hubVisitorsGroup3) {
        Remove-PnPGroup -Identity $hubVisitorsGroupId3 -Force -ErrorAction SilentlyContinue
        Write-Host "Removed group with ID '$hubVisitorsGroupId3' from $siteUrl" -ForegroundColor Yellow
    }

        # Check if "Hub Visitors" group with ID 15 exists on the site and remove if found
    $hubVisitorsGroup4 = Get-PnPGroup | Where-Object { $_.Id -eq $hubVisitorsGroupId4 }
    if ($hubVisitorsGroup4) {
        Remove-PnPGroup -Identity $hubVisitorsGroupId4 -Force -ErrorAction SilentlyContinue
        Write-Host "Removed group with ID '$hubVisitorsGroupId4' from $siteUrl" -ForegroundColor Yellow
    }

    # Check if the default "SharePoint Group Visitors" group exists using the defined ID
    $visitorsGroup = Get-PnPGroup | Where-Object { $_.Id -eq $visitorGroupId }

    if ($visitorsGroup) {
        # Loop through each group email provided in the current site object
        foreach ($groupEmail in $site["Groups"]) {
            # Add the group email as a member to the "Visitors" group
            Add-PnPGroupMember -Identity $visitorGroupId -EmailAddress $groupEmail
            Write-Host "Added $groupEmail to Visitors group on $siteUrl" -ForegroundColor Green
        }
    } else {
        Write-Warning "Default 'Visitors' group not found on $siteUrl"
    }

    Disconnect-PnPOnline
    Write-Host "Finished setting up security groups on: $siteUrl" -ForegroundColor Yellow
    Write-Host
}

# Section 13: Disables inheritance, removes default SharePoint Visitor Group and Hub Visitors on all libraries and sites list in. (References Variable 8)
# Sets library permissions on all sites listed in. (References Variable 8)
foreach ($siteUrl in $sitesLibrariesSecurityGroupsMapping.Keys) {
    Write-Host "Connected to $siteUrl" -ForegroundColor Yellow

    Connect-PnPOnline -Url $siteUrl -Interactive

    # Iterate over each library in the site
    foreach ($library in $sitesLibrariesSecurityGroupsMapping[$siteUrl].Keys) {
        Write-Host
        Write-Host "Disabling inheritance, removing default visitor group, and configuring permissions for: $library" -ForegroundColor Green

        $permissions = $sitesLibrariesSecurityGroupsMapping[$siteUrl][$library]

        # Break role inheritance on the library
        $libraryObject = Get-PnPList -Identity $library
        $libraryObject.BreakRoleInheritance($true, $false)

        # Remove default visitor group (ID 4) if it exists within the library
        try {
            $visitorsGroup = Get-PnPGroup -Identity 4 -ErrorAction Stop
            $libraryObject.RoleAssignments.GetByPrincipal($visitorsGroup).DeleteObject()
            Invoke-PnPQuery
            Write-Host "The default Visitors SharePoint group has been removed from the library '$library' on $($siteUrl)" -ForegroundColor Green
        } catch [Microsoft.SharePoint.Client.ServerException] {
            if ($_.Exception.Message -match "Can not find the principal with id: 4") {
                Write-Host "Default Visitors SharePoint group not found using group ID 4. No action needed." -ForegroundColor Cyan
            } else {
                Write-Host "An error occurred while removing the default Visitors SharePoint group from the library '$library' on $($siteUrl)." -ForegroundColor Red
                Write-Host "Error details: $_" -ForegroundColor Red
            }
        }

        # Remove Hub Visitors group by ID 12 if it exists within the library
        $hubVisitorsGroup = Get-PnPGroup -Identity 12 -ErrorAction SilentlyContinue
        if ($hubVisitorsGroup) {
            try {
                $libraryObject.RoleAssignments.GetByPrincipal($hubVisitorsGroup).DeleteObject()
                Invoke-PnPQuery
                Write-Host "The 'Hub Visitors' SharePoint group (ID 12) has been removed from the library '$library' on $($siteUrl)" -ForegroundColor Green
            } catch {
                Write-Host "An error occurred while removing the 'Hub Visitors' SharePoint group (ID 12) from the library '$library' in $($siteUrl)." -ForegroundColor Red
                Write-Host "Error details: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "'Hub Visitors' SharePoint group with ID 12 not found in $($siteUrl). No action needed." -ForegroundColor Cyan
        }

        # Configure permissions for the library
        foreach ($permissionLevel in $permissions.Keys) {
            $groupsToAdd = $permissions[$permissionLevel]
            foreach ($group in $groupsToAdd) {
                Set-PnPListPermission -Identity $library -User $group -AddRole $permissionLevel
                Write-Host "Added permissions for $group to $library with permission level $permissionLevel" -ForegroundColor Cyan
            }
        }

        Write-Host "Permissions configured for $library on $siteUrl" -ForegroundColor Green
    }

    Disconnect-PnPOnline
}

# Section 14: Bespoke CloudMapper Config:
# Disables inheritance, removes default SharePoint Visitor Group and Hub Visitors. 
# Sets security group permissions with a custom permission level for the Home ASPX page on each listed site. (References Variable 9)
Write-Host
Write-Host "Custom configuration for CloudMapper" -ForegroundColor Magenta
foreach ($sitePageConfig in $sitePagesAndRoles) {
    $siteUrl = $sitePageConfig.SitePage.Substring(0, $sitePageConfig.SitePage.IndexOf("/SitePages/Home.aspx"))
    $groupsToAddPermissions = $sitePageConfig.Groups

    Connect-PnPOnline -Url $siteUrl -Interactive

    # Get the Home.aspx page
    $homePage = Get-PnPListItem -List "Site Pages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>Home.aspx</Value></Eq></Where></Query></View>"

    if ($homePage) {
        # Break inheritance on the Home.aspx page
        Set-PnPListItemPermission -List "Site Pages" -Identity $homePage.Id -InheritPermissions:$false

        # Remove default SharePoint Visitor Group (ID 4) if it exists
        $defaultVisitorGroup = Get-PnPGroup -Identity 4 -ErrorAction SilentlyContinue
        if ($defaultVisitorGroup) {
        Set-PnPListItemPermission -List "Site Pages" -Identity $homePage.Id -Group $defaultVisitorGroup.Id -RemoveRole "Read" -ErrorAction SilentlyContinue
        } else {
        Write-Warning "Default visitor group with ID 4 not found or already removed. Skipping removal of permissions."
        }

        # Remove default Sharepoint Hub Visitors group using group ID 12 if it exists
        $hubVisitorsGroup = Get-PnPGroup -Identity 12 -ErrorAction SilentlyContinue
        if ($hubVisitorsGroup) {
        Set-PnPListItemPermission -List "Site Pages" -Identity $homePage.Id -Group $hubVisitorsGroup.Id -RemoveRole "Read" -ErrorAction SilentlyContinue
        } else {
        Write-Warning "Hub Visitors group with ID 12 not found or already removed. Skipping removal of permissions."
        }
                
        # Assign the custom role to the Home.aspx page for each group
        foreach ($group in $groupsToAddPermissions) {
            Set-PnPListItemPermission -List "Site Pages" -Identity $homePage.Id -User $group -AddRole "Home ASPX"  # Bespoke CloudMapper Role
        }
        Write-Host "Permissions configured for Home.aspx page on $($siteUrl)" -ForegroundColor Green    
    } else {
        Write-Host "Home.aspx page not found on $($siteUrl). Skipping." -ForegroundColor Yellow        
    }
    Disconnect-PnPOnline
}

# Section 15: Bespoke CloudMapper Config: Change the default Read permission level permissions for each site listed in. (References Variable 10)
Write-Host "Custom configuration for CloudMapper" -ForegroundColor Magenta
foreach ($SitePermissions in $SitesPermissions) {
    
    Connect-PnPOnline -Url $SitePermissions.SiteURL -Interactive

    # Set the role definition by clearing specified permissions
    Set-PnPRoleDefinition -Identity "Read" -Clear $SitePermissions.ClearPermissions | Out-Null
    Write-Host "Removed permissions for bespoke CloudMapper configuration on $($SitePermissions.SiteURL)" -ForegroundColor Green

    # Set the role definition by adding specified permissions
    Set-PnPRoleDefinition -Identity "Read" -Select $SitePermissions.AddPermissions | Out-Null
    Write-Host "Added bespoke CloudMapper permissions $($SitePermissions.SiteURL)" -ForegroundColor Green
    
    Disconnect-PnPOnline | Out-Null
}
Write-Host
Write-Host "-------------------------------- SharePoint design has now completed, proceed to visual design --------------------------------" -ForegroundColor Cyan