# Specify the user list file path
$userListPath = "C:\Users.txt"
# Example
# user1@contoso.com
# user2@contoso.com
# user3@contoso.com


# Connect to Microsoft Online Services and SharePoint Online
Connect-MsolService
Connect-SPOService -URL https://your-school-admin.sharepoint.com

# If you need to unblock credentials for users who are blocked from signing in - run this command
Get-Content -Path $userListPath | ForEach-Object { Set-MsolUser -UserPrincipalName $_ -BlockCredential $False }

# Request personal sites for the specified users
$users = Get-Content -Path $userListPath
Request-SPOPersonalSite -UserEmails $users

######################################################################################################################

$TenantUrl = Read-Host "Enter the SharePoint admin center URL"

# Set the log file path on the desktop
$LogFile = [Environment]::GetFolderPath("Desktop") + "\OneDriveSites.log"

# Connect to SharePoint Online
Connect-SPOService -Url $TenantUrl

# Get a list of OneDrive URLs
Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object -ExpandProperty Url | Out-File $LogFile -Force

# Inform the user about the completion and the log file location
Write-Host "Done! File saved as $($LogFile)."
