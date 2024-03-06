<#
Developed by: Rhys Saddul
#>

# Install and import the Microsoft.Graph.Users module if not already installed
# Check if the module is installed
if (-not (Get-Module -Name Microsoft.Graph.Users -ListAvailable)) {
    # Install the module
    Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Force
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
$ExportPath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\Users\Guests_Export.csv"

# Connect to Microsoft Graph
Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome | Out-Null

# Retrieve Office 365 guests and their IDs
$GuestUserIds = Get-MgUser -All -Filter "UserType eq 'Guest'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id

# Initialize an empty array to store guest user details
$GuestUsers = @()

# Retrieve detailed information for each guest user
foreach ($userId in $GuestUserIds) {
    $userDetail = Get-MgUser -UserId $userId
    $GuestUsers += [PSCustomObject]@{
        GivenName = $userDetail.GivenName
        Surname = $userDetail.Surname
        DisplayName = $userDetail.DisplayName
        Mail = $userDetail.Mail
        UserPrincipalName = $userDetail.UserPrincipalName
        Id = $userDetail.Id
    }
}

# Export guest user details to CSV file with UTF-8 encoding without BOM
$GuestUsers | Export-Csv -Path $ExportPath -Encoding UTF8 -NoTypeInformation

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

# Check if any of the guest users have empty or null
# $GuestUsers | Where-Object { $_.GivenName -eq $null -or $_.GivenName -eq '' -or $_.Surname -eq $null -or $_.Surname -eq '' }
