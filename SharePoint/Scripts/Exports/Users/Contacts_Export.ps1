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
$ExportPath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\Users\Contacts_Export.csv"

# Connect to Exchange Online with progress display
Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false

# Retrieve guest contacts with specified properties and export them to CSV file
Get-Contact -RecipientTypeDetails MailContact | 
    Select-Object DisplayName, FirstName, LastName, WindowsEmailAddress, RecipientType, DistinguishedName, IsDirSynced | 
    Export-Csv -Path $ExportPath -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
