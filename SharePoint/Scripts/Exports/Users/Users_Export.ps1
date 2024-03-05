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
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\BFLT\Users_Export.csv"

# Connect to Exchange Online with progress display
Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false

# Get all mailbox users
$mailboxes = Get-ExoMailbox

# Initialize an array to store mailbox information
$mailboxInfo = @()

# Iterate through each mailbox
# Iterate through each mailbox
foreach ($mailbox in $mailboxes) {
    # Check if the RecipientTypeDetails is UserMailbox
    if ($mailbox.RecipientTypeDetails -eq "UserMailbox") {
        # Create an object to store mailbox information
        $mailboxObject = [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            Alias = $mailbox.Alias
            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            RecipientType = $mailbox.RecipientTypeDetails
        }

        # Initialize an array to store email addresses
        $emailAddresses = @()

        # Add each email address to the array
        foreach ($emailAddress in $mailbox.EmailAddresses) {
            $emailAddresses += $emailAddress
        }

        # Add the email addresses as a property to the mailbox object
        $mailboxObject | Add-Member -MemberType NoteProperty -Name EmailAddresses -Value ($emailAddresses -join ";")

        # Add the mailbox object to the array
        $mailboxInfo += $mailboxObject
    }
}

# Export mailbox information to CSV
$mailboxInfo | Export-Csv -Path $outputFilePath -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false


Get-ExoMailbox vdeoliveria@kfos.co.uk | Select-Object EmailAddresses