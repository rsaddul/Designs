# Install and import the ExchangeOnlineManagement module if not already installed
if (-not (Get-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online with progress display
Connect-ExchangeOnline -ShowProgress $true

# Define the output file path
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\KST\Shared_Mailboxes.csv"

# Create an array to store mailbox permissions
$permissionsArray = @()

# Get all shared mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox

foreach ($mailbox in $mailboxes) {
    $mailboxName = $mailbox.DisplayName
    
    # Get Read and manage permissions
    $readManagePermissions = Get-MailboxPermission -Identity $mailboxName | Where-Object { $_.User -like "*@*" -and $_.AccessRights -eq "FullAccess" }
    foreach ($permission in $readManagePermissions) {
        $permissionsArray += [PSCustomObject]@{
            "Mailbox Name" = $mailboxName
            "Shared Mailbox Address" = $mailbox.PrimarySmtpAddress
            "Permission Type" = "Read and manage"
            "User" = $permission.User
        }
    }

    # Get Send as permissions
    $sendAsPermissions = Get-RecipientPermission -Identity $mailboxName | Where-Object { $_.Trustee -like "*@*" -and $_.AccessRights -eq "SendAs" }
    foreach ($permission in $sendAsPermissions) {
        $permissionsArray += [PSCustomObject]@{
            "Mailbox Name" = $mailboxName
            "Shared Mailbox Address" = $mailbox.PrimarySmtpAddress
            "Permission Type" = "Send as"
            "User" = $permission.Trustee
        }
    }

    # Get Send on behalf of permissions
    $sendOnBehalfOfPermissions = Get-MailboxPermission -Identity $mailboxName | Where-Object { $_.User -like "*@*" -and $_.AccessRights -eq "SendAs" }
    foreach ($permission in $sendOnBehalfOfPermissions) {
        $permissionsArray += [PSCustomObject]@{
            "Mailbox Name" = $mailboxName
            "Shared Mailbox Address" = $mailbox.PrimarySmtpAddress
            "Permission Type" = "Send on behalf of"
            "User" = $permission.User
        }
    }
}

# Export permissions to CSV
$permissionsArray | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Mailbox permissions exported to $outputFilePath"