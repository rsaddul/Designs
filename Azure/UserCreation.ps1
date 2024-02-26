# Run this on the server with ADDS
# Declare variables
$DateTime = Get-Date
$Date = $DateTime.ToShortDateString()
$Time = $DateTime.ToLongTimeString()
$DateTime.ToLongTimeString()
$Date = $Date.Replace("/", "-")
$Time = $Time.Replace(":", "-")
$CRLF = "`r`n"
$Separator = "==========================================================================================="
$n = 1

# Change these variables to suit
$CSVFile = "E:\UserCreation\SMB_UserCreation_Staff.csv"
$LogFile = "E:\UserCreation\SMB_UserCreation_Staff_Log.log"
$ReportPath = "E:\UserCreation\SMB_UserCreation_Staff_Report.csv"

# Check if the CSV file exists
if (-not (Test-Path $CSVFile)) {
    Write-Host "CSV file not found: $CSVFile" -ForegroundColor Red
    Out-File -FilePath $LogFile -Input "CSV file not found: $CSVFile $Time on $Date$($CRLF)$Separator$($CRLF)$Separator" -Encoding Unicode -Append -NoClobber
    Break
}

# Check if the variables are changed correctly
if (([System.Windows.Forms.MessageBox]::Show("Have you changed the variables correctly?", 'Variables', 'YesNo', 'Question')) -eq "No") {
    Write-Host "Please change the variables at the top of this script" -ForegroundColor Red
    Out-File -FilePath $LogFile -Input "Please change the variables at the top of this script $Time on $Date$($CRLF)$Separator$($CRLF)$Separator" -Encoding Unicode -Append -NoClobber
    Break
}

# Check if the CSV file has blank lines
if (([System.Windows.Forms.MessageBox]::Show("Have you checked the CSV for blank lines?", 'Variables', 'YesNo', 'Question')) -eq "No") {
    Write-Host "Please check the CSV for blank lines" -ForegroundColor Red
    Out-File -FilePath $LogFile -Input "Please check the CSV for blank lines $Time on $Date$($CRLF)$Separator$($CRLF)$Separator" -Encoding Unicode -Append -NoClobber
    Break
}

Out-File -FilePath $LogFile -Input "Variables declared at $Time $Date$($CRLF)CSV File: $CSVFile$($CRLF)" -Encoding Unicode -Append -NoClobber

# Import Active Directory Module for running CMDlets
Import-Module ActiveDirectory | Out-Null
Out-File -FilePath $LogFile -Input "Active Directory Module Imported$($CRLF)" -Encoding Unicode -Append -NoClobber
Write-Host "Active Directory Module Imported$($CRLF)$($Separator)$($CRLF)" -ForegroundColor Green

# Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-CSV -Delimiter "," -Path $CSVFile
Out-File -FilePath $LogFile -Input "Users Imported into variable ADUsers from file: $CSVFile$($CRLF)$($Separator)" -Encoding Unicode -Append -NoClobber
Write-Host "Users Imported into variable ADUsers from file: $CSVFile$($CRLF)$($Separator)$($CRLF)" -ForegroundColor Green

# Loop through each row containing user details in the CSV file
ForEach ($User in $ADUsers) {
    $DateTime = Get-Date
    $Date = $DateTime.ToShortDateString()
    $Time = $DateTime.ToShortTimeString()

    # Read user data from each field in each row and assign the data to a variable
    $Firstname             = $User.Firstname
    $Lastname              = $User.Lastname
    $Username              = $User.Username
    $Password              = ConvertTo-SecureString $User.Password -AsPlainText -Force
    $OU                    = $User.OU
    $EmailAddress          = "$Username@$($User.EmailDomain)"
    $HomeDirectory         = "$($User.NewHomeDirectory)\$Username"
    $HomeDrive             = $User.HomeDrive
    $ChangePasswordAtLogon = $false
    $Description           = $User.Description
    $HomeDriveSecurity     = "$($User.NewHomeDirectoryLocalPath)\$Username"
    $HomeDirectoryLocalPath = "$($User.NewHomeDirectoryLocalPath)\$Username"
    $DomainName            = $User.DomainName
    $NewHomeDir            = $HomeDirectoryLocalPath

    # Check to see if the user already exists in AD
    if (Get-ADUser -Filter {SamAccountName -eq $Username}) {
        # If user does exist, give a warning
        Write-Warning "A user account with username $Username already exists in Active Directory, appending a digit$($CRLF)"
        Out-File -FilePath $LogFile -Input "$Username Already exists in AD - $OU" -Encoding Unicode -Append -NoClobber

        $Append = 0

        do {
            ++$Append
            $Username = $User.Username + $Append
            $Lastname = $User.Lastname + $Append
            if (Get-ADUser -Filter {SamAccountName -eq $Username}) {
                $CheckAD = 'True'
            } else {
                $CheckAD = 'False'
            }
        } until ($CheckAD -eq 'False')

        $HomeDirectory         = "$($User.NewHomeDirectory)\$Username"
        $HomeDriveSecurity     = "$($User.NewHomeDirectoryLocalPath)\$Username"
        $HomeDirectoryLocalPath = "$($User.NewHomeDirectoryLocalPath)\$Username"
        $EmailAddress          = "$Username@$($User.EmailDomain)"
        $NewHomeDir            = $HomeDirectoryLocalPath

        Write-Host "Username for above User has been changed to $Username" -ForegroundColor DarkYellow
        Out-File -FilePath $LogFile -Input "Username for above User has been changed to $Username$($CRLF)" -Encoding Unicode -Append -NoClobber
    }

    # Create random print code and check if in use in AD already
    $PrintCode = Get-Random -Minimum 20000 -Maximum 89999

    if (Get-ADUser -Filter {pager -eq $PrintCode}) {
        do {
            $PrintCode = Get-Random -Minimum 20000 -Maximum 89999
            if (Get-ADUser -Filter {pager -eq $PrintCode}) {
                $CheckADPrintCode = 'True'
            } else {
                $CheckADPrintCode = 'False'
            }
        } until ($CheckADPrintCode -eq 'False')
    }

    # User does not exist then proceed to create the new user account
    # Account will be created in the OU provided by the $OU variable read from the CSV file
New-ADUser `
	   -SamAccountName $Username `
           -UserPrincipalName "$Username@$DomainName" `
           -Name "$Firstname $Lastname" `
           -GivenName $Firstname `
           -Surname $Lastname `
           -Enabled $true `
           -DisplayName "$Firstname $Lastname" `
           -Path $OU `
           -EmailAddress $EmailAddress `
           -HomeDirectory $HomeDirectory `
           -HomeDrive $HomeDrive `
           -AccountPassword $Password `
           -ChangePasswordAtLogon $ChangePasswordAtLogon `
           -Description $Description

   # Check to see if the user exists in AD, if so carry on, if not give a warning (at bottom)
if (Get-ADUser -Filter {SamAccountName -eq $Username}) {
    # If user does exist, give a message
    Write-Host "User $Username has been created" -ForegroundColor Green
    $n++
    Out-File -FilePath $LogFile -Input "Username for is $Username$($CRLF)" -Encoding Unicode -Append -NoClobber

    # Set Pager field for Print Code
    Set-ADUser $UserName -Replace @{Pager = $PrintCode}

    Write-Host "Print Code for $Username is $PrintCode" -ForegroundColor Green
    Out-File -FilePath $LogFile -Input "Print Code for above User is now $PrintCode$($CRLF)" -Encoding Unicode -Append -NoClobber

    # Once user is made, see if Home Directories clash with existing folders and, if not, create them
    if (-not (Test-Path -Path $HomeDirectoryLocalPath)) {
        New-Item $HomeDirectoryLocalPath -Type Directory | Out-Null

        # Test to see if creation is successful
        if (Test-Path -Path $HomeDirectoryLocalPath) {
            Write-Host "Home Directory '$HomeDirectoryLocalPath' created" -ForegroundColor Green
            Out-File -FilePath $LogFile -Input "Home Directory '$HomeDirectoryLocalPath' created$($CRLF)" -Encoding Unicode -Append -NoClobber
        }

        # Set Security on newly created folders
        cmd /c "icacls" $HomeDriveSecurity "/setowner" "${Username}" | Out-Null
        cmd /c "icacls" $HomeDriveSecurity "/grant" "${Username}:(OI)(CI)F" | Out-Null
    } else {
        Write-Warning "A Home Directory folder with the path '$HomeDirectoryLocalPath' already exists. This will now have to be created manually"
        Out-File -FilePath $LogFile -Input "A Home Directory with the path '$HomeDirectoryLocalPath' already exists. This will now have to be created manually$($CRLF)" -Encoding Unicode -Append -NoClobber
    }

    # Output info to CSV
    $ContactObject = New-Object PSObject -Property @{
        FullName                   = "$Firstname $Lastname"
        SAMAccountName             = $Username
        Password                   = $User.password
        Description                = $Description
        "Email Address"            = $EmailAddress
        "Print Code"               = "-$PrintCode-"
        "Time Created"             = "-$Time-   -$Date-"
        "Change Password on first Logon" = $ChangePasswordAtLogon
    }

    $Report += $ContactObject

    Out-File -FilePath $LogFile -Input "$($Separator)$($CRLF)" -Encoding Unicode -Append -NoClobber
    Write-Host "$($CRLF)$($Separator)$($CRLF)"
} else {
    Write-Warning "User account --$Username-- was not created"
    $n++
    Out-File -FilePath $LogFile -Input "ERROR================>> User account --$Username-- was not created$($CRLF)" -Encoding Unicode -Append -NoClobber
}

$Report | Export-Csv "$ReportPath" -NoTypeInformation

Out-File -FilePath $LogFile -Input "$($CRLF)$($Separator)$($CRLF)Script complete at $Time on $Date$($CRLF)$($Separator)" -Encoding Unicode -Append -NoClobber
Write-Host "$($Separator)$($CRLF)Script complete at $Time on $Date$($CRLF)$($Separator)" -ForegroundColor Green
