# Run ADUserExport.ps1 (include objectguid)
# Copy CSV to downloads
# Make sure to hash out "Set-AzureADUser -ObjectId $AzureADUserObjectID -ImmutableId $ConvertedADObjectGUID" for testing
# Once run sucesfully run the AD Sync from Entra

#Run this on the machine with module AzureAD installed
#Declare variables
$Report =@()
$DateTime = Get-Date
$Date = $DateTime.ToShortDateString()
$Time = $DateTime.ToLongTimeString()
$DateTime.ToLongTimeString()
$Date = $Date.Replace("/","-")
$Time = $Time.Replace(":","-")
$CRLF = "`r`n"
$Separator = "==========================================================================================="
$n = 1

#Change these variables to suit
$CSVFile = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Migrations\Azure\User_Creation\NGA_Staff_AD_User_Export_NewDomain.csv"
$LogFile = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Migrations\Azure\User_Creation\NGA_Staff_AD_User_Export_NewDomain_Log.log"
#$ReportPath = "E:\UserCreation\SJP_AzureHardMatch_Staff_Report.csv"

If(([System.Windows.Forms.MessageBox]::Show("Have you installed AzureAD module?",'Variables','YesNo','Question')) -eq "No"){Write-Host "Please Install-Module AzureAD" -ForegroundColor Red;Out-File -filepath $LogFile -input "Please Install-Module AzureAD $Time on $Date$($CRLF)$Separator$($CRLF)$Separator" -encoding "Unicode" -append -noclobber;Break}

If(([System.Windows.Forms.MessageBox]::Show("Have you changed the variables correctly?",'Variables','YesNo','Question')) -eq "No"){Write-Host "Please change the variables at the top of this script" -ForegroundColor Red;Out-File -filepath $LogFile -input "Please change the variables at the top of this script $Time on $Date$($CRLF)$Separator$($CRLF)$Separator" -encoding "Unicode" -append -noclobber;Break}

If(([System.Windows.Forms.MessageBox]::Show("Have you checked the csv for blank lines?",'Variables','YesNo','Question')) -eq "No"){Write-Host "Please check the csv for blank lines" -ForegroundColor Red;Out-File -filepath $LogFile -input "Please check the csv for blank lines $Time on $Date$($CRLF)$Separator$($CRLF)$Separator" -encoding "Unicode" -append -noclobber;Break}

Out-File -filepath $LogFile -input "Variables declared at $Time $Date$($CRLF)CSV File: $CSVFile$($CRLF)" -encoding "Unicode" -append -noclobber

# Import CSV file and set variables

$ADUsers = Import-CSV -Delimiter "," -Path $CSVFile
Out-File -filepath $LogFile -input "Users Imported into variable ADUsers from file: $CSVFile$($CRLF)$($Separator)" -encoding "Unicode" -append -noclobber
Write-Host "Users Imported into variable ADUsers from file: $CSVFile$($CRLF)$($Separator)$($CRLF)" -ForegroundColor Green

# Interactively connect to AzureAD, to avoid the MFA issues with PSCredential objects
Out-File -filepath $LogFile -input "Interactively connect to AzureAD$($CRLF)" -encoding "Unicode" -append -noclobber

#Connect-AzureAD

#Loop through each row containing user details in the CSV file 
ForEach ($User in $ADUsers)
{

    $ADUsername = $User.SAMAccountName
    $ADObjectGUID = $User.ObjectGUID
    $EmailAddress = $User.mail

    Try{
    # Search for the Azure User using:
    Write-Host "Finding user: $($EmailAddress)"
    Out-File -filepath $LogFile -input "Interactively connect to AzureAD$($CRLF)" -encoding "Unicode" -append -noclobber

    $AzureADUser = Get-AzureADUser -ObjectID $EmailAddress

    $AzureADUserObjectID = $AzureADUser.ObjectID
    Write-Host "$($EmailAddress) has this Azure ObjectID: $($AzureADUserObjectID)"
    Out-File -filepath $LogFile -input "$($EmailAddress) has this Azure ObjectID: $($AzureADUserObjectID)" -encoding "Unicode" -append -noclobber

    # Convert ObjectGUID using the following command:

    $ConvertedADObjectGUID = [Convert]::ToBase64String([guid]::New($($ADObjectGUID)).ToByteArray())
    Write-Host "Converted Azure ObjectID: $($ConvertedADObjectGUID)"
    Out-File -filepath $LogFile -input "Converted Azure ObjectID: $($ConvertedADObjectGUID)" -encoding "Unicode" -append -noclobber

    # Set the Azure ImmutableID using the following command:

    Write-Host "Setting Immutable ID for $($EmailAddress) to $($ConvertedADObjectGUID)"
    Out-File -filepath $LogFile -input "Setting Immutable ID for $($EmailAddress) to $($ConvertedADObjectGUID)" -encoding "Unicode" -append -noclobber
    # Set-AzureADUser -ObjectId $AzureADUserObjectID -ImmutableId $ConvertedADObjectGUID

    # Sync the account and wait

    }Catch{
        
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Out-File $LogFile -Append -NoClobber -InputObject "Error for user: $($EmailAddress)$($CRLF)$ErrorMessage"
        Write-Warning "Error for user: $($EmailAddress)"

    }
    
    
}