<#
Developed by: Rhys Saddul
#>

# Check if the Microsoft Teams PowerShell module is installed
if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
    # If not installed, install the module
    Install-Module -Name MicrosoftTeams -Force -AllowClobber -Scope CurrentUser -Repository PSGallery -ErrorAction Continue
}

# Import the Microsoft Teams PowerShell module
Import-Module MicrosoftTeams

# Prompt to ensure varibles have been setup correctly prior to running the script
$dialogResult = [System.Windows.Forms.MessageBox]::Show("Have you set the varible for the output file path?", "Variable Setup", "YesNo", "Question")
if ($dialogResult -eq "Yes") {
    Write-Host "User confirmed variables have been set. Proceeding with the script." -ForegroundColor Green
} else {
    Write-Host "User confirmed Variables have not been set. Exiting the script." -ForegroundColor Red
    return
}

# Varible for the output file path
$outputFilePath = "C:\Users\RhysSaddul\OneDrive - eduthing\Documents\Exports\Teams\Teams_Export.csv"

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Get all Teams
$teams = Get-Team 

$totalTeams = $teams.Count

# Create an array to store Teams information
$teamsInfo = @()

foreach ($team in $teams) {
    $teamPrivacy = $team.Visibility
    $groupId = $team.GroupId  # Store GroupId for querying team details
    $teamDetails = Get-Team -GroupId $groupId  # Get additional details of the team
    $teamName = $teamDetails.DisplayName  # Retrieve team name from team details
    $teamMembers = Get-TeamUser -GroupId $groupId | Select-Object -ExpandProperty User
    $channels = Get-TeamChannel -GroupId $groupId

    foreach ($channel in $channels) {
        $channelName = $channel.DisplayName
        $channelMembers = Get-TeamChannelUser -GroupId $groupId -DisplayName $channel.DisplayName | Select-Object -ExpandProperty User

        $teamMemberNames = $teamMembers
        $channelMemberNames = $channelMembers

        $teamsInfo += [PSCustomObject]@{
            "Privacy" = $teamPrivacy
            "Team Name" = $teamName
            "Team Email" = $groupId  # Use GroupId as team email substitute
            "Member Name" = $teamMemberNames -join ", "
            "Channel Name" = $channelName
            "Channel Members" = $channelMemberNames -join ", "
        }
    }
}

# Export Teams information to CSV
$teamsInfo | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Teams information exported to $outputFilePath" -ForegroundColor Green
