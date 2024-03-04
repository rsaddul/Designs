# Prompt box to confirm whether the folder paths have been updated
$Confirmation = [System.Windows.Forms.MessageBox]::Show(
    "Have you updated the folder paths? Click 'Yes' to continue with the analysis or 'No' to exit.",
    "Folder Paths Update Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($Confirmation -eq 'No') {
    Write-Host "Analysis aborted. Please update the folder paths and run the script again."
    exit
}

# List of folders to analyze
$FolderPaths = @(
    "\\NHJ-FILE01\NHJSharedData$\Office"
    # Add more folder paths as needed
)

# Set the path limit (260 characters for NTFS)
$PathLimit = 260

# Create an array to store results
$Results = @()

foreach ($FolderPath in $FolderPaths) {
    Write-Host "Analyzing folder: $FolderPath"

    # Get all files and subfolders recursively
    $Items = Get-ChildItem -Path $FolderPath -Recurse

    # Check each item for path length
    foreach ($Item in $Items) {
        $ItemPathLength = $Item.FullName.Length
        if ($ItemPathLength -gt $PathLimit) {
            $Results += [PSCustomObject]@{
                'Folder' = $FolderPath
                'Item'   = $Item.FullName
                'Length' = $ItemPathLength
            }
        }
    }
}

# Output results to CSV file
$Results | Export-Csv -Path 'C:\PathExceedingLimit.csv' -NoTypeInformation

Write-Host "Analysis complete. Results exported to PathExceedingLimit.csv"