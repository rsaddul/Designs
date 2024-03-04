# Prompt box to confirm whether the folder paths have been updated
$Confirmation = [System.Windows.Forms.MessageBox]::Show(
    "Have you updated the folder paths? Click 'Yes' to continue with removing empty folders or 'No' to exit.",
    "Folder Paths Update Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($Confirmation -eq 'No') {
    Write-Host "Empty folder removal aborted. Please update the folder paths and run the script again."
    exit
}

# Specify the folder paths where you want to start deleting empty folders
$FolderPaths = @(
    "C:\Test"
)

# Function to remove empty folders
function Remove-EmptyFolders {
    param (
        [string]$rootDirectory
    )

    # Specify log file paths
    $removalLogPath = "C:\RemovalLog.txt"
    $skippedLogPath = "C:\SkippedLog.txt"

    # Get a list of all directories in the root directory
    $directories = Get-ChildItem -Path $rootDirectory -Directory

    # Loop through each directory in reverse order
    for ($i = $directories.Count - 1; $i -ge 0; $i--) {
        $directory = $directories[$i]

        # Recursively call the function on each subdirectory
        Remove-EmptyFolders -rootDirectory $directory.FullName

        # Check if the directory is empty
        $isEmpty = @(Get-ChildItem -Path $directory.FullName -File).Count -eq 0 -and @(Get-ChildItem -Path $directory.FullName -Directory).Count -eq 0

        # Output information about the directory
        Write-Host "Checking folder: $($directory.FullName)"

        # If the directory is empty, log removal and remove it
        if ($isEmpty) {
            Write-Host "Removing empty folder: $($directory.FullName)" -ForegroundColor Red
            $directory.FullName | Out-File -Append -FilePath $removalLogPath
            Remove-Item -Path $directory.FullName -Force
        }
        else {
            Write-Host "Skipped non-empty folder: $($directory.FullName)"
            $directory.FullName | Out-File -Append -FilePath $skippedLogPath
        }
    }
}

# Call the function for each folder path to remove empty folders
foreach ($FolderPath in $FolderPaths) {
    Remove-EmptyFolders -rootDirectory $FolderPath
}
