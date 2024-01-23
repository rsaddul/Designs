# Prompt box to confirm whether the folder paths array has been updated
$Confirmation = [System.Windows.Forms.MessageBox]::Show(
    "Have you updated the folder paths array? Click 'Yes' to continue with the analysis or 'No' to exit.",
    "Folder Paths Update Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($Confirmation -eq 'No') {
    Write-Host "Analysis aborted. Please update the folder paths array and run the script again."
    exit
}

# List of folders to analyze
$FolderPaths = @(
    "\\NHJ-FILE01\NHJSharedData$\Curriculum",
    "\\NHJ-FILE01\NHJSharedData$\Staffroom",
    "\\NHJ-FILE01\NHJSharedData$\Media",
    "\\NHJ-FILE01\NHJSharedData$\Pupil",
    "\\NHJ-FILE01\NHJSharedData$\Safeguarding",
    "\\NHJ-FILE01\NHJSharedData$\SLT",
    "\\NHJ-FILE01\NHJSharedData$\Office"
)

foreach ($FolderPath in $FolderPaths) {
    Write-Host "Analyzing folder: $FolderPath"

    # Calculate the total size of the folder in MB
    $TotalSize = (Get-ChildItem -Path $FolderPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum

    if ($TotalSize -ge 1GB) {
        # Convert to GB if the size is 1GB or larger
        $TotalSizeFormatted = [math]::Round($TotalSize / 1GB, 2)
        $SizeUnit = "GB"
    } else {
        # Otherwise, keep the size in MB
        $TotalSizeFormatted = [math]::Round($TotalSize / 1MB, 2)
        $SizeUnit = "MB"
    }

    # Assuming you want to move files older than 3 years based on LastAccessTime
    $ThreeYearsAgo = (Get-Date).AddYears(-3)

    # Calculate the size of files older than 3 years based on LastAccessTime in MB
    $OldFilesSize = (Get-ChildItem -Path $FolderPath -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastAccessTime -lt $ThreeYearsAgo } -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum

    if ($OldFilesSize -ge 1GB) {
        # Convert to GB if the size is 1GB or larger
        $OldFilesSizeFormatted = [math]::Round($OldFilesSize / 1GB, 2)
        $OldFilesSizeUnit = "GB"
    } else {
        # Otherwise, keep the size in MB
        $OldFilesSizeFormatted = [math]::Round($OldFilesSize / 1MB, 2)
        $OldFilesSizeUnit = "MB"
    }

    # Calculate the potential space saved in MB
    $SpaceSaved = $TotalSize - $OldFilesSize

    if ($SpaceSaved -ge 1GB) {
        # Convert to GB if the size is 1GB or larger
        $SpaceSavedFormatted = [math]::Round($SpaceSaved / 1GB, 2)
        $SpaceSavedUnit = "GB"
    } else {
        # Otherwise, keep the size in MB
        $SpaceSavedFormatted = [math]::Round($SpaceSaved / 1MB, 2)
        $SpaceSavedUnit = "MB"
    }

    # Display the results for each folder
    Write-Host "Total size of data in the folder: $($TotalSizeFormatted) $SizeUnit" -ForegroundColor Cyan
    Write-Host "Size of data older than 3 years: $($OldFilesSizeFormatted) $OldFilesSizeUnit" -ForegroundColor Red
    Write-Host "Potential space saved by moving old data: $($SpaceSavedFormatted) $SpaceSavedUnit" -ForegroundColor Green
    Write-Host "----------------------------------------"
}
