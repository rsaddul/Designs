# Create a tag file just so Intune knows this was installed
if (-not (Test-Path "$($env:ProgramData)\Microsoft\RemoveWindowsBloatware"))
{
    Mkdir "$($env:ProgramData)\Microsoft\RemoveWindowsBloatware"
}
Set-Content -Path "$($env:ProgramData)\Microsoft\RemoveWindowsBloatware\RemoveWindowsBloatware.ps1.tag" -Value "Installed"

# Start logging
Start-Transcript "$($env:ProgramData)\Microsoft\RemoveWindowsBloatware\RemoveWindowsBloatware.log"

# List of built-in apps to remove
$UninstallPackages = @(
	"Microsoft.549981C3F5F10"
	"Microsoft.BingNews"
	"Microsoft.BingWeather"
	"Microsoft.GamingApp"
	"Microsoft.GetHelp"
	"Microsoft.Getstarted"
	"Microsoft.Microsoft3DViewer"
	"Microsoft.MicrosoftOfficeHub"
	"Microsoft.MicrosoftSolitaireCollection"
	"Microsoft.MicrosoftStickyNotes"
	"Microsoft.MixedReality.Portal"
	"Microsoft.Office.OneNote"
	"Microsoft.Office.Onenote"
	"Microsoft.OneConnect"
	"Microsoft.OneDriveSync"
	"Microsoft.People"
	"Microsoft.SkypeApp"
	"Microsoft.Windows.Alarms"
	"Microsoft.Windows.DevHome"
	"Microsoft.WindowsFeedbackHub"
	"Microsoft.WindowsMaps"
	"Microsoft.Xbox.TCUI"
	"Microsoft.XboxApp"
	"Microsoft.XboxIdentityProvider"
	"Microsoft.XboxSpeechToTextOverlay"
	"Microsoft.YourPhone"
	"Microsdft.Wallet"
	"Microsoft.ZuneMusic"
	"Microsoft.ZuneVideo"
	"Microsoft.windowscommunicationsapps"
	"MicrosoftTeams"
)

$InstalledPackages = Get-AppxPackage -AllUsers | Where {($UninstallPackages -contains $_.Name)}

$ProvisionedPackages = Get-AppxProvisionedPackage -Online | Where {($UninstallPackages -contains $_.DisplayName)}

$InstalledPrograms = Get-Package | Where {$UninstallPrograms -contains $_.Name}

# Remove provisioned packages first
ForEach ($ProvPackage in $ProvisionedPackages) {

    Write-Host -Object "Attempting to remove provisioned package: [$($ProvPackage.DisplayName)]..."

    Try {
        $Null = Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -ErrorAction Stop
        Write-Host -Object "Successfully removed provisioned package: [$($ProvPackage.DisplayName)]"
    }
    Catch {Write-Warning -Message "Failed to remove provisioned package: [$($ProvPackage.DisplayName)]"}
}

# Remove appx packages
ForEach ($AppxPackage in $InstalledPackages) {
                                            
    Write-Host -Object "Attempting to remove Appx package: [$($AppxPackage.Name)]..."

    Try {
        $Null = Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction Stop
        Write-Host -Object "Successfully removed Appx package: [$($AppxPackage.Name)]"
    }
    Catch {Write-Warning -Message "Failed to remove Appx package: [$($AppxPackage.Name)]"}
}

# Remove installed programs
$InstalledPrograms | ForEach {

    Write-Host -Object "Attempting to uninstall: [$($_.Name)]..."

    Try {
        $Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
        Write-Host -Object "Successfully uninstalled: [$($_.Name)]"
    }
    Catch {Write-Warning -Message "Failed to uninstall: [$($_.Name)]"}
}

Stop-Transcript
