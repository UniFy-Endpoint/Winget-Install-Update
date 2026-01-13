<#
.SYNOPSIS
    Requirement/Install Script for Winget in System Context (x64 + ARM64 support) with logging.

.DESCRIPTION
    Detects installed Winget version, compares to the latest stable GitHub release,
    installs/updates Winget and required dependencies for the detected architecture.
    Uses DISM.exe for compatibility when Appx module is unavailable.

.PARAMETER TestMode
    When enabled, performs all checks but skips actual installation. Useful for testing.

.NOTES
    Author: Yoennis Olmo
    Version: v1.4
    Release Date: 2025-12-19

.EXAMPLE
    # Normal installation
    .\Winget-Install-Update_v1.ps1

.EXAMPLE
    # Test mode (no installation)
    .\Winget-Install-Update_v1.ps1 -TestMode
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$TestMode
)

# ---------------------------
# Configure Logging
# ---------------------------
$logPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Winget_Install.txt"
$logDir = Split-Path $logPath

if (-not (Test-Path $logDir)) {
    try {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Write-Output "Log directory created: $logDir"
    } catch {
        Write-Output "ERROR: Cannot create log directory: $_"
        exit 1
    }
}

try {
    Start-Transcript -Path $logPath -Append -ErrorAction Stop
} catch {
    Write-Output "WARNING: Could not start transcript logging. $_"
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Output "=== Winget Requirement/Install Script Start ==="
Write-Output "Running as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Output "PowerShell Version: $($PSVersionTable.PSVersion)"
Write-Output "PowerShell Edition: $($PSVersionTable.PSEdition)"

if ($TestMode) {
    Write-Output "*** TEST MODE ENABLED - No installations will be performed ***"
}

# ---------------------------
# Check Installation Method
# ---------------------------
$useAppxCmdlets = $false
$useDISM = $false

# Try to import Appx module
try {
    Import-Module Appx -ErrorAction Stop
    if (Get-Command Get-AppxPackage -ErrorAction SilentlyContinue) {
        $useAppxCmdlets = $true
        Write-Output "Using Appx PowerShell cmdlets for installation."
    }
} catch {
    Write-Output "Appx module not available. Will use DISM.exe instead."
}

# Check if DISM.exe is available
if (-not $useAppxCmdlets) {
    $dismPath = "$env:SystemRoot\System32\dism.exe"
    if (Test-Path $dismPath) {
        $useDISM = $true
        Write-Output "Using DISM.exe for installation."
    } else {
        Write-Output "ERROR: Neither Appx cmdlets nor DISM.exe are available."
        Write-Output "Cannot proceed with installation."
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
}

# ---------------------------
# Detect Architecture
# ---------------------------
$arch = $env:PROCESSOR_ARCHITECTURE
if ($arch -match "ARM64|AARCH64") {
    $arch = "arm64"
} elseif ($arch -match "AMD64|x86_64") {
    $arch = "x64"
} else {
    $regArch = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment").PROCESSOR_ARCHITECTURE
    if ($regArch -match "ARM64|AARCH64") {
        $arch = "arm64"
    } else {
        $arch = "x64"
    }
}
Write-Output "Detected architecture: $arch"

# ---------------------------
# Function to parse version
# ---------------------------
function Get-VersionFromTag {
    param([string]$tag)
    return [version]($tag -replace '^v', '')
}

# ---------------------------
# Detect Installed Winget
# ---------------------------
$installedVersion = $null
$wingetPath = $null

Write-Output "Checking for existing winget installation..."

# Method 1: Try using Get-AppxPackage if available
if ($useAppxCmdlets) {
    try {
        $appxPackage = Get-AppxPackage -Name Microsoft.DesktopAppInstaller -AllUsers -ErrorAction SilentlyContinue
        if ($appxPackage) {
            $wingetPath = Join-Path $appxPackage.InstallLocation "winget.exe"
            if (Test-Path $wingetPath) {
                try {
                    $wingetOutput = & $wingetPath -v 2>$null
                    if ($LASTEXITCODE -eq 0 -and $wingetOutput) {
                        $versionString = $wingetOutput -replace '^v', ''
                        if ($versionString -as [version]) {
                            $installedVersion = [version]$versionString
                            Write-Output "Detected winget version: $installedVersion at $wingetPath"
                        }
                    }
                } catch {
                    Write-Output "Error detecting version from $wingetPath : $_"
                }
            }
        }
    } catch {
        Write-Output "WARNING: Get-AppxPackage failed: $_"
    }
}

# Method 2: Search WindowsApps folder manually
if (-not $installedVersion) {
    $windowsAppsPath = "$env:ProgramFiles\WindowsApps"
    if (Test-Path $windowsAppsPath) {
        $appInstallerFolders = Get-ChildItem -Path $windowsAppsPath -Filter "Microsoft.DesktopAppInstaller*" -Directory -ErrorAction SilentlyContinue
        foreach ($folder in $appInstallerFolders) {
            $wingetPath = Join-Path $folder.FullName "winget.exe"
            if (Test-Path $wingetPath) {
                try {
                    $wingetOutput = & $wingetPath -v 2>$null
                    if ($LASTEXITCODE -eq 0 -and $wingetOutput) {
                        $versionString = $wingetOutput -replace '^v', ''
                        if ($versionString -as [version]) {
                            $installedVersion = [version]$versionString
                            Write-Output "Detected winget version: $installedVersion at $wingetPath"
                            break
                        }
                    }
                } catch {
                    continue
                }
            }
        }
    }
}

# Method 3: Check System32
if (-not $installedVersion) {
    $wingetPath = "$env:SystemRoot\System32\winget.exe"
    if (Test-Path $wingetPath) {
        try {
            $wingetOutput = & $wingetPath -v 2>$null
            if ($LASTEXITCODE -eq 0 -and $wingetOutput) {
                $versionString = $wingetOutput -replace '^v', ''
                if ($versionString -as [version]) {
                    $installedVersion = [version]$versionString
                    Write-Output "Detected winget version: $installedVersion at $wingetPath"
                }
            }
        } catch {
            Write-Output "Error detecting version from $wingetPath : $_"
        }
    }
}

# Method 4: Try PATH
if (-not $installedVersion) {
    try {
        $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetCmd) {
            $wingetOutput = & winget -v 2>$null
            if ($LASTEXITCODE -eq 0 -and $wingetOutput) {
                $versionString = $wingetOutput -replace '^v', ''
                if ($versionString -as [version]) {
                    $installedVersion = [version]$versionString
                    Write-Output "Detected winget version: $installedVersion in PATH"
                }
            }
        }
    } catch {
        Write-Output "Winget not found in PATH."
    }
}

if (-not $installedVersion) {
    Write-Output "Winget not currently installed."
}

# ---------------------------
# Get Latest Winget Release Info
# ---------------------------
Write-Output "Querying GitHub for latest winget release..."
try {
    $releases = Invoke-RestMethod 'https://api.github.com/repos/microsoft/winget-cli/releases?per_page=20' -UseBasicParsing
} catch {
    Write-Output "ERROR: Could not query GitHub API. $_"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$latestRelease = $releases | Where-Object { -not $_.prerelease } |
                 Sort-Object -Property published_at -Descending |
                 Select-Object -First 1

if (-not $latestRelease) {
    Write-Output "ERROR: No stable winget release found."
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$latestVersion = Get-VersionFromTag $latestRelease.tag_name
Write-Output "Latest stable winget version: $latestVersion"

# ---------------------------
# Version Check
# ---------------------------
if ($installedVersion -and $installedVersion -ge $latestVersion) {
    Write-Output "Winget is already up to date (v$installedVersion). No action needed."
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 0
}

Write-Output "Winget needs to be installed or updated."

# ---------------------------
# Download Dependencies and Winget
# ---------------------------
$tempDir = "$env:TEMP\WingetInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}
Write-Output "Using temp directory: $tempDir"

# Find the .msixbundle asset
$msixAsset = $latestRelease.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1
if (-not $msixAsset) {
    Write-Output "ERROR: Could not find .msixbundle in release assets."
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# Find the license file
$licenseAsset = $latestRelease.assets | Where-Object { $_.name -like "*License1.xml" } | Select-Object -First 1
if (-not $licenseAsset) {
    Write-Output "ERROR: Could not find License file in release assets."
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$msixPath = Join-Path $tempDir $msixAsset.name
$licensePath = Join-Path $tempDir $licenseAsset.name

# Download Winget
Write-Output "Downloading winget: $($msixAsset.name)..."
if (-not $TestMode) {
    try {
        Invoke-WebRequest -Uri $msixAsset.browser_download_url -OutFile $msixPath -UseBasicParsing
        Write-Output "Downloaded: $msixPath"
    } catch {
        Write-Output "ERROR: Failed to download winget. $_"
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
} else {
    Write-Output "[TEST MODE] Would download: $($msixAsset.browser_download_url)"
}

# Download License
Write-Output "Downloading license: $($licenseAsset.name)..."
if (-not $TestMode) {
    try {
        Invoke-WebRequest -Uri $licenseAsset.browser_download_url -OutFile $licensePath -UseBasicParsing
        Write-Output "Downloaded: $licensePath"
    } catch {
        Write-Output "ERROR: Failed to download license. $_"
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
} else {
    Write-Output "[TEST MODE] Would download: $($licenseAsset.browser_download_url)"
}

# ---------------------------
# Install Dependencies
# ---------------------------

# VCLibs dependency
Write-Output "Installing VCLibs dependency..."
if ($arch -eq "x64") {
    $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
} elseif ($arch -eq "arm64") {
    $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.arm64.14.00.Desktop.appx"
} else {
    Write-Output "ERROR: Unsupported architecture for VCLibs: $arch"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$vcLibsPath = Join-Path $tempDir "Microsoft.VCLibs.$arch.14.00.Desktop.appx"
if (-not $TestMode) {
    try {
        Invoke-WebRequest -Uri $vcLibsUrl -OutFile $vcLibsPath -UseBasicParsing
        Write-Output "Downloaded VCLibs: $vcLibsPath"

        if ($useAppxCmdlets) {
            Add-AppxPackage -Path $vcLibsPath -ErrorAction Stop
            Write-Output "VCLibs installed successfully (Appx)."
        } elseif ($useDISM) {
            $dismArgs = "/Online /Add-ProvisionedAppxPackage /PackagePath:`"$vcLibsPath`" /SkipLicense"
            $dismResult = Start-Process -FilePath "$env:SystemRoot\System32\dism.exe" -ArgumentList $dismArgs -Wait -PassThru -NoNewWindow
            if ($dismResult.ExitCode -eq 0) {
                Write-Output "VCLibs installed successfully (DISM)."
            } else {
                Write-Output "WARNING: VCLibs installation returned exit code: $($dismResult.ExitCode)"
            }
        }
    } catch {
        Write-Output "WARNING: VCLibs installation failed or already installed. $_"
    }
} else {
    Write-Output "[TEST MODE] Would download and install VCLibs from: $vcLibsUrl"
}

# UI.Xaml dependency
Write-Output "Installing UI.Xaml dependency..."
$uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
if ($arch -eq "arm64") {
    $uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.arm64.appx"
}

$uiXamlPath = Join-Path $tempDir "Microsoft.UI.Xaml.2.8.$arch.appx"
if (-not $TestMode) {
    try {
        Invoke-WebRequest -Uri $uiXamlUrl -OutFile $uiXamlPath -UseBasicParsing
        Write-Output "Downloaded UI.Xaml: $uiXamlPath"

        if ($useAppxCmdlets) {
            Add-AppxPackage -Path $uiXamlPath -ErrorAction Stop
            Write-Output "UI.Xaml installed successfully (Appx)."
        } elseif ($useDISM) {
            $dismArgs = "/Online /Add-ProvisionedAppxPackage /PackagePath:`"$uiXamlPath`" /SkipLicense"
            $dismResult = Start-Process -FilePath "$env:SystemRoot\System32\dism.exe" -ArgumentList $dismArgs -Wait -PassThru -NoNewWindow
            if ($dismResult.ExitCode -eq 0) {
                Write-Output "UI.Xaml installed successfully (DISM)."
            } else {
                Write-Output "WARNING: UI.Xaml installation returned exit code: $($dismResult.ExitCode)"
            }
        }
    } catch {
        Write-Output "WARNING: UI.Xaml installation failed or already installed. $_"
    }
} else {
    Write-Output "[TEST MODE] Would download and install UI.Xaml from: $uiXamlUrl"
}

# ---------------------------
# Install Winget
# ---------------------------
Write-Output "Installing winget package..."
if (-not $TestMode) {
    $installSuccess = $false

    # Method 1: Try DISM first (most reliable for system context)
    if ($useDISM) {
        try {
            Write-Output "Attempting installation with DISM.exe..."
            $dismArgs = "/Online /Add-ProvisionedAppxPackage /PackagePath:`"$msixPath`" /LicensePath:`"$licensePath`""
            Write-Output "DISM Command: dism.exe $dismArgs"

            $dismResult = Start-Process -FilePath "$env:SystemRoot\System32\dism.exe" -ArgumentList $dismArgs -Wait -PassThru -NoNewWindow

            if ($dismResult.ExitCode -eq 0) {
                Write-Output "Winget installed successfully for all users (DISM)."
                $installSuccess = $true
            } else {
                Write-Output "DISM installation returned exit code: $($dismResult.ExitCode)"
            }
        } catch {
            Write-Output "ERROR: DISM installation failed. $_"
        }
    }

    # Method 2: Try Appx cmdlets if DISM failed
    if (-not $installSuccess -and $useAppxCmdlets) {
        try {
            Write-Output "Attempting installation with Add-AppxProvisionedPackage..."
            Add-AppxProvisionedPackage -Online -PackagePath $msixPath -LicensePath $licensePath -ErrorAction Stop
            Write-Output "Winget installed successfully for all users (Appx)."
            $installSuccess = $true
        } catch {
            Write-Output "ERROR: Add-AppxProvisionedPackage failed. $_"

            # Fallback: Try Add-AppxPackage
            try {
                Write-Output "Attempting fallback with Add-AppxPackage..."
                Add-AppxPackage -Path $msixPath -ErrorAction Stop
                Write-Output "Winget installed successfully (current user)."
                $installSuccess = $true
            } catch {
                Write-Output "ERROR: Add-AppxPackage also failed. $_"
            }
        }
    }

    if (-not $installSuccess) {
        Write-Output "ERROR: All installation methods failed."
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
} else {
    Write-Output "[TEST MODE] Would install winget using:"
    if ($useDISM) {
        Write-Output "  DISM.exe /Online /Add-ProvisionedAppxPackage /PackagePath:`"$msixPath`" /LicensePath:`"$licensePath`""
    } else {
        Write-Output "  Add-AppxProvisionedPackage -Online -PackagePath $msixPath -LicensePath $licensePath"
    }
}

# ---------------------------
# Verify Installation
# ---------------------------
if (-not $TestMode) {
    Write-Output "Verifying installation..."
    Start-Sleep -Seconds 5

    # Try multiple verification methods
    $verified = $false

    # Method 1: Check with Get-AppxPackage
    if ($useAppxCmdlets) {
        try {
            $appxPackage = Get-AppxPackage -Name Microsoft.DesktopAppInstaller -AllUsers -ErrorAction SilentlyContinue
            if ($appxPackage) {
                $wingetPath = Join-Path $appxPackage.InstallLocation "winget.exe"
                if (Test-Path $wingetPath) {
                    $wingetOutput = & $wingetPath -v 2>$null
                    Write-Output "Winget verification: $wingetOutput"
                    $verified = $true
                }
            }
        } catch {
            Write-Output "Verification method 1 failed: $_"
        }
    }

    # Method 2: Search WindowsApps folder
    if (-not $verified) {
        $windowsAppsPath = "$env:ProgramFiles\WindowsApps"
        $appInstallerFolders = Get-ChildItem -Path $windowsAppsPath -Filter "Microsoft.DesktopAppInstaller*" -Directory -ErrorAction SilentlyContinue
        foreach ($folder in $appInstallerFolders) {
            $wingetPath = Join-Path $folder.FullName "winget.exe"
            if (Test-Path $wingetPath) {
                try {
                    $wingetOutput = & $wingetPath -v 2>$null
                    Write-Output "Winget verification: $wingetOutput"
                    $verified = $true
                    break
                } catch {
                    continue
                }
            }
        }
    }

    if ($verified) {
        Write-Output "Installation completed successfully!"
    } else {
        Write-Output "WARNING: Could not verify winget installation. It may require a system restart."
    }
}

# ---------------------------
# Cleanup
# ---------------------------
if (-not $TestMode) {
    Write-Output "Cleaning up temporary files..."
    try {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Cleanup completed."
    } catch {
        Write-Output "WARNING: Could not clean up temp directory. $_"
    }
} else {
    Write-Output "[TEST MODE] Would clean up: $tempDir"
}

Write-Output "=== Winget Requirement/Install Script Complete ==="
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

if ($TestMode) {
    Write-Output "`n*** TEST MODE - No changes were made to the system ***"
}

exit 0
