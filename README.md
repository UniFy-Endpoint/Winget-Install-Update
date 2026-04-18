# Winget Install/Update Script for Windows Autopilot

Automated PowerShell script for installing and updating **Windows Package Manager (winget)** in **SYSTEM context** during Windows Autopilot and Autopilot Device Preparation deployments.

## Features

- ✅ **System Context Installation** - Works in SYSTEM context for Autopilot/Intune deployments
- ✅ **Architecture Detection** - Automatic support for x64 and ARM64 systems
- ✅ **Version Management** - Detects installed version and updates only when needed
- ✅ **Dependency Handling** - Automatically installs VCLibs and UI.Xaml dependencies
- ✅ **Test Mode** - Dry-run capability for testing without making changes
- ✅ **Comprehensive Logging** - Detailed logs for troubleshooting
- ✅ **Error Handling** - Robust error handling with fallback methods

## Requirements

- Windows 10 (1809+) or Windows 11
- PowerShell 5.1 or later
- Internet connectivity
- Administrator privileges (SYSTEM context for Autopilot)

## Quick Start

### Standard Installation
```powershell
.\Winget-Install-Update_v1.ps1
```

### Test Mode (No Installation)
```powershell
.\Winget-Install-Update_v1.ps1 -TestMode
```

## 📦 What Gets Installed

1. **Microsoft.DesktopAppInstaller** (winget) - Latest stable release
2. **VCLibs** - Visual C++ Runtime Libraries (architecture-specific)
3. **UI.Xaml** - Microsoft UI framework (architecture-specific)
4. **License File** - Required for system-wide provisioning

## Deployment Methods

### Intune (Microsoft Endpoint Manager)

1. **Create Win32 App Package:**
   - Package the script using [Microsoft Win32 Content Prep Tool](https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool)
   - Install command: `powershell.exe -ExecutionPolicy Bypass -File Winget-Install-Update_v1.ps1`
   - Uninstall command: `cmd.exe /c echo "Not applicable"`

2. **Detection Rule:**
   - Rule type: **File**
   - Path: `%ProgramFiles%\WindowsApps`
   - File/folder: `Microsoft.DesktopAppInstaller_*`
   - Detection method: **Folder exists**

3. **Assignment:**
   - Assign to device groups
   - Install during Autopilot ESP (Enrollment Status Page)

### Autopilot Device Preparation

1. Upload script to Intune
2. Assign to Autopilot profile
3. Script runs in SYSTEM context during device setup

### Manual Deployment
```powershell
# Run as Administrator
Set-ExecutionPolicy Bypass -Scope Process -Force
.\Winget-Install-Update_v1.ps1
```

## Logging

Logs are automatically created at:
```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Winget_Install.txt
```

**Log Contents:**
- Script execution details
- Architecture detection
- Version comparison
- Download progress
- Installation results
- Error messages

## Troubleshooting

### Issue: Script fails to download from GitHub

**Solution:** Check internet connectivity and proxy settings
```powershell
# Test GitHub API access
Invoke-RestMethod 'https://api.github.com/repos/microsoft/winget-cli/releases/latest'
```

### Issue: "Add-AppxProvisionedPackage" fails

**Solution:** The script includes a fallback to `Add-AppxPackage`. Check logs for specific error details.

### Issue: Winget not found after installation

**Solution:** 
1. Restart PowerShell session
2. Check if package is installed: `Get-AppxPackage -Name Microsoft.DesktopAppInstaller -AllUsers`
3. Verify path: `$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*`

### Issue: Dependencies fail to install

**Solution:** Dependencies (VCLibs, UI.Xaml) may already be installed. Check logs - warnings are normal if already present.

## How It Works

1. **Architecture Detection** - Identifies x64 or ARM64 system
2. **Version Check** - Queries GitHub API for latest stable release
3. **Comparison** - Compares installed version (if any) with latest
4. **Download** - Downloads winget, license, and dependencies to temp folder
5. **Installation** - Installs dependencies first, then winget with license
6. **Verification** - Confirms successful installation
7. **Cleanup** - Removes temporary files

## Script Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TestMode` | Switch | `$false` | Enables test mode - performs checks without installation |

## Security Considerations

- Script uses TLS 1.2 for secure downloads
- Downloads only from official Microsoft sources:
  - GitHub (microsoft/winget-cli)
  - Microsoft CDN (aka.ms)
- Validates file downloads before installation
- Runs with minimal required permissions

## 📌 Version History

### v1.2 (2025-12-01)
- ✅ Fixed incomplete installation logic
- ✅ Added dependency installation (VCLibs, UI.Xaml)
- ✅ Added license file handling
- ✅ Fixed path detection using Get-AppxPackage
- ✅ Added TestMode parameter
- ✅ Improved error handling with fallback methods
- ✅ Added installation verification
- ✅ Enhanced logging

### v1.1 (2025-08-15)
- Initial release with basic functionality

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Microsoft winget team for the excellent package manager
- Community contributors for testing and feedback

## Additional Resources

- [Winget Official Documentation](https://learn.microsoft.com/windows/package-manager/)
- [Windows Autopilot Documentation](https://learn.microsoft.com/mem/autopilot/)
- [Intune Win32 App Management](https://learn.microsoft.com/mem/intune/apps/apps-win32-app-management)

## Support

For issues and questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Review logs at `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\Winget_Install.txt`

---

**Made with ❤️ for the Windows Autopilot community**
