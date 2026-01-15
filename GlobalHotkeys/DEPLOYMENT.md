# GlobalHotkeys Deployment

This document explains how to publish and deploy GlobalHotkeys to your system.

## Overview

After successfully publishing the project, you can use the provided script to:
1. Kill any running instances of `GlobalHotkeys.exe`
2. Copy the published single-file executable to `C:\Program Files\GlobalHotkeys`
3. Optionally restart the app

## Prerequisites

- **Administrator privileges** - Required to write to `C:\Program Files`
- **Successful build** - Ensure the project builds without errors
- **Windows 10/11** - The scripts are designed for modern Windows systems

## Deployment Script

- **File**: `scripts/publish-and-deploy.ps1`
- **What it does**: `dotnet publish` (single-file) + stop/copy/restart
- **Usage**: Run it normally; it will self-elevate to admin when needed.

## Quick Deployment Steps

1. **Publish + deploy**:
   ```powershell
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\scripts\publish-and-deploy.ps1 -Configuration Release
   ```

2. **Verify installation**:
   - Check `C:\Program Files\GlobalHotkeys\GlobalHotkeys.exe`

## What Gets Deployed

The script deploys a single executable:
- `GlobalHotkeys.exe` - single-file published executable

Notes:
- `ProcessList.txt` and `SoundDevices.txt` are expected to be present under the install directory. If you prefer to keep them elsewhere, use symlinks/junctions pointing into `C:\Program Files\GlobalHotkeys`.

## Troubleshooting

### "Access Denied" Errors
- Ensure you're running as Administrator
- Check if antivirus software is blocking the operation
- Verify the target directory permissions

### "Publish completed but no .exe found" Error
- Ensure the publish step is succeeding and producing a single-file executable.

### Process Termination Issues
- Some processes may be protected by Windows
- Try closing the application manually before deployment
- Restart the deployment script

## Manual Deployment

If the scripts don't work, you can manually:

1. Stop any running GlobalHotkeys.exe processes
2. Create `C:\Program Files\GlobalHotkeys` directory
3. Copy the published `GlobalHotkeys.exe` into the target directory
4. Ensure `ProcessList.txt` and `SoundDevices.txt` are available in the install directory (copy or symlink)

## Post-Deployment

After successful deployment:
- The application is available at `C:\Program Files\GlobalHotkeys\GlobalHotkeys.exe`
- You can run the application from that location
- The application will use the configuration files in the installation directory

## Updating

To update an existing installation:
1. Build the new version
2. Run the deployment script again
3. The script will automatically replace the old files
4. Any running instances will be terminated and replaced

## Security Notes

- The deployment scripts require administrator privileges
- Files are copied to `C:\Program Files` which is a protected system directory
- Only run these scripts from trusted sources
- The scripts include error checking and validation
