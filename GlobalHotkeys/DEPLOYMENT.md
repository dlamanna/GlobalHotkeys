# GlobalHotkeys2 Deployment

This document explains how to deploy GlobalHotkeys2 to your system after building the project.

## Overview

After successfully building the project, you can use the provided deployment scripts to:
1. Kill any running instances of GlobalHotkeys2.exe
2. Copy the executable and supporting files to `C:\Program Files\GlobalHotkeys`

## Prerequisites

- **Administrator privileges** - Required to write to `C:\Program Files`
- **Successful build** - Ensure the project builds without errors
- **Windows 10/11** - The scripts are designed for modern Windows systems

## Deployment Scripts

### Option 1: PowerShell Script (Recommended)
- **File**: `post-build-deploy.ps1`
- **Features**: Better error handling, colored output, progress indicators
- **Usage**: Right-click and "Run as administrator"

### Option 2: Batch File
- **File**: `post-build-deploy.bat`
- **Features**: Simple, works on all Windows systems
- **Usage**: Right-click and "Run as administrator"

### Option 3: Launcher Batch File
- **File**: `run-deploy-as-admin.bat`
- **Features**: Launches the PowerShell script with proper privileges
- **Usage**: Double-click (will prompt for admin credentials)

## Quick Deployment Steps

1. **Build the project**:
   ```bash
   dotnet build --configuration Release
   ```

2. **Run deployment script**:
   - Right-click on `post-build-deploy.ps1`
   - Select "Run as administrator"
   - Follow the on-screen prompts

3. **Verify installation**:
   - Check `C:\Program Files\GlobalHotkeys\GlobalHotkeys2.exe`

## What Gets Deployed

The deployment script copies the following files:
- `GlobalHotkeys2.exe` - Main executable
- `*.dll` - All required DLL files (NAudio, Keyboard, etc.)
- `*.config` - Configuration files
- `*.json` - Runtime configuration files
- `ProcessList.txt` - List of applications to mute
- `SoundDevices.txt` - Audio device configuration

## Troubleshooting

### "Access Denied" Errors
- Ensure you're running as Administrator
- Check if antivirus software is blocking the operation
- Verify the target directory permissions

### "Source executable not found" Error
- Build the project first: `dotnet build --configuration Release`
- Check that the build output is in `bin\Release\net8.0-windows\win-x86\`

### Process Termination Issues
- Some processes may be protected by Windows
- Try closing the application manually before deployment
- Restart the deployment script

## Manual Deployment

If the scripts don't work, you can manually:

1. Stop any running GlobalHotkeys2.exe processes
2. Create `C:\Program Files\GlobalHotkeys` directory
3. Copy all files from `bin\Release\net8.0-windows\win-x86\` to the target directory
4. Copy `ProcessList.txt` and `SoundDevices.txt` from the project root

## Post-Deployment

After successful deployment:
- The application is available at `C:\Program Files\GlobalHotkeys\GlobalHotkeys2.exe`
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
