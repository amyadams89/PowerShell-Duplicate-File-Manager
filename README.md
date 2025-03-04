# PowerShell Duplicate File Manager

A collection of PowerShell scripts for finding and managing duplicate files across drives and cloud storage.

## Features

- Identifies duplicate files using multiple detection methods
- Preserves folder structure while removing duplicates
- Provides simulation mode to preview changes before execution
- Prioritizes keeping newest versions of files
- Handles common duplicate naming patterns

## Scripts

* **OneDrive-Duplicate-Finder.ps1**: Identifies and removes duplicates in OneDrive folders
* **E-Drive-Duplicate-Remover.ps1**: Removes duplicates from organized folders on external drives
* **Basic-Duplicate-Organizer.ps1**: Organizes files by size to help identify duplicates

## Usage

### OneDrive Duplicate Finder

```powershell
# Edit the script to set your OneDrive path
$baseDir = "C:\Users\YourUsername\OneDrive - YourOrganization"

# Run the script
.\scripts\OneDrive-Duplicate-Finder.ps1
