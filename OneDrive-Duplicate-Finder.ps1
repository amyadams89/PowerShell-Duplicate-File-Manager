# PowerShell script to find and remove actual duplicate files
# -----------------------------------------------------------
# Author: Your Name
# Version: 1.1
# Last Updated: March 3, 2025
# GitHub: https://github.com/yourusername/PowerShell-Duplicate-File-Manager
# 
# Description:
#   This script identifies and removes duplicate files in the specified folder
#   while preserving folder structure. It uses file size and name pattern matching
#   to identify duplicates and keeps the newest file by default.
#
# Parameters:
#   -FolderPath     : The root folder to scan for duplicates (default: current user's OneDrive)
#   -SimulateOnly   : Run in simulation mode without removing files
#   -AutoConfirm    : Skip confirmation prompts (use with caution)
#   -LogPath        : Path to save the log file (default: same directory as script)
#   -VerboseOutput  : Show detailed progress information
# 
# Example usage:
#   .\OneDrive-Duplicate-Finder.ps1 -FolderPath "D:\MyFiles" -SimulateOnly
#   .\OneDrive-Duplicate-Finder.ps1 -FolderPath "$env:USERPROFILE\OneDrive" -LogPath "C:\Logs\duplicate-scan.log"
# -----------------------------------------------------------

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$FolderPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$SimulateOnly = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$AutoConfirm = $false,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput = $false
)

# -----------------------------------------------------------
# Script initialization
# -----------------------------------------------------------

# Set up error handling
$ErrorActionPreference = "Stop"

# Initialize log file
if ([string]::IsNullOrEmpty($LogPath)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $scriptDir = Split-Path -Parent $scriptPath
    $LogPath = Join-Path -Path $scriptDir -ChildPath "DuplicateScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

# If no folder path is provided, try to detect OneDrive path
if ([string]::IsNullOrEmpty($FolderPath)) {
    # Try to automatically detect OneDrive folder
    $possibleOneDrivePaths = @(
        "$env:USERPROFILE\OneDrive",
        "$env:USERPROFILE\OneDrive - *"
    )
    
    foreach ($path in $possibleOneDrivePaths) {
        $matchingPaths = Resolve-Path -Path $path -ErrorAction SilentlyContinue
        if ($matchingPaths) {
            $FolderPath = $matchingPaths[0].Path
            break
        }
    }
    
    if ([string]::IsNullOrEmpty($FolderPath)) {
        Write-Host "Could not automatically detect OneDrive folder. Please specify the folder path." -ForegroundColor Red
        exit 1
    }
}

# Check if folder path exists
if (-not (Test-Path -Path $FolderPath)) {
    Write-Host "Folder path '$FolderPath' does not exist. Please check the path and try again." -ForegroundColor Red
    exit 1
}

# -----------------------------------------------------------
# Logging functions
# -----------------------------------------------------------

function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Log to file
    Add-Content -Path $LogPath -Value $logMessage
    
    # Log to console with color coding if not suppressed
    if (-not $NoConsole) {
        $color = switch ($Level) {
            "INFO"    { "White" }
            "WARNING" { "Yellow" }
            "ERROR"   { "Red" }
            "SUCCESS" { "Green" }
            default   { "White" }
        }
        
        Write-Host $logMessage -ForegroundColor $color
    }
}

# -----------------------------------------------------------
# Function to find true duplicate files
# -----------------------------------------------------------

function Find-TrueDuplicates {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    
    Write-Log "Scanning for duplicates in $FolderPath" -Level "INFO"
    Write-Log "This might take a while for large folders..." -Level "WARNING"
    
    # Get all files recursively
    Write-Log "Getting file list..." -Level "INFO" -NoConsole:(-not $VerboseOutput)
    $allFiles = Get-ChildItem -Path $FolderPath -File -Recurse -ErrorAction SilentlyContinue
    
    Write-Log "Found $($allFiles.Count) total files to analyze." -Level "INFO"
    
    # Group files by size (first level filter)
    Write-Log "Grouping files by size..." -Level "INFO" -NoConsole:(-not $VerboseOutput)
    $sizeGroups = $allFiles | Group-Object -Property Length | Where-Object { $_.Count -gt 1 }
    
    Write-Log "Found $($sizeGroups.Count) groups of files with the same size." -Level "INFO"
    
    $duplicateSets = @()
    $totalGroups = $sizeGroups.Count
    $groupCounter = 0
    
    foreach ($sizeGroup in $sizeGroups) {
        $groupCounter++
        $filesInGroup = $sizeGroup.Group
        
        # Show progress every 20 groups
        if ($groupCounter % 20 -eq 0) {
            $percentComplete = [math]::Round(($groupCounter / $totalGroups) * 100, 1)
            Write-Progress -Activity "Analyzing file groups" -Status "Group $groupCounter of $totalGroups ($percentComplete%)" -PercentComplete $percentComplete
            
            if ($VerboseOutput) {
                Write-Log "Progress: Analyzing group $groupCounter of $totalGroups ($percentComplete%)" -Level "INFO" -NoConsole:(-not $VerboseOutput)
            }
        }
        
        # Files with exactly the same name are definite duplicates
        $nameGroups = $filesInGroup | Group-Object -Property Name
        
        foreach ($nameGroup in $nameGroups) {
            if ($nameGroup.Count -gt 1) {
                # These are exact duplicates - same size, same filename
                $duplicateSets += ,$nameGroup.Group
            }
        }
        
        # Now check for similar filenames within the same size group
        # Group files to compare efficiently
        $filesToCheck = @($filesInGroup)
        $alreadyProcessed = @{}
        
        for ($i = 0; $i -lt $filesToCheck.Count; $i++) {
            $file = $filesToCheck[$i]
            
            # Skip if this file was already processed as part of a group
            if ($alreadyProcessed.ContainsKey($file.FullName)) { continue }
            
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $ext = [System.IO.Path]::GetExtension($file.Name)
            
            # Get base name without common duplicate suffixes
            $cleanBaseName = $baseName
            # Remove " (1)", " (2)", etc.
            $cleanBaseName = $cleanBaseName -replace " \(\d+\)$", ""
            # Remove "(1)", "(2)", etc.
            $cleanBaseName = $cleanBaseName -replace "\(\d+\)$", ""
            # Remove " - Copy", " - Copy (1)" etc.
            $cleanBaseName = $cleanBaseName -replace " - Copy(\(\d+\))?$", ""
            # Remove "_1", "_2", etc.
            $cleanBaseName = $cleanBaseName -replace "_\d+$", ""
            # Remove "_copy", "_copy1" etc.
            $cleanBaseName = $cleanBaseName -replace "_copy\d*$", ""
            
            # Find similar files
            $similarFiles = @($file)
            
            for ($j = 0; $j -lt $filesToCheck.Count; $j++) {
                if ($i -eq $j) { continue } # Skip self-comparison
                
                $compareFile = $filesToCheck[$j]
                $compareBaseName = [System.IO.Path]::GetFileNameWithoutExtension($compareFile.Name)
                $compareExt = [System.IO.Path]::GetExtension($compareFile.Name)
                
                # Skip if extension doesn't match
                if ($compareExt -ne $ext) { continue }
                
                # Skip if already processed
                if ($alreadyProcessed.ContainsKey($compareFile.FullName)) { continue }
                
                # Check for duplicate patterns
                $isMatch = $false
                
                # Exact match without suffix
                if ($compareBaseName -eq $cleanBaseName) {
                    $isMatch = $true
                }
                # Pattern: original vs "original (1)"
                elseif ($compareBaseName -match "^$([regex]::Escape($cleanBaseName)) \(\d+\)$") {
                    $isMatch = $true
                }
                # Pattern: original vs "original(1)"
                elseif ($compareBaseName -match "^$([regex]::Escape($cleanBaseName))\(\d+\)$") {
                    $isMatch = $true
                }
                # Pattern: original vs "original - Copy"
                elseif ($compareBaseName -match "^$([regex]::Escape($cleanBaseName)) - Copy(\(\d+\))?$") {
                    $isMatch = $true
                }
                # Pattern: original vs "original_1"
                elseif ($compareBaseName -match "^$([regex]::Escape($cleanBaseName))_\d+$") {
                    $isMatch = $true
                }
                # Pattern: original vs "original_copy"
                elseif ($compareBaseName -match "^$([regex]::Escape($cleanBaseName))_copy\d*$") {
                    $isMatch = $true
                }
                
                if ($isMatch) {
                    $similarFiles += $compareFile
                    $alreadyProcessed[$compareFile.FullName] = $true
                }
            }
            
            # If we found similar files, add them as a duplicate set
            if ($similarFiles.Count -gt 1) {
                $duplicateSets += ,$similarFiles
                
                # Mark all as processed
                foreach ($sf in $similarFiles) {
                    $alreadyProcessed[$sf.FullName] = $true
                }
            }
        }
    }
    
    Write-Progress -Activity "Analyzing file groups" -Completed
    
    Write-Log "Found $($duplicateSets.Count) sets of duplicate files." -Level "SUCCESS"
    
    return $duplicateSets
}

# -----------------------------------------------------------
# Function to remove duplicates
# -----------------------------------------------------------

function Remove-Duplicates {
    param (
        [Parameter(Mandatory=$true)]
        $DuplicateSets,
        
        [Parameter(Mandatory=$false)]
        [switch]$WhatIf
    )
    
    $totalRemoved = 0
    $remainingCount = 0
    $setCounter = 0
    $totalBytes = 0
    $errorCount = 0
    
    foreach ($set in $DuplicateSets) {
        $setCounter++
        
        # Keep the newest file
        $keepFile = $set | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
        $remainingCount++
        
        # Find files to remove
        $filesToRemove = $set | Where-Object { $_.FullName -ne $keepFile.FullName }
        $totalRemoved += $filesToRemove.Count
        
        # Calculate space saved
        $fileSize = $keepFile.Length
        $spaceSaved = $fileSize * $filesToRemove.Count
        $totalBytes += $spaceSaved
        
        Write-Log "`nDuplicate set $setCounter/$($DuplicateSets.Count): $($keepFile.Name) (Size: $(($fileSize / 1MB).ToString('0.00')) MB)" -Level "INFO"
        Write-Log "  Keeping: $($keepFile.FullName) (Last modified: $($keepFile.LastWriteTime))" -Level "SUCCESS"
        
        foreach ($file in $filesToRemove) {
            if ($WhatIf) {
                Write-Log "  Would remove: $($file.FullName) (Last modified: $($file.LastWriteTime))" -Level "WARNING"
            } else {
                try {
                    Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                    Write-Log "  Removed: $($file.FullName) (Last modified: $($file.LastWriteTime))" -Level "INFO"
                }
                catch {
                    $errorCount++
                    Write-Log "  Error removing $($file.FullName): $_" -Level "ERROR"
                }
            }
        }
    }
    
    $savedSpace = [math]::Round($totalBytes / 1MB, 2)
    
    if ($WhatIf) {
        Write-Log "`nSimulation summary:" -Level "INFO"
        Write-Log "Duplicate sets found: $($DuplicateSets.Count)" -Level "INFO"
        Write-Log "Files that would be kept: $remainingCount" -Level "INFO"
        Write-Log "Files that would be removed: $totalRemoved" -Level "INFO"
        Write-Log "Potential space savings: $savedSpace MB" -Level "INFO"
    } else {
        Write-Log "`nOperation summary:" -Level "INFO"
        Write-Log "Duplicate sets processed: $($DuplicateSets.Count)" -Level "INFO"
        Write-Log "Files kept: $remainingCount" -Level "INFO"
        Write-Log "Files removed: $($totalRemoved - $errorCount)" -Level "SUCCESS"
        Write-Log "Files with removal errors: $errorCount" -Level "ERROR"
        Write-Log "Space freed: $savedSpace MB" -Level "SUCCESS"
    }
    
    # Return results for potential future use
    return @{
        DuplicateSets = $DuplicateSets.Count
        KeptFiles = $remainingCount
        RemovedFiles = $totalRemoved - $errorCount
        ErrorCount = $errorCount
        SpaceSaved = $savedSpace
    }
}

# -----------------------------------------------------------
# Main script
# -----------------------------------------------------------

try {
    $startTime = Get-Date
    
    Write-Log "============================================================" -Level "INFO"
    Write-Log "Starting duplicate file scan on $(Get-Date)" -Level "INFO"
    Write-Log "Target directory: $FolderPath" -Level "INFO"
    Write-Log "Simulation mode: $SimulateOnly" -Level "INFO"
    Write-Log "Log file: $LogPath" -Level "INFO"
    Write-Log "============================================================" -Level "INFO"
    
    # Find duplicates
    $duplicateSets = Find-TrueDuplicates -FolderPath $FolderPath
    
    if ($duplicateSets.Count -eq 0) {
        Write-Log "No duplicate files found." -Level "SUCCESS"
        exit
    }
    
    # Process duplicates based on parameters
    if ($SimulateOnly) {
        # If SimulateOnly is specified, just run in WhatIf mode
        $results = Remove-Duplicates -DuplicateSets $duplicateSets -WhatIf
    } elseif ($AutoConfirm) {
        # If AutoConfirm is specified, remove duplicates without confirmation
        $results = Remove-Duplicates -DuplicateSets $duplicateSets
    } else {
        # Show menu
        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "1. Simulate removal (show what would be removed)" -ForegroundColor White
        Write-Host "2. Remove duplicates (keeping newest file in each set)" -ForegroundColor White
        Write-Host "3. Exit without changes" -ForegroundColor White
        
        $option = Read-Host -Prompt "Select an option (1-3)"
        
        switch ($option) {
            "1" {
                $results = Remove-Duplicates -DuplicateSets $duplicateSets -WhatIf
            }
            "2" {
                $confirmation = Read-Host -Prompt "This will permanently remove duplicate files. Continue? (Y/N)"
                if ($confirmation -eq "Y" -or $confirmation -eq "y") {
                    $results = Remove-Duplicates -DuplicateSets $duplicateSets
                } else {
                    Write-Log "Operation cancelled by user." -Level "WARNING"
                }
            }
            "3" {
                Write-Log "Operation cancelled by user." -Level "WARNING"
            }
            default {
                Write-Log "Invalid option. Operation cancelled." -Level "ERROR"
            }
        }
    }
    
    # Calculate duration
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Log "Script completed in $($duration.TotalMinutes.ToString('0.00')) minutes." -Level "INFO"
    Write-Log "============================================================" -Level "INFO"
}
catch {
    Write-Log "An error occurred: $_" -Level "ERROR"
    Write-Log $_.ScriptStackTrace -Level "ERROR"
}

if (-not $AutoConfirm) {
    Write-Host "`nScript complete. Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
