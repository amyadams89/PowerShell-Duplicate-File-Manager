# PowerShell Duplicate File Manager - Usage Examples

This document provides practical examples for using the duplicate file management scripts in different scenarios.

## Basic Usage Examples

### 1. Analyzing OneDrive for Duplicates (Simulation Mode)

This will scan your OneDrive folder and report potential duplicates without making any changes:

```powershell
.\scripts\OneDrive-Duplicate-Finder.ps1 -SimulateOnly
```

**Sample Output:**
```
[2025-03-03 10:15:42] [INFO] Scanning for duplicates in C:\Users\Username\OneDrive - Organization
[2025-03-03 10:15:42] [WARNING] This might take a while for large folders...
[2025-03-03 10:15:50] [INFO] Found 15428 total files to analyze.
[2025-03-03 10:15:52] [INFO] Found 3294 groups of files with the same size.
[2025-03-03 10:18:35] [SUCCESS] Found 7571 sets of duplicate files.

[2025-03-03 10:18:36] [INFO] Duplicate set 1/7571: presentation.pptx (Size: 4.25 MB)
[2025-03-03 10:18:36] [SUCCESS]   Keeping: C:\Users\Username\OneDrive - Organization\Presentations\Final\presentation.pptx (Last modified: 02/28/2025 13:22:36)
[2025-03-03 10:18:36] [WARNING]   Would remove: C:\Users\Username\OneDrive - Organization\Presentations\Draft\presentation (1).pptx (Last modified: 02/25/2025 09:15:22)

[2025-03-03 10:20:15] [INFO] Simulation summary:
[2025-03-03 10:20:15] [INFO] Duplicate sets found: 7571
[2025-03-03 10:20:15] [INFO] Files that would be kept: 7571
[2025-03-03 10:20:15] [INFO] Files that would be removed: 11209
[2025-03-03 10:20:15] [INFO] Potential space savings: 1513 MB
```

### 2. Cleaning Up External Drive Duplicates

To clean duplicate files on an external drive:

```powershell
.\scripts\OneDrive-Duplicate-Finder.ps1 -FolderPath "E:\Research" -VerboseOutput
```

The script will prompt you to:
1. Simulate removal (recommended first step)
2. Remove duplicates
3. Exit without changes

### 3. Scheduled Cleanup with No Interaction

For automation in scheduled tasks:

```powershell
.\scripts\OneDrive-Duplicate-Finder.ps1 -FolderPath "D:\Backups" -AutoConfirm -LogPath "C:\Logs\backup-cleanup.log"
```

This will run without user interaction and save the log to the specified location.

## Advanced Use Cases

### 1. Combining with Other PowerShell Commands

Pipe the results to get a report of the largest duplicate sets:

```powershell
$duplicateSets = .\scripts\OneDrive-Duplicate-Finder.ps1 -SimulateOnly -AutoConfirm
$duplicateSets | Sort-Object -Property SpaceSaved -Descending | Select-Object -First 10 | Format-Table
```

### 2. Processing Multiple Folders Sequentially

```powershell
$folders = @("D:\Photos", "D:\Videos", "D:\Documents")
$totalSaved = 0

foreach ($folder in $folders) {
    Write-Host "Processing $folder..." -ForegroundColor Cyan
    $result = .\scripts\OneDrive-Duplicate-Finder.ps1 -FolderPath $folder -AutoConfirm
    $totalSaved += $result.SpaceSaved
}

Write-Host "Total space saved across all folders: $totalSaved MB" -ForegroundColor Green
```

### 3. Custom Duplicate Handling

If you want to keep files in certain folders regardless of date:

```powershell
# Example of how to modify the script for priority folders
# Add this to the script or create a customized version

$priorityFolders = @(
    "C:\Users\Username\OneDrive\Important",
    "D:\Critical Data"
)

# Modify the sorting logic in Remove-Duplicates function:
# Instead of just sorting by date, prioritize certain locations
$keepFile = $set | Sort-Object -Property {
    # Check if file is in a priority folder
    foreach ($pFolder in $priorityFolders) {
        if ($_.FullName -like "$pFolder*") {
            return [DateTime]::MaxValue  # Ensure priority files are kept
        }
    }
    # Otherwise sort by last write time
    return $_.LastWriteTime
} -Descending | Select-Object -First 1
```

## Visual Examples

Here's what the duplicate identification process looks like. The script identifies these variations as duplicates:

1. **Same name, different locations:**
   ```
   Documents\Report.docx
   Documents\Archive\Report.docx
   ```

2. **Numbered duplicates:**
   ```
   Photo.jpg
   Photo (1).jpg
   Photo(2).jpg
   ```

3. **Copy suffix variations:**
   ```
   Presentation.pptx
   Presentation - Copy.pptx
   Presentation_copy.pptx
   ```

4. **Numeric suffix variations:**
   ```
   Data.xlsx
   Data_1.xlsx
   ```

## Common Questions

**Q: Will this delete the original files?**
A: No, the script keeps one copy of each file (by default, the newest version).

**Q: How does it determine which file to keep?**
A: By default, it keeps the file with the most recent modification date.

**Q: Is it safe to use on my main folders?**
A: Always run in simulation mode first (-SimulateOnly) to review what would be removed.
