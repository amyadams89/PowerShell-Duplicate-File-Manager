Below is a **comprehensive plan** and **annotated PowerShell example** for reorganizing both Outlook folders (via the Outlook COM object) and local OneDrive directories. The solution includes:

1. A **step-by-step approach** to identifying empty or duplicate folders, merging or archiving them, and handling exclusions.  
2. A **“dry run” mode** that reports intended changes without modifying anything.  
3. A **“commit” mode** that carries out the folder merges, renames, or deletions.  
4. Guidance for **naming conventions**, **error handling**, and **best practices** to help preserve data integrity.

---

## 1. Recommended Approach

### 1.1 Identify Requirements and Plan Your Cleanup

1. **Take Inventory**  
   - List all your Outlook folders (including subfolders) and note which ones are essential, which ones are duplicates, and which ones are outdated or empty.  
   - Do the same on your **OneDrive** path structure—determine which folders are empty, near-empty, duplicated, or have ambiguous names.

2. **Define Exclusions**  
   - Maintain a central list (Array, CSV, or text file) of folder names or paths you want to exclude. For Outlook, this might include shared mailboxes, calendars, or organizational folders you do not own. For OneDrive, it might be system-generated folders or placeholder directories you must keep.

3. **Decide on a “Merge” vs. “Archive” vs. “Delete” Strategy**  
   - For each folder candidate (in Outlook or OneDrive), decide whether it should be:  
     1. **Merged** (contents moved into a primary or “master” folder).  
     2. **Renamed** (if the folder name is outdated or unclear).  
     3. **Archived** (moved to an “Archive” folder for safekeeping).  
     4. **Deleted** (only if confirmed empty or truly not needed).

4. **Safety Check**  
   - Prepare for a **Dry Run**. You want to see a summary of proposed actions before actually altering your folder structure.  
   - Confirm you have **recent backups** (PST exports for Outlook, and a local backup or snapshot of the OneDrive directory).

### 1.2 Implement the Cleanup in PowerShell

1. **Outlook COM Automation**  
   - Use the `Outlook.Application` COM object.  
   - Traverse folders recursively starting from your primary mailbox or a given root folder.  
   - Compare each folder name against your **exclusion list**. Skip anything excluded.  
   - Count items using `folder.Items.Count` to determine if it’s empty (or near-empty).  
   - Detect duplicates (exact name matches, or name plus a certain threshold of identical subfolders/items) if needed.  
   - In **dry run mode**, output an action plan (rename, merge into Archive, etc.).  
   - In **commit mode**, set `folder.Name = 'NewName'`, move messages if merging, or `folder.Delete()` to remove unwanted folders.  
   - **Handle errors** in case certain folders cannot be deleted or renamed (e.g., default Outlook folders).

2. **Local OneDrive File System**  
   - Recursively search directories in your OneDrive path (e.g., `C:\Users\YourUser\OneDrive`).  
   - Exclude certain directories from processing (using your same or a separate exclusion list).  
   - Check if a folder is empty or if it’s a near-duplicate (the script can detect duplicates by checking if subfolder/file structures are identical, or simply if names match).  
   - In **dry run mode**, print out planned operations: rename, move/merge, or delete.  
   - In **commit mode**, perform these filesystem operations with standard `Move-Item`, `Rename-Item`, or `Remove-Item` cmdlets.  
   - **Log** any errors so you can address permission or locking issues.

### 1.3 Naming Scheme Suggestions

- For **projects or date-oriented** folders, consider a format like:  
  ```
  YYYY_MM_Description
  ```
  or
  ```
  YYYY_Semester_ProjectName
  ```  
  Examples:  
  - `2025_Spring_MarketingPlan`  
  - `2024_10_NewWebsiteAssets`

- For **personal or student-based** folders, you could do:  
  ```
  Year_Semester_StudentName
  ```  
  Example: `2025_Fall_JohnDoe_Thesis`

- Keep the structure **consistent** across Outlook and OneDrive—this improves discoverability and prevents confusion.

### 1.4 Best Practices and Disclaimers

1. **Always Backup First**  
   - Export critical Outlook folders to a PST or use the built-in Outlook Export wizard for insurance.  
   - Backup critical data from OneDrive before running large-scale deletions or merges.

2. **Test on a Small Subset**  
   - Before performing a mass reorganization, run the script on a single test mailbox folder or a small sample directory structure in OneDrive.

3. **Incremental Cleanup**  
   - Instead of reorganizing everything at once, consider doing it in batches (e.g., handle outdated folders first, then rename or unify duplicates next).

4. **Irreversible Deletions**  
   - Deletions in Outlook may or may not be recoverable depending on your corporate or personal retention policy.  
   - Deletions on OneDrive often go to the OneDrive Recycle Bin, but do not rely on that as your only backup plan.

5. **References**  
   - **Outlook COM Object** documentation:  
     [Microsoft Docs: Outlook VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook)  
   - **FileSystem** operations in PowerShell:  
     [Microsoft Docs: About FileSystem Provider](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_filesystem_provider)

---

## 2. Detailed PowerShell Script Example

> **Note**: This script is illustrative. You may need to adapt folder paths, mailbox names, or rename logic to match your environment and needs. Also ensure that **Outlook is installed** and that your Windows PowerShell or PowerShell 7 environment can load COM objects.

```powershell
<#
.SYNOPSIS
    Reorganize Outlook folders and local OneDrive directories
    in either 'dry run' mode or 'commit' mode.

.DESCRIPTION
    - Connects to Outlook via COM automation.
    - Recursively enumerates folders in Outlook and local OneDrive paths.
    - Identifies empty or near-empty folders, duplicates, or outdated names.
    - Logs intended actions in Dry Run mode.
    - Optionally executes folder merges/renames/deletions in Commit mode.

.PARAMETER DryRun
    If specified, the script will only show what actions would be taken.

.PARAMETER Commit
    If specified, the script will actually perform the reorganization.

.EXAMPLE
    .\Reorganize-Folders.ps1 -DryRun

    Shows a report of proposed changes, but makes no actual modifications.

.EXAMPLE
    .\Reorganize-Folders.ps1 -Commit

    Performs the actual folder renaming, merging, or deletion operations.
#>

param(
    [switch]$DryRun,
    [switch]$Commit
)

# --- USER CONFIGURATIONS ---

# 1) Set your OneDrive root path. Adjust as needed.
$OneDriveRoot = "C:\Users\<YourUsername>\OneDrive"

# 2) Outlook root folder name (this could be your primary mailbox name).
#    For example: "Mailbox - Jane Doe" or "YourName@company.com"
$OutlookRootFolderName = "Mailbox - Jane Doe"

# 3) Exclusion list for Outlook folder names (exact match).
#    Add more as needed.
$OutlookExclusions = @(
    "Calendar",
    "Contacts",
    "Shared Folders",
    "Some Public Folder",
    "RSS Feeds"
)

# 4) Exclusion list for OneDrive directories (exact or partial match).
$OneDriveExclusions = @(
    "Microsoft Teams Chat Files",
    "System-Generated",
    "DoNotDelete"
)

# 5) Archive folder name in Outlook (will create if it doesn’t exist).
$OutlookArchiveFolderName = "Old Folders"

# 6) Archive folder name in OneDrive.
$OneDriveArchiveFolderName = "Archive"

# --- LOGGING FUNCTION ---
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$timestamp] [$Level] $Message"
}

# --- OUTLOOK FOLDER REORGANIZATION FUNCTIONS ---

function Get-OutlookNamespace {
    try {
        Write-Log "Creating Outlook Application COM object..."
        $outlookApp = New-Object -ComObject Outlook.Application
        $namespace = $outlookApp.GetNameSpace("MAPI")
        return $namespace
    }
    catch {
        Write-Log "Failed to create Outlook COM object. $_" "ERROR"
        return $null
    }
}

function Get-OutlookFolderByName {
    param(
        [Microsoft.Office.Interop.Outlook._NameSpace]$Namespace,
        [string]$FolderName
    )
    foreach ($folder in $Namespace.Folders) {
        if ($folder.Name -eq $FolderName) {
            return $folder
        }
    }
    return $null
}

function Ensure-OutlookArchiveFolder {
    param(
        [Microsoft.Office.Interop.Outlook.MAPIFolder]$RootFolder,
        [string]$ArchiveFolderName
    )

    # Check if archive subfolder already exists:
    foreach ($sub in $RootFolder.Folders) {
        if ($sub.Name -eq $ArchiveFolderName) {
            return $sub
        }
    }

    # If not exist, create a new folder under the root:
    Write-Log "Creating archive folder '$ArchiveFolderName' under '$($RootFolder.Name)'"
    if ($Commit) {
        return $RootFolder.Folders.Add($ArchiveFolderName)
    } else {
        return $null  # in DryRun, we do not actually create it
    }
}

function Process-OutlookFolders {
    param(
        [Microsoft.Office.Interop.Outlook.MAPIFolder]$RootFolder,
        [string[]]$ExclusionList,
        [Microsoft.Office.Interop.Outlook.MAPIFolder]$ArchiveFolder
    )
    # Recursive function to traverse Outlook folders

    foreach ($folder in $RootFolder.Folders) {
        # Skip if in exclusion list
        if ($ExclusionList -contains $folder.Name) {
            Write-Log "Skipping Outlook folder '$($folder.Name)' (excluded)"
            continue
        }

        # Count items
        $itemCount = $folder.Items.Count
        $subfolderCount = $folder.Folders.Count

        # Example logic: If folder is empty and has no subfolders -> move or delete
        if ($itemCount -eq 0 -and $subfolderCount -eq 0) {
            Write-Log "Empty Outlook folder found: '$($folder.Name)'. Will move to archive or delete."

            if ($ArchiveFolder -and $folder.Name -ne $OutlookArchiveFolderName) {
                Write-Log "Would move '$($folder.Name)' to archive folder '$($ArchiveFolder.Name)'"
                if ($Commit) {
                    try {
                        $folder.Move($ArchiveFolder) | Out-Null
                        Write-Log "Moved '$($folder.Name)' to archive successfully."
                    }
                    catch {
                        Write-Log "Failed to move '$($folder.Name)' to archive. $_" "ERROR"
                    }
                }
            }
            else {
                # Or optionally delete if you prefer actual deletion of empty folders
                Write-Log "Would delete empty folder '$($folder.Name)'."
                if ($Commit) {
                    try {
                        $folder.Delete()
                        Write-Log "Deleted folder '$($folder.Name)'."
                    }
                    catch {
                        Write-Log "Failed to delete '$($folder.Name)'. $_" "ERROR"
                    }
                }
            }
        }
        else {
            # Potential rename logic or detection of duplicates
            # For demonstration, let's say we rename any folder with "Old_" prefix to "Archive_"
            if ($folder.Name -like "Old_*") {
                $newName = $folder.Name.Replace("Old_", "Archive_")
                Write-Log "Would rename folder '$($folder.Name)' to '$newName'."
                if ($Commit) {
                    try {
                        $folder.Name = $newName
                        Write-Log "Renamed folder to '$newName'."
                    }
                    catch {
                        Write-Log "Failed to rename '$($folder.Name)'. $_" "ERROR"
                    }
                }
            }

            # Recurse subfolders
            Process-OutlookFolders -RootFolder $folder -ExclusionList $ExclusionList -ArchiveFolder $ArchiveFolder
        }
    }
}

# --- ONEDRIVE FOLDER REORGANIZATION FUNCTIONS ---

function Process-OneDriveFolders {
    param(
        [string]$Path,
        [string[]]$Exclusions
    )

    # Get subdirectories
    $subDirs = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | Sort-Object Name

    foreach ($dir in $subDirs) {
        # If in exclusions, skip
        if ($Exclusions -contains $dir.Name) {
            Write-Log "Skipping OneDrive folder '$($dir.FullName)' (excluded)"
            continue
        }

        # Check if the directory is empty
        $contents = Get-ChildItem -Path $dir.FullName -Force -ErrorAction SilentlyContinue
        if (!$contents) {
            Write-Log "Empty OneDrive folder found: '$($dir.FullName)'. Will move or delete."
            # Example: move to Archive folder if not the Archive folder itself
            if ($dir.Name -ne $OneDriveArchiveFolderName) {
                $archivePath = Join-Path -Path $Path -ChildPath $OneDriveArchiveFolderName
                Write-Log "Would move '$($dir.FullName)' to '$archivePath'."

                if ($Commit) {
                    try {
                        if (!(Test-Path -Path $archivePath)) {
                            Write-Log "Creating OneDrive archive folder: $archivePath"
                            New-Item -Path $archivePath -ItemType Directory | Out-Null
                        }
                        # Move folder
                        $targetPath = Join-Path -Path $archivePath -ChildPath $dir.Name
                        Move-Item -Path $dir.FullName -Destination $targetPath
                        Write-Log "Moved folder '$($dir.FullName)' to '$targetPath'."
                    }
                    catch {
                        Write-Log "Failed to move folder '$($dir.FullName)'. $_" "ERROR"
                    }
                }
            }
            else {
                # If it's the actual Archive folder and is empty, we might just keep it
                Write-Log "Archive folder is empty but will keep it."
            }
        }
        else {
            # Potential rename or duplication checks:
            # Example: rename "Old_Projects" -> "Archive_Projects"
            if ($dir.Name -like "Old_*") {
                $newName = $dir.Name.Replace("Old_", "Archive_")
                $newFullPath = Join-Path -Path $dir.Parent.FullName -ChildPath $newName
                Write-Log "Would rename OneDrive folder '$($dir.FullName)' to '$newFullPath'."
                if ($Commit) {
                    try {
                        Rename-Item -Path $dir.FullName -NewName $newName
                        Write-Log "Renamed folder '$($dir.FullName)' to '$newFullPath'."
                    }
                    catch {
                        Write-Log "Failed to rename folder '$($dir.FullName)'. $_" "ERROR"
                    }
                }
            }

            # Recurse into subfolders
            Process-OneDriveFolders -Path $dir.FullName -Exclusions $Exclusions
        }
    }
}

# --- MAIN SCRIPT LOGIC ---

Write-Log "Starting folder reorganization script..."

if (-not $DryRun -and -not $Commit) {
    Write-Log "No mode specified. Please use -DryRun or -Commit." "WARNING"
    return
}

# 1) Connect to Outlook (if available) and process Outlook folders
$namespace = Get-OutlookNamespace
if ($null -ne $namespace) {
    $rootOutlookFolder = Get-OutlookFolderByName -Namespace $namespace -FolderName $OutlookRootFolderName
    if ($rootOutlookFolder) {
        # Ensure we have an Archive folder (in commit mode, it will be created if missing)
        $archiveFolder = Ensure-OutlookArchiveFolder -RootFolder $rootOutlookFolder -ArchiveFolderName $OutlookArchiveFolderName
        Write-Log "Processing Outlook folders under '$($rootOutlookFolder.Name)'..."
        Process-OutlookFolders -RootFolder $rootOutlookFolder -ExclusionList $OutlookExclusions -ArchiveFolder $archiveFolder
    }
    else {
        Write-Log "Could not find Outlook root folder named '$OutlookRootFolderName'." "ERROR"
    }
}
else {
    Write-Log "Skipping Outlook processing because Outlook COM object was not created."
}

# 2) Process OneDrive folders
if (Test-Path $OneDriveRoot) {
    Write-Log "Processing OneDrive folders at '$OneDriveRoot'..."
    Process-OneDriveFolders -Path $OneDriveRoot -Exclusions $OneDriveExclusions
}
else {
    Write-Log "OneDrive root path '$OneDriveRoot' does not exist. Skipping."
}

Write-Log "Folder reorganization script completed."
```

### How to Use This Script

1. **Save** the code as `Reorganize-Folders.ps1` (or a name of your choice).  
2. **Adjust the configuration** near the top:
   - `$OneDriveRoot` to your actual OneDrive path.  
   - `$OutlookRootFolderName` to match your mailbox or main Outlook folder name.  
   - `$OutlookExclusions` and `$OneDriveExclusions` arrays to capture all the folders you want to skip.  
3. **Open PowerShell** (ensure you have script execution privileges, e.g., `Set-ExecutionPolicy RemoteSigned` if needed).  
4. **Run a Dry Run**:  
   ```powershell
   .\Reorganize-Folders.ps1 -DryRun
   ```  
   - This should generate a log of what the script **would** do, but make no changes.  
5. **Check the Log** carefully to confirm the intended actions look correct.  
6. **Run in Commit Mode**:  
   ```powershell
   .\Reorganize-Folders.ps1 -Commit
   ```  
   - This will proceed to rename, move, archive, or delete Outlook folders and OneDrive directories as configured.  
7. **Verify** the results in Outlook (under your mailbox structure) and in your OneDrive directory.

---

## 3. Practical Tips & Best Practices

1. **Back Up**  
   - Export key Outlook folders to PST if you have critical email data.  
   - For OneDrive, make a copy (e.g., a simple ZIP or external drive backup) before large-scale changes.

2. **Use Incremental Steps**  
   - Clean up only your known empty or old folders first, confirm success, then handle potential duplicates or merges later.  
   - For merging Outlook folders, you may need additional logic to **move items** from one folder to another, if you truly want to consolidate mail content. The script above demonstrates how to move the entire folder, but item-by-item merges can also be done if required.

3. **Handle Default Outlook Folders Cautiously**  
   - Some folders (like Inbox, Drafts, Sent Items) cannot be deleted or renamed by default. The script should gracefully handle exceptions.

4. **Logging & Recovery**  
   - Maintain logs of all actions. If something goes wrong, you can identify which folder was renamed or removed at each step.  
   - Check your OneDrive Recycle Bin or Outlook “Deleted Items” (if such a route is used by your environment) for quick recoveries.

5. **Tailor Duplicate Detection**  
   - This example shows a very basic rename approach for “Old_*” folders. In production, you might want to compare folder item counts or last modified dates to identify duplicates vs. truly separate folders.

---

## 4. Disclaimers & References

1. **Data Loss Warning**  
   - Any script that automates folder deletion can result in **irreversible data loss** if used incorrectly. Always run a **dry run**, verify logs, and maintain backups.

2. **Outlook COM Limitations**  
   - The [Outlook COM API](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook) can behave differently based on your system’s bitness (32-bit vs. 64-bit) and Outlook security policies. If you encounter errors, ensure you’re running PowerShell in the same bitness as Outlook.

3. **Adapt for Your Environment**  
   - Folder names, mailbox references, and some methods (like `folder.Delete()`) may require special permissions or might be restricted by corporate Group Policy or retention policies.

4. **No Warranty**  
   - This is example code. Test thoroughly in your environment. The authors (and ChatGPT) assume **no responsibility** for any accidental data deletion or environment disruption.

---

# Final Summary

1. **Plan** which folders to keep, rename, archive, merge, or delete.  
2. **Maintain** an exclusion list for both Outlook and OneDrive directories.  
3. **Run** in **Dry Run** mode first; review the log to ensure correctness.  
4. **Execute** in **Commit** mode only after verifying backups.  
5. **Adopt** a consistent naming scheme like `YYYY_Semester_ProjectName` to keep your new folder structure intuitive and future-proof.  

With these steps, your email and local directories should become far more organized, with minimal risk to important data. Good luck with your cleanup!
