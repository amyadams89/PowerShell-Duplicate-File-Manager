# PowerShell Duplicate File Manager

A collection of PowerShell scripts designed to identify and manage duplicate files across local drives and cloud storage like OneDrive. These scripts help reclaim storage space while preserving folder structure and ensuring data integrity.

## 🔍 Features

- **Smart Duplicate Detection**: Identifies duplicates using file size and name pattern recognition
- **Structure Preservation**: Maintains your folder organization while removing redundant files
- **Safety First**: Simulation mode to preview changes before making them
- **Intelligent Selection**: Keeps the newest version of each file by default
- **Multiple Storage Support**: Works with local drives, flash drives, and OneDrive folders

## 📁 Scripts

The repository contains several scripts for different duplicate management scenarios:

### OneDrive-Duplicate-Finder.ps1

Identifies and removes duplicate files within OneDrive folders while preserving folder structure. Particularly useful for cloud storage where duplicate files waste both space and bandwidth.

### E-Drive-Duplicate-Remover.ps1

Removes duplicates from external drives or flash drives, designed to work with previously organized folders.

### Basic-Duplicate-Organizer.ps1

Organizes files by size to help identify potential duplicates manually or prepare for automated cleaning.

## 🚀 Usage

### OneDrive Duplicate Finder

```powershell
# Clone the repository
git clone https://github.com/yourusername/PowerShell-Duplicate-File-Manager.git
cd PowerShell-Duplicate-File-Manager

# Edit the script to set your OneDrive path or use parameters
$baseDir = "C:\Users\YourUsername\OneDrive - YourOrganization"

# Run the script
.\scripts\OneDrive-Duplicate-Finder.ps1
```

When running the script, you'll be presented with three options:
1. **Simulate removal**: Shows what would be removed without making changes (recommended first step)
2. **Remove duplicates**: Permanently removes duplicate files, keeping the newest file in each set
3. **Exit without changes**: Closes the script without making any changes

### Sample Output

```
Scanning for duplicates in C:\Users\Username\OneDrive - Organization
This might take a while for large folders...
Found 15,428 total files to analyze.
Found 3,294 groups of files with the same size.
Found 7,571 sets of duplicate files.

Duplicate set 1542/7571: Research-Paper-Draft.docx (Size: 2.34 MB)
  Keeping: C:\Users\Username\OneDrive - Organization\Documents\Research\Final\Research-Paper-Draft.docx (Last modified: 11/15/2024 14:22:36)
  Would remove: C:\Users\Username\OneDrive - Organization\Documents\Research\Old\Research-Paper-Draft (1).docx (Last modified: 11/10/2024 09:15:22)

Simulation summary:
Duplicate sets found: 7,571
Files that would be kept: 7,571
Files that would be removed: 11,209
Potential space savings: 1,513 MB
```

## ⚙️ Configuration

### Base Directory

By default, each script targets a specific directory. You can modify the `$baseDir` variable at the top of each script:

```powershell
# Set the base directory to scan
$baseDir = "C:\Users\YourUsername\OneDrive - YourOrganization"
```

### Duplicate Detection Logic

The script identifies duplicates based on:
- Files with identical size
- Files with identical names
- Files with similar names following common duplication patterns:
  - `filename (1).ext`, `filename (2).ext`
  - `filename - Copy.ext`, `filename - Copy (1).ext`
  - `filename_1.ext`, `filename_2.ext`
  - `filename_copy.ext`, `filename_copy1.ext`

## 🛡️ Safety Features

- **Simulation Mode**: Preview changes before executing them
- **Selective Preservation**: Keeps the newest file in each duplicate set by default
- **Detailed Logging**: Clear information about what files are being kept and removed
- **Confirmation Prompts**: Requires explicit confirmation before removing files

## 🛠️ Requirements

- Windows PowerShell 5.1 or newer
- Administrator privileges may be required for some file operations
- Write access to the target directories

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📊 Planned Enhancements

- [ ] Hash-based verification for more accurate duplicate detection
- [ ] Priority folder configuration to control which copies are kept
- [ ] Enhanced logging with HTML reports
- [ ] GUI interface for easier use
- [ ] Support for additional cloud storage providers

## ⚠️ Disclaimer

Always back up your data before running file management scripts. While these scripts are designed to be safe, unexpected issues can occur. The author is not responsible for any data loss resulting from the use of these scripts.
