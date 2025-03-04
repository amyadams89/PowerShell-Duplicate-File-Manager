# Contributing to PowerShell Duplicate File Manager

Thank you for your interest in contributing to the PowerShell Duplicate File Manager! This document provides guidelines and instructions for contributing.

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment for everyone.

## How Can I Contribute?

### Reporting Bugs

Before submitting a bug report:

1. Check the Issues section to see if the bug has already been reported
2. Use the bug report template when creating a new issue
3. Include detailed steps to reproduce the problem
4. Include the PowerShell version and operating system information

### Suggesting Features

Feature suggestions are always welcome:

1. Check if the feature has already been suggested or implemented
2. Use the feature request template
3. Provide clear use-cases for the feature
4. Explain how the feature would benefit users

### Pull Requests

1. Fork the repository
2. Create a new branch for your feature or fix
3. Write your code following the coding standards below
4. Add or update relevant tests
5. Update documentation to reflect your changes
6. Submit a pull request with a clear description of the changes

## Development Setup

1. Clone the repository: `git clone https://github.com/yourusername/PowerShell-Duplicate-File-Manager.git`
2. Install recommended PowerShell modules for development:
   ```powershell
   Install-Module -Name PSScriptAnalyzer -Scope CurrentUser
   Install-Module -Name Pester -Scope CurrentUser
   ```

## PowerShell Coding Standards

### Script Format

- Use 4 spaces for indentation, not tabs
- Use meaningful variable and function names
- Add comment-based help to all functions
- Keep line length under 100 characters when possible

### PowerShell Best Practices

- Use the PowerShell verb-noun naming convention for functions
- Use Pascal case for function names (e.g., `Get-DuplicateFile`)
- Use camel case for variable names (e.g., `$fileCount`)
- Use proper error handling with try/catch blocks
- Use parameter validation attributes
- Avoid global variables

### Example

```powershell
function Get-FileHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath
    )
    
    try {
        $hashValue = Get-FileHash -Path $FilePath -Algorithm SHA256
        return $hashValue.Hash
    }
    catch {
        Write-Error "Failed to get hash for file: $_"
    }
}
```

## Testing Guidelines

- Write Pester tests for all functions
- Tests should cover both normal operation and error cases
- Run all tests before submitting pull requests
- Use mock objects when testing functions that interact with the file system

## Documentation

- Update README.md with any new features or significant changes
- Add examples for new functionality
- Keep the parameter documentation up to date

Thank you for contributing to make PowerShell Duplicate File Manager better!
