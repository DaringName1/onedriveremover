# OneDrive Remover

Current version: `v1.1.3`

Removes OneDrive from Windows 10 and Windows 11 while preserving user files.

## What It Does

- Removes the OneDrive app
- Cleans OneDrive startup entries and related registry entries
- Removes Explorer sidebar integration
- Blocks normal OneDrive reinstall behavior
- Preserves your files
- Moves Desktop, Documents, and Pictures out of OneDrive first when needed
- Stops if those folders contain cloud-only files that are not fully downloaded
- Verifies the cleanup after removal and lets you save a log

## Requirements

- Windows 10 or Windows 11
- Administrator rights
- Not supported on Windows S Mode

## Download

Download the latest `OneDrive-Remover.exe` from the Releases page:

[Latest Release](https://github.com/DaringName1/onedriveremover/releases/latest)

## In The App

The app shows:

- The current version in the window title
- A short plain-English note explaining what the tool does
- Step-by-step progress while it runs
- A `Save Log` button for troubleshooting or support

## How To Use

1. Download `OneDrive-Remover.exe`
2. Right-click it and choose `Run as administrator`
3. Click `Run - Remove OneDrive`
4. If prompted, allow the tool to move Desktop, Documents, or Pictures out of OneDrive first
5. If the tool says some files are cloud-only, open OneDrive and choose `Always keep on this device`, then run the tool again
6. Restart Windows when the tool finishes

## Ways To Run It

### Option 1: Run the EXE

This is the easiest option for most people.

1. Download `OneDrive-Remover.exe`
2. Right-click it
3. Choose `Run as administrator`
4. Click `Run - Remove OneDrive`

### Option 2: Run the PS1 directly from an Administrator PowerShell window

1. Open PowerShell as Administrator
2. Go to the folder that contains the script
3. Run:

```powershell
.\Remove-OneDrive.ps1
```

### Option 3: Run the PS1 with execution-policy bypass

Use this if normal script execution is blocked.

1. Open PowerShell as Administrator
2. Go to the folder that contains the script
3. Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Remove-OneDrive.ps1
```

### Option 4: Run the PS1 from the current PowerShell session after setting process-only bypass

This only changes policy for the current PowerShell window.

```powershell
Set-ExecutionPolicy Bypass -Scope Process
.\Remove-OneDrive.ps1
```

## Advanced Usage

These switches are mainly useful when running the `PS1` directly.

### Check current OneDrive state without removing anything

```powershell
.\Remove-OneDrive.ps1 -CheckOnly
```

### Verify whether OneDrive remnants are still present

```powershell
.\Remove-OneDrive.ps1 -VerifyOnly
```

### Auto-approve moving Desktop, Documents, and Pictures out of OneDrive

```powershell
.\Remove-OneDrive.ps1 -AutoApproveFolderMove
```

### Skip Office save-location cleanup

```powershell
.\Remove-OneDrive.ps1 -SkipOfficeFix
```

### Save the log automatically to a file

```powershell
.\Remove-OneDrive.ps1 -SaveLogPath "C:\Temp\OneDrive-Removal-Log.txt"
```

### Example with multiple switches

```powershell
.\Remove-OneDrive.ps1 -AutoApproveFolderMove -SkipOfficeFix -SaveLogPath "C:\Temp\OneDrive-Removal-Log.txt"
```

## Safety Notes

- Your files are not deleted by this tool
- If Desktop, Documents, or Pictures are inside OneDrive, the tool copies them back to local Windows folders first
- The original OneDrive copy is left in place as a backup copy during that move step
- The tool shuts down OneDrive, removes its scheduled tasks, clears local OneDrive account state, and then uninstalls it

## Included Files

- `OneDrive-Remover.exe`
- `Remove-OneDrive.ps1`

## Release

Current public release:

[OneDrive Remover v1.1.3](https://github.com/DaringName1/onedriveremover/releases/tag/v1.1.3)

## Disclaimer

Use at your own risk. This tool is intended for normal personal Windows 10 and Windows 11 PCs.
