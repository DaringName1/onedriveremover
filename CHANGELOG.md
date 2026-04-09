# Changelog

## 1.1.0 - 2026-04-09

- Added advanced PowerShell switches for technical users
- Added `-CheckOnly` to report current OneDrive state without removing anything
- Added `-VerifyOnly` to report remaining OneDrive items after cleanup
- Added `-AutoApproveFolderMove` to skip the move confirmation prompt
- Added `-SkipOfficeFix` to skip Office cleanup
- Added `-SaveLogPath` to save the log automatically to a chosen file path

## 1.0.0 - 2026-04-09

- Added a Windows GUI with step-by-step progress and logging
- Added safer handling for Desktop, Documents, and Pictures when they are still inside OneDrive
- Added cloud-only file detection before folder migration
- Added stronger OneDrive shutdown and local account-state cleanup
- Added guarded delete logic to avoid unsafe path removal
- Added Explorer/sidebar cleanup and reinstall-block policy
- Added final verification checks after removal
- Added a Save Log button for troubleshooting
- Added README, LICENSE, and GitHub release support files
