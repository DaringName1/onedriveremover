# Changelog

## 1.1.3 - 2026-04-09

- Extended final verification to flag leftover OneDrive AppX and provisioned packages
- Fixed permission-repair fallback to use the Administrators SID for non-English Windows installs
- Improved delete fallback logging when permission repair still cannot remove a path

## 1.1.2 - 2026-04-09

- Fixed known-folder detection so OneDrive for Business paths are treated as OneDrive-managed
- Fixed `-VerifyOnly` to use real verification logic instead of the check-only report
- Fixed `-SaveLogPath` so it also works with non-GUI `-CheckOnly` and `-VerifyOnly` modes

## 1.1.1 - 2026-04-09

- Fixed the AppX cleanup step error caused by a stray `else`
- Made the interface larger and easier to use
- Increased the log area and improved button layout

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
