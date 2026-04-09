#Requires -Version 5.1
# ============================================================
#  OneDrive Remover v1.1.0 - Handles clean, partial & broken installs
#  Launched automatically as Admin via the desktop shortcut
# ============================================================

[CmdletBinding()]
param(
    [switch]$AutoApproveFolderMove,
    [switch]$SkipOfficeFix,
    [string]$SaveLogPath,
    [switch]$CheckOnly,
    [switch]$VerifyOnly
)

try {

$AppVersion = "1.1.0"
$selfPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $null }
$isCompiledExe = $selfPath -and ([System.IO.Path]::GetExtension($selfPath) -ieq ".exe")
$scriptDir = if ($selfPath) { Split-Path -Parent $selfPath } else { (Get-Location).Path }
$iconPath = Join-Path $scriptDir "app-icon.ico"

# Unblock only the raw script file. In compiled EXE mode there may be no PS1 path.
if ($selfPath -and -not $isCompiledExe -and (Test-Path -LiteralPath $selfPath)) {
    Unblock-File -Path $selfPath -ErrorAction SilentlyContinue
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-OneDriveStatusReport {
    $odBase = "$env:USERPROFILE\OneDrive"
    $folderMap = @{
        Desktop   = [Environment]::GetFolderPath("Desktop")
        Documents = [Environment]::GetFolderPath("MyDocuments")
        Pictures  = [Environment]::GetFolderPath("MyPictures")
    }
    $knownFoldersInOneDrive = @(
        $folderMap.Keys | Where-Object { $folderMap[$_] -like "$odBase*" }
    )
    $tasks = @(Get-ScheduledTask -EA SilentlyContinue | Where-Object { $_.TaskName -match "OneDrive" })
    $processes = @(Get-Process -Name "OneDrive","OneDriveSetup","FileCoAuth","FileSyncHelper" -EA SilentlyContinue | Select-Object -ExpandProperty ProcessName)
    $sidebarIds = @("{018D5C66-4533-4307-9B53-224DE2ED1FE6}","{04271989-C4F2-4BF7-BDD3-6A4CE03B45F2}","{0E5AAE11-A475-4c5b-AB00-C66DE400274E}")
    $sidebarPresent = $false
    foreach($id in $sidebarIds){
        foreach($ns in @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$id","HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$id")){
            if(Test-Path $ns){ $sidebarPresent = $true; break }
        }
        if($sidebarPresent){ break }
    }
    $gp = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
    $policyBlock = (Get-ItemProperty -Path $gp -Name "DisableFileSyncNGSC" -EA SilentlyContinue).DisableFileSyncNGSC -eq 1
    $uninstallers = @(
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
        "$env:SystemRoot\System32\OneDriveSetup.exe",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\Update\OneDriveSetup.exe",
        "$env:ProgramFiles\Microsoft OneDrive\OneDriveSetup.exe",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDriveSetup.exe"
    ) | Where-Object { Test-Path $_ }
    $appxPackages = @(Get-AppxPackage -AllUsers -EA SilentlyContinue | Where-Object { $_.Name -match "OneDrive" } | Select-Object -ExpandProperty Name)

    [PSCustomObject]@{
        OneDriveDataFolder     = $odBase
        KnownFoldersInOneDrive = $knownFoldersInOneDrive
        RunningProcesses       = $processes
        ScheduledTasks         = @($tasks | Select-Object -ExpandProperty TaskName)
        UninstallerPaths       = $uninstallers
        AppxPackages           = $appxPackages
        SidebarEntriesPresent  = $sidebarPresent
        PolicyBlockEnabled     = $policyBlock
    }
}

function Write-OneDriveStatusReport($title, $report) {
    Write-Output $title
    Write-Output ("Data folder: {0}" -f $report.OneDriveDataFolder)
    Write-Output ("Known folders in OneDrive: {0}" -f ($(if($report.KnownFoldersInOneDrive.Count){ $report.KnownFoldersInOneDrive -join ", " } else { "none" })))
    Write-Output ("Running processes: {0}" -f ($(if($report.RunningProcesses.Count){ $report.RunningProcesses -join ", " } else { "none" })))
    Write-Output ("Scheduled tasks: {0}" -f ($(if($report.ScheduledTasks.Count){ $report.ScheduledTasks -join ", " } else { "none" })))
    Write-Output ("Uninstallers found: {0}" -f ($(if($report.UninstallerPaths.Count){ $report.UninstallerPaths -join "; " } else { "none" })))
    Write-Output ("AppX packages: {0}" -f ($(if($report.AppxPackages.Count){ $report.AppxPackages -join ", " } else { "none" })))
    Write-Output ("Explorer sidebar entries present: {0}" -f $report.SidebarEntriesPresent)
    Write-Output ("Reinstall block policy enabled: {0}" -f $report.PolicyBlockEnabled)
}

if ($CheckOnly) {
    Write-OneDriveStatusReport "OneDrive check-only report" (Get-OneDriveStatusReport)
    exit
}

if ($VerifyOnly) {
    Write-OneDriveStatusReport "OneDrive verify-only report" (Get-OneDriveStatusReport)
    exit
}

# -- Admin check - re-launch elevated if not already ----------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
if (-not $isAdmin) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This tool needs Administrator rights to remove OneDrive safely.`n`nClick Yes to restart with Administrator access now.",
        "Administrator Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($result -eq "Yes") {
        if ($isCompiledExe -and $selfPath) {
            Start-Process -FilePath $selfPath -Verb RunAs
        } elseif ($selfPath) {
            Start-Process powershell.exe -ArgumentList "-Sta -NoProfile -ExecutionPolicy Bypass -File `"$selfPath`"" -Verb RunAs
        } else {
            throw "Could not determine the current script or executable path for elevation."
        }
    }
    exit
}

# -- Environment checks before building UI --------------------

# 1. Windows S Mode - blocks all unsigned apps and scripts entirely
$sMode = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy" `
          -Name "SkuPolicyRequired" -ErrorAction SilentlyContinue).SkuPolicyRequired
if ($sMode -eq 1) {
    [System.Windows.Forms.MessageBox]::Show(
        "Windows 11 S Mode detected.`n`nS Mode blocks all unsigned scripts and apps. OneDrive cannot be removed while S Mode is active.`n`nTo fix: Go to Settings - System - Activation - Switch out of S Mode (free). Then run this tool again.",
        "Windows S Mode - Cannot Continue",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit
}

# 2. Execution policy - detect if a machine-level lock is overriding Bypass
$effPolicy = (Get-ExecutionPolicy -Scope MachinePolicy)
if ($effPolicy -eq "Restricted" -or $effPolicy -eq "AllSigned") {
    $override = [System.Windows.Forms.MessageBox]::Show(
        "PowerShell is restricted on this PC ($effPolicy).`n`nThe tool can still try to continue, but some cleanup steps may be blocked.`n`nIf this is a work or school computer, IT may need to allow it.`n`nContinue anyway?",
        "Execution Policy Warning",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($override -ne "Yes") { exit }
}

# 3. UAC disabled - elevation silently fails with no prompt
$uacEnabled = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
               -Name "EnableLUA" -ErrorAction SilentlyContinue).EnableLUA
if ($uacEnabled -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "User Account Control (UAC) is disabled on this system.`n`nThis tool requires UAC to be enabled to run with Administrator rights.`n`nTo fix: Search 'UAC' in Start - set slider to at least the second notch - restart - run this tool again.",
        "UAC Disabled - Cannot Continue",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit
}


# 4. Known Folder Move detection - if Desktop/Documents/Pictures
#    are inside OneDrive, copy them back to local profile folders first.
$odBase    = "$env:USERPROFILE\OneDrive"
$kfmHit    = @()
$folderMap  = @{
    Desktop   = [Environment]::GetFolderPath("Desktop")
    Documents = [Environment]::GetFolderPath("MyDocuments")
    Pictures  = [Environment]::GetFolderPath("MyPictures")
}
$folderTargets = @{
    Desktop   = Join-Path $env:USERPROFILE "Desktop"
    Documents = Join-Path $env:USERPROFILE "Documents"
    Pictures  = Join-Path $env:USERPROFILE "Pictures"
}
$folderRegistryNames = @{
    Desktop   = "Desktop"
    Documents = "Personal"
    Pictures  = "My Pictures"
}
foreach ($name in $folderMap.Keys) {
    if ($folderMap[$name] -like "$odBase*") { $kfmHit += $name }
}
if ($kfmHit.Count -gt 0) {
    if (-not $AutoApproveFolderMove) {
        $detailLines = ($kfmHit | ForEach-Object { "  {0}: {1}" -f $_, $folderMap[$_] }) -join "`n"
        $kfmResult = [System.Windows.Forms.MessageBox]::Show(
            "Some Windows folders are still inside OneDrive:`n$detailLines`n`nThis tool can copy them back to this PC first, switch Windows to those local folders, and then remove OneDrive. The old OneDrive copy will stay there as a backup.`n`nDo you want to do that now?",
            "Move Folders Out Of OneDrive",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($kfmResult -ne "Yes") { exit }
    }
}

$BG      = [System.Drawing.Color]::FromArgb(13, 13, 18)
$PANEL   = [System.Drawing.Color]::FromArgb(22, 22, 32)
$ACCENT  = [System.Drawing.Color]::FromArgb(30, 130, 255)
$SUCCESS = [System.Drawing.Color]::FromArgb(0, 210, 110)
$WARN    = [System.Drawing.Color]::FromArgb(255, 185, 0)
$DANGER  = [System.Drawing.Color]::FromArgb(255, 65, 65)
$TEXT    = [System.Drawing.Color]::FromArgb(235, 235, 245)
$DIM     = [System.Drawing.Color]::FromArgb(110, 110, 135)

$FTitle = New-Object System.Drawing.Font("Segoe UI", 17, [System.Drawing.FontStyle]::Bold)
$FSub   = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Regular)
$FMono  = New-Object System.Drawing.Font("Consolas",  9, [System.Drawing.FontStyle]::Regular)
$FBtn   = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$FStep  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)

# -- Form -----------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text            = "OneDrive Remover v$AppVersion"
$form.Size            = New-Object System.Drawing.Size(580, 730)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $BG
$form.ForeColor       = $TEXT
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
if (Test-Path -LiteralPath $iconPath) {
    $form.Icon = New-Object System.Drawing.Icon($iconPath)
} else {
    $form.Icon = [System.Drawing.SystemIcons]::Shield
}

# Header
$hdr = New-Object System.Windows.Forms.Panel
$hdr.Size = New-Object System.Drawing.Size(580, 88); $hdr.Location = New-Object System.Drawing.Point(0,0); $hdr.BackColor = $PANEL
$t1 = New-Object System.Windows.Forms.Label; $t1.Text = "  OneDrive Remover"; $t1.Font = $FTitle; $t1.ForeColor = $TEXT; $t1.Location = New-Object System.Drawing.Point(10,10); $t1.AutoSize = $true
$t2 = New-Object System.Windows.Forms.Label; $t2.Text = "  v$AppVersion - Safe OneDrive removal for Windows 10 and 11"; $t2.Font = $FSub; $t2.ForeColor = $DIM; $t2.Location = New-Object System.Drawing.Point(12,50); $t2.AutoSize = $true
$hdr.Controls.AddRange(@($t1,$t2)); $form.Controls.Add($hdr)

# Banner
$ban = New-Object System.Windows.Forms.Panel; $ban.Size = New-Object System.Drawing.Size(528,44); $ban.Location = New-Object System.Drawing.Point(26,102); $ban.BackColor = [System.Drawing.Color]::FromArgb(0,55,28)
$bl  = New-Object System.Windows.Forms.Label; $bl.Text = "  This tool removes the OneDrive app, keeps your files, and can move Windows folders back to this PC first."; $bl.Font = $FSub; $bl.ForeColor = $SUCCESS; $bl.Location = New-Object System.Drawing.Point(8,7); $bl.Size = New-Object System.Drawing.Size(514,30)
$ban.Controls.Add($bl); $form.Controls.Add($ban)

# Steps
$stepDefs = @(
    @{I="Move";   T="Copy known folders out of OneDrive when needed"},
    @{I="Stop";   T="Kill all OneDrive, Office and Teams processes"},
    @{I="Unlink"; T="Unlink account and clear all scheduled tasks"},
    @{I="Uninst"; T="Run every available uninstaller (all locations)"},
    @{I="AppX";   T="Remove AppX / Microsoft Store version"},
    @{I="Reg";    T="Scrub all registry keys and startup entries"},
    @{I="Side";   T="Remove from Explorer sidebar - restart Explorer"},
    @{I="Files";  T="Force-delete all leftover app files and folders"},
    @{I="GPO";    T="Block OneDrive from reinstalling via Group Policy"},
    @{I="Office"; T="Fix Office save location and clean up side effects"}
)
$dots = @(); $sLbls = @(); $yB = 162
for ($i=0; $i -lt $stepDefs.Count; $i++) {
    $d = New-Object System.Windows.Forms.Label; $d.Text = "-"; $d.Font = New-Object System.Drawing.Font("Segoe UI",11); $d.ForeColor = $DIM; $d.Location = New-Object System.Drawing.Point(26,($yB+$i*45)); $d.AutoSize=$true
    $l = New-Object System.Windows.Forms.Label; $l.Text = $stepDefs[$i].T; $l.Font = $FStep; $l.ForeColor = $DIM; $l.Location = New-Object System.Drawing.Point(52,($yB+$i*45+2)); $l.AutoSize=$true
    $form.Controls.AddRange(@($d,$l)); $dots+=$d; $sLbls+=$l
}

# Progress + status
$prog = New-Object System.Windows.Forms.ProgressBar; $prog.Size = New-Object System.Drawing.Size(528,7); $prog.Location = New-Object System.Drawing.Point(26,535); $prog.Minimum=0; $prog.Maximum=$stepDefs.Count; $prog.Style="Continuous"; $form.Controls.Add($prog)
$lblSt = New-Object System.Windows.Forms.Label; $lblSt.Text="Ready - press Run to begin"; $lblSt.Font=$FSub; $lblSt.ForeColor=$DIM; $lblSt.Location=New-Object System.Drawing.Point(26,550); $lblSt.Size=New-Object System.Drawing.Size(528,18); $form.Controls.Add($lblSt)

# Log
$log = New-Object System.Windows.Forms.RichTextBox; $log.Size=New-Object System.Drawing.Size(528,96); $log.Location=New-Object System.Drawing.Point(26,574); $log.BackColor=[System.Drawing.Color]::FromArgb(8,8,13); $log.ForeColor=$DIM; $log.Font=$FMono; $log.ReadOnly=$true; $log.BorderStyle="None"; $log.ScrollBars="Vertical"; $form.Controls.Add($log)

# Buttons
$btn = New-Object System.Windows.Forms.Button; $btn.Text="Run - Remove OneDrive"; $btn.Font=$FBtn; $btn.Size=New-Object System.Drawing.Size(360,52); $btn.Location=New-Object System.Drawing.Point(26,676); $btn.BackColor=$ACCENT; $btn.ForeColor=$TEXT; $btn.FlatStyle="Flat"; $btn.FlatAppearance.BorderSize=0; $btn.Cursor=[System.Windows.Forms.Cursors]::Hand; $form.Controls.Add($btn)
$btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text="Save Log"; $btnSave.Font=$FBtn; $btnSave.Size=New-Object System.Drawing.Size(156,52); $btnSave.Location=New-Object System.Drawing.Point(398,676); $btnSave.BackColor=[System.Drawing.Color]::FromArgb(32,32,44); $btnSave.ForeColor=$TEXT; $btnSave.FlatStyle="Flat"; $btnSave.FlatAppearance.BorderSize=0; $btnSave.Cursor=[System.Windows.Forms.Cursors]::Hand; $btnSave.Enabled=$false; $form.Controls.Add($btnSave)

# Helpers
function L($m,$c){ $log.SelectionStart=$log.TextLength; $log.SelectionLength=0; if($c){$log.SelectionColor=$c}; $log.AppendText("$m`n"); $log.SelectionColor=$DIM; $log.ScrollToCaret(); [System.Windows.Forms.Application]::DoEvents() }
function Mark($i,$s){ switch($s){ "run"{$dots[$i].ForeColor=$ACCENT;$sLbls[$i].ForeColor=$TEXT} "ok"{$dots[$i].ForeColor=$SUCCESS;$sLbls[$i].ForeColor=$SUCCESS} "warn"{$dots[$i].ForeColor=$WARN;$sLbls[$i].ForeColor=$WARN} "err"{$dots[$i].ForeColor=$DANGER;$sLbls[$i].ForeColor=$DANGER} }; [System.Windows.Forms.Application]::DoEvents() }
function Export-RemovalLog(){
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = "Save Removal Log"
    $dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $dialog.FileName = "OneDrive-Removal-Log-{0}.txt" -f (Get-Date -Format "yyyyMMdd-HHmmss")
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        [System.IO.File]::WriteAllText($dialog.FileName, $log.Text, [System.Text.UTF8Encoding]::new($false))
        [System.Windows.Forms.MessageBox]::Show(
            "Log saved to:`n$($dialog.FileName)",
            "Log Saved",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}
function Save-LogIfRequested(){
    if([string]::IsNullOrWhiteSpace($SaveLogPath)){ return }
    $parent = Split-Path -Parent $SaveLogPath
    if(-not [string]::IsNullOrWhiteSpace($parent)){
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null
    }
    [System.IO.File]::WriteAllText($SaveLogPath, $log.Text, [System.Text.UTF8Encoding]::new($false))
}
function Get-CloudOnlyFiles($path){
    if(-not (Test-Path -LiteralPath $path)){ return @() }

    $offlineFlag = [System.IO.FileAttributes]::Offline
    Get-ChildItem -LiteralPath $path -Recurse -Force -File -EA SilentlyContinue |
        Where-Object { (($_.Attributes -band $offlineFlag) -eq $offlineFlag) }
}
function Sync-KnownFolderToLocal($name, $sourcePath, $targetPath, $registryName){
    if(-not (Test-Path -LiteralPath $targetPath)){
        New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
    }

    if((Test-Path -LiteralPath $sourcePath) -and ((Resolve-Path -LiteralPath $sourcePath).Path.TrimEnd('\') -ne (Resolve-Path -LiteralPath $targetPath).Path.TrimEnd('\'))){
        $cloudOnly = @(Get-CloudOnlyFiles -path $sourcePath)
        if($cloudOnly.Count -gt 0){
            $sample = ($cloudOnly | Select-Object -First 3 | ForEach-Object { $_.FullName }) -join "; "
            throw "Some $name files are still cloud-only. Open OneDrive, right-click that folder, choose 'Always keep on this device', wait for sync to finish, then run this tool again. Example: $sample"
        }
        L "    Copying $name to $targetPath" $DIM
        $robocopyArgs = @(
            "`"$sourcePath`"",
            "`"$targetPath`"",
            "/E","/COPY:DAT","/DCOPY:DAT","/R:1","/W:1","/XJ"
        )
        & robocopy @robocopyArgs | Out-Null
        $rc = $LASTEXITCODE
        if($rc -ge 8){
            throw "Copy failed for $name (robocopy exit $rc)."
        }
        L "    Copied $name to local profile." $SUCCESS
    } else {
        L "    $name already points to a local folder." $DIM
    }

    $userShellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $shellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    Set-ItemProperty -Path $userShellFolders -Name $registryName -Value $targetPath -EA Stop
    Set-ItemProperty -Path $shellFolders -Name $registryName -Value $targetPath -EA SilentlyContinue
    L "    Windows now uses local $name at $targetPath" $SUCCESS
}
function Clear-OneDriveAccountState(){
    foreach($settingsPath in @(
        "$env:LOCALAPPDATA\Microsoft\OneDrive\settings",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\logs\Personal",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\logs\Business1",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\logs\Business2"
    )){
        if(Test-Path -LiteralPath $settingsPath){
            try{
                Remove-Item -LiteralPath $settingsPath -Recurse -Force -EA Stop
                L "  Cleared account state: $settingsPath" $SUCCESS
            }catch{
                L "  Could not clear account state yet: $settingsPath" $WARN
            }
        }
    }
}
function Verify-RemovalState(){
    $issues = New-Object System.Collections.ArrayList

    $running = Get-Process -Name "OneDrive","FileCoAuth","FileSyncHelper" -EA SilentlyContinue
    if($running){ [void]$issues.Add("OneDrive processes are still running.") }

    $tasks = Get-ScheduledTask -EA SilentlyContinue | Where-Object {$_.TaskName -match "OneDrive"}
    if($tasks){ [void]$issues.Add("OneDrive scheduled tasks still exist.") }

    foreach($id in @("{018D5C66-4533-4307-9B53-224DE2ED1FE6}","{04271989-C4F2-4BF7-BDD3-6A4CE03B45F2}","{0E5AAE11-A475-4c5b-AB00-C66DE400274E}")){
        foreach($ns in @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$id","HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$id")){
            if(Test-Path $ns){ [void]$issues.Add("Explorer sidebar entry still exists: $id"); break }
        }
    }

    $gp = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
    $gpValue = (Get-ItemProperty -Path $gp -Name "DisableFileSyncNGSC" -EA SilentlyContinue).DisableFileSyncNGSC
    if($gpValue -ne 1){ [void]$issues.Add("Reinstall block policy is missing.") }

    return $issues
}
function Test-OneDriveRemovalTarget($p){
    if([string]::IsNullOrWhiteSpace($p)){ return $false }

    try {
        $resolved = [System.IO.Path]::GetFullPath($p.TrimEnd('\'))
    } catch {
        return $false
    }

    $protected = @(
        [System.IO.Path]::GetFullPath("$env:USERPROFILE\OneDrive".TrimEnd('\')),
        [System.IO.Path]::GetFullPath($env:USERPROFILE.TrimEnd('\')),
        [System.IO.Path]::GetFullPath($env:SystemDrive.TrimEnd('\') + '\')
    )
    if($protected -contains $resolved){ return $false }

    $allowedBases = @(
        [System.IO.Path]::GetFullPath("$env:LOCALAPPDATA\Microsoft\OneDrive".TrimEnd('\')),
        [System.IO.Path]::GetFullPath("$env:ProgramFiles\Microsoft OneDrive".TrimEnd('\')),
        [System.IO.Path]::GetFullPath("${env:ProgramFiles(x86)}\Microsoft OneDrive".TrimEnd('\')),
        [System.IO.Path]::GetFullPath("$env:ProgramData\Microsoft OneDrive".TrimEnd('\')),
        [System.IO.Path]::GetFullPath("$env:ProgramData\Microsoft\OneDrive".TrimEnd('\'))
    ) | Select-Object -Unique

    foreach($base in $allowedBases){
        if($resolved -eq $base -or $resolved.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)){
            return $true
        }
    }

    return $false
}
function FDel($p){
    if(-not(Test-Path -LiteralPath $p)){return}
    if(-not (Test-OneDriveRemovalTarget $p)){
        L "    Skipped unsafe delete target: $p" $WARN
        return
    }
    try{
        Remove-Item -LiteralPath $p -Recurse -Force -EA Stop
        L "    Removed: $p" $SUCCESS
    }catch{
        try{
            & takeown /f "$p" /r /d y 2>$null|Out-Null
            & icacls "$p" /grant administrators:F /t 2>$null|Out-Null
            Remove-Item -LiteralPath $p -Recurse -Force -EA SilentlyContinue
            L "    Force-removed: $p" $WARN
        }catch{
            L "    Locked (will clear on reboot): $p" $DIM
        }
    }
}

$script:done = $false

$btnSave.Add_Click({
    Export-RemovalLog
})

$btn.Add_Click({
    # Guard: if already finished, treat button as restart prompt
    if($script:done){
        if([System.Windows.Forms.MessageBox]::Show("Restart now?","Restart",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes"){Restart-Computer -Force}
        return
    }
    try {
        $btn.Enabled=$false; $btn.Text="Working..."; $btn.BackColor=[System.Drawing.Color]::FromArgb(40,40,60); $log.Clear()

        # STEP 0 - Move known folders back to local profile if needed
        Mark 0 "run"; $lblSt.Text="Copying known folders out of OneDrive when needed..."
        if($kfmHit.Count -gt 0){
            try{
                foreach($name in $kfmHit){
                    Sync-KnownFolderToLocal -name $name -sourcePath $folderMap[$name] -targetPath $folderTargets[$name] -registryName $folderRegistryNames[$name]
                }
                L "  Local folder copies created. OneDrive originals were left in place." $SUCCESS
                $folderMap = @{
                    Desktop   = $folderTargets["Desktop"]
                    Documents = $folderTargets["Documents"]
                    Pictures  = $folderTargets["Pictures"]
                }
                [Environment]::SetEnvironmentVariable("OneDrive", $null, "User")
            } catch {
                Mark 0 "err"
                $lblSt.ForeColor=$DANGER; $lblSt.Text="Could not move folders out of OneDrive safely."
                L "  Failed before removal: $($_.Exception.Message)" $DANGER
                $btn.Enabled=$true; $btn.Text="Run - Remove OneDrive"; $btn.BackColor=$ACCENT
                $btnSave.Enabled=$true
                return
            }
        } else {
            L "  No Desktop/Documents/Pictures folders are inside OneDrive." $SUCCESS
        }
        $prog.Value=1; Mark 0 "ok"

        # STEP 1 - Kill all processes
        Mark 1 "run"; $lblSt.Text="Killing all OneDrive, Office and Teams processes..."
        $killList = @(
            # OneDrive processes
            "OneDrive","OneDriveSetup","OneDriveStandaloneUpdater","FileCoAuth","FileSyncHelper",
            # Office processes that hold OneDrive handles
            "WINWORD","EXCEL","POWERPNT","OUTLOOK","ONENOTE","MSACCESS","MSPUB",
            # Teams - has its own OneDrive integration
            "Teams","ms-teams"
        )
        foreach($p in $killList){
            Get-Process -Name $p -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
            & taskkill /f /im "$p.exe" 2>$null | Out-Null
            L "  Stopped: $p" $DIM
        }
        Start-Sleep -Milliseconds 700; $prog.Value=2; Mark 1 "ok"

        # STEP 2 - Unlink + scheduled tasks
        Mark 2 "run"; $lblSt.Text="Unlinking account & clearing scheduled tasks..."
        $odExe = @(
            "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
            "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe",
            "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
        ) | Where-Object {Test-Path $_} | Select-Object -First 1
        if($odExe){
            Start-Process $odExe -ArgumentList "/shutdown" -EA SilentlyContinue
            Start-Sleep -Milliseconds 1200
            Get-Process -Name "OneDrive","FileCoAuth","FileSyncHelper" -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
            Start-Sleep -Milliseconds 500
        }
        Get-ScheduledTask -EA SilentlyContinue | Where-Object {$_.TaskName -match "OneDrive"} | ForEach-Object {
            Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false -EA SilentlyContinue
            L "  Removed task: $($_.TaskName)" $SUCCESS
        }
        foreach($hive in @("HKCU:","HKLM:")){
            $run="$hive\Software\Microsoft\Windows\CurrentVersion\Run"
            Remove-ItemProperty $run -Name "OneDrive" -EA SilentlyContinue
            Remove-ItemProperty $run -Name "OneDriveSetup" -EA SilentlyContinue
        }
        Clear-OneDriveAccountState
        $prog.Value=3; Mark 2 "ok"

        # STEP 3 - All uninstallers
        Mark 3 "run"; $lblSt.Text="Running all uninstallers..."
        $unins = @(
            "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
            "$env:SystemRoot\System32\OneDriveSetup.exe",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\Update\OneDriveSetup.exe",
            "$env:ProgramFiles\Microsoft OneDrive\OneDriveSetup.exe",
            "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDriveSetup.exe"
        )
        $ran=$false
        foreach($u in $unins){
            if(Test-Path $u){
                L "  Running: $u" $DIM
                $r = Start-Process $u -ArgumentList "/uninstall /silent /allusers" -Wait -PassThru -EA SilentlyContinue
                if ($r -and $r.ExitCode -eq 0) {
                    L "  Uninstall succeeded (exit 0)." $SUCCESS; $ran=$true
                } else {
                    $code = if($r) { $r.ExitCode } else { "n/a" }
                    L "  /allusers failed (exit $code) - retrying user-mode..." $WARN
                    $r2 = Start-Process $u -ArgumentList "/uninstall /silent" -Wait -PassThru -EA SilentlyContinue
                    if ($r2 -and $r2.ExitCode -eq 0) {
                        L "  User-mode uninstall succeeded." $SUCCESS; $ran=$true
                    } else {
                        $code2 = if($r2) { $r2.ExitCode } else { "n/a" }
                        L "  Both uninstall modes failed (exit $code2) - forcing file removal." $WARN
                    }
                }
            }
        }
        if(-not $ran){
            L "  Uninstallers failed or not found - trying winget..." $WARN
            $wingetCmd = Get-Command winget.exe -ErrorAction SilentlyContinue
            if($wingetCmd){
                Start-Process $wingetCmd.Source -ArgumentList "uninstall --id Microsoft.OneDrive --silent --accept-source-agreements --force" -Wait -EA SilentlyContinue
            } else {
                L "  winget not available in this admin session." $DIM
            }
            L "  Skipped Win32_Product fallback to avoid Windows Installer self-repair side effects." $DIM
        }
        Start-Sleep -Milliseconds 500; $prog.Value=4; Mark 3 "ok"

        # STEP 4 - AppX
        Mark 4 "run"; $lblSt.Text="Removing AppX / Store version..."
        Get-AppxPackage -AllUsers -EA SilentlyContinue | Where-Object {$_.Name -match "OneDrive"} | ForEach-Object {
            Remove-AppxPackage -Package $_.PackageFullName -AllUsers -EA SilentlyContinue
            L "  Removed AppX: $($_.Name)" $SUCCESS
        } else {
            L "  No AppX OneDrive package found." $DIM
        }
        Get-AppxProvisionedPackage -Online -EA SilentlyContinue | Where-Object {$_.PackageName -match "OneDrive"} | ForEach-Object {
            Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -EA SilentlyContinue | Out-Null
            L "  Removed provisioned: $($_.PackageName)" $SUCCESS
        }
        $prog.Value=5; Mark 4 "ok"

        # STEP 5 - Registry & Environment
        Mark 5 "run"; $lblSt.Text="Scrubbing registry & variables..."
        foreach($hive in @("HKCU:","HKLM:")){
            $run="$hive\Software\Microsoft\Windows\CurrentVersion\Run"
            Remove-ItemProperty $run -Name "OneDrive" -EA SilentlyContinue
            Remove-ItemProperty $run -Name "OneDriveSetup" -EA SilentlyContinue
        }
        # Clean stray environment variables
        foreach($var in @("OneDrive","OneDriveConsumer")){
            [Environment]::SetEnvironmentVariable($var, $null, "User")
            [Environment]::SetEnvironmentVariable($var, $null, "Machine")
        }
        $keys=@(
            "HKCU:\Software\Microsoft\OneDrive",
            "HKLM:\SOFTWARE\Microsoft\OneDrive",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{04271989-C4F2-4BF7-BDD3-6A4CE03B45F2}",
            "HKCU:\Software\Classes\Directory\Background\shellex\ContextMenuHandlers\FileSyncEx",
            "HKCU:\Software\Classes\Directory\shellex\ContextMenuHandlers\FileSyncEx",
            "HKCU:\Software\Classes\*\shellex\ContextMenuHandlers\FileSyncEx"
        )
        foreach($k in $keys){ try{ if(Test-Path $k){ Remove-Item $k -Recurse -Force -EA SilentlyContinue; L "  Removed: $k" $SUCCESS } }catch{} }
        $prog.Value=6; Mark 5 "ok"

        # STEP 6 - Explorer sidebar
        Mark 6 "run"; $lblSt.Text="Removing from Explorer sidebar..."
        $ids=@("{018D5C66-4533-4307-9B53-224DE2ED1FE6}","{04271989-C4F2-4BF7-BDD3-6A4CE03B45F2}","{0E5AAE11-A475-4c5b-AB00-C66DE400274E}")
        foreach($id in $ids){
            foreach($root in @("HKCU:\Software\Classes\CLSID","HKLM:\SOFTWARE\Classes\CLSID")){
                $p="$root\$id"; try{ if(-not(Test-Path $p)){New-Item $p -Force -EA SilentlyContinue|Out-Null}; Set-ItemProperty -Path $p -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Type DWord -EA SilentlyContinue }catch{}
            }
            foreach($ns in @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$id","HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$id")){ try{Remove-Item $ns -Recurse -Force -EA SilentlyContinue}catch{} }
        }
        L "  Sidebar cleared." $SUCCESS
        # Restart Explorer so sidebar changes take effect immediately
        L "  Restarting Explorer..." $DIM
        Stop-Process -Name "explorer" -Force -EA SilentlyContinue
        Start-Sleep -Milliseconds 1200
        Start-Process "explorer.exe"
        L "  Explorer restarted." $SUCCESS
        $prog.Value=7; Mark 6 "ok"

        # STEP 7 - Force-delete leftover files
        Mark 7 "run"; $lblSt.Text="Force-deleting leftover files (NOT your data)..."
        foreach($p in @("OneDrive","OneDriveSetup","FileCoAuth")){ & taskkill /f /im "$p.exe" 2>$null|Out-Null }
        Start-Sleep -Milliseconds 400
        foreach($t in @(
            "$env:LOCALAPPDATA\Microsoft\OneDrive\setup",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\logs",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\Update",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\StandaloneUpdater",
            "$env:LOCALAPPDATA\Microsoft\OneDrive\extensions",
            "$env:LOCALAPPDATA\Microsoft\OneDrive",
            "$env:ProgramFiles\Microsoft OneDrive",
            "${env:ProgramFiles(x86)}\Microsoft OneDrive",
            "$env:ProgramData\Microsoft OneDrive",
            "$env:ProgramData\Microsoft\OneDrive"
        )){ FDel $t }
        # Remove the OneDriveSetup.exe files left in System32/SysWOW64
        foreach($f in @("$env:SystemRoot\SysWOW64\OneDriveSetup.exe","$env:SystemRoot\System32\OneDriveSetup.exe")){
            if(Test-Path $f){ try{Remove-Item $f -Force -EA SilentlyContinue; L "    Removed: $f" $SUCCESS}catch{L "    Locked (system): $f" $DIM} }
        }
        L "  User data preserved: $env:USERPROFILE\OneDrive" $SUCCESS
        $prog.Value=8; Mark 7 "ok"

        # STEP 8 - Group Policy block
        Mark 8 "run"; $lblSt.Text="Blocking reinstall via Group Policy..."
        try{
            $gp="HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
            if(-not(Test-Path $gp)){New-Item $gp -Force|Out-Null}
            Set-ItemProperty -Path $gp -Name "DisableFileSyncNGSC" -Value 1 -Type DWord
            Set-ItemProperty -Path $gp -Name "DisableLibrariesDefaultSaveToOneDrive" -Value 1 -Type DWord -EA SilentlyContinue
            L "  Group Policy applied." $SUCCESS
        }catch{ L "  GP failed (non-admin). Other steps still complete." $WARN }
        $prog.Value=9; Mark 8 "ok"

        # STEP 9 - Fix Office default save location + clean side effects
        Mark 9 "run"; $lblSt.Text="Fixing Office save location and side effects..."
        if($SkipOfficeFix){
            L "  Skipped Office cleanup because -SkipOfficeFix was used." $WARN
        } else {
            L "  Resetting Office default save path to Documents..." $DIM
            $docsPath = [Environment]::GetFolderPath("MyDocuments")
            foreach ($ver in @("16.0","15.0","14.0")) {
                foreach ($app in @("Word","Excel","PowerPoint","OneNote")) {
                    $oPath = "HKCU:\Software\Microsoft\Office\$ver\$app\Options"
                    if (Test-Path $oPath) {
                        Set-ItemProperty -Path $oPath -Name "DefaultPath" -Value $docsPath -EA SilentlyContinue
                        L "    Reset $app $ver save path - $docsPath" $SUCCESS
                    }
                }
                # Reset common Office cloud storage preference
                $commonPath = "HKCU:\Software\Microsoft\Office\$ver\Common\General"
                if (Test-Path $commonPath) {
                    Set-ItemProperty -Path $commonPath -Name "PreferCloudSaveLocations" -Value 0 -Type DWord -EA SilentlyContinue
                    Set-ItemProperty -Path $commonPath -Name "SkyDriveSignInOption" -Value 0 -Type DWord -EA SilentlyContinue
                }
            }
            # Clear Windows Backup OneDrive reference
            $buPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CloudExperienceHost\Intent\officeHub"
            if (Test-Path $buPath) {
                Set-ItemProperty -Path $buPath -Name "Intent" -Value 0 -Type DWord -EA SilentlyContinue
                L "  Cleared Windows Backup OneDrive reference." $SUCCESS
            }
            L "  Office side effects cleaned." $SUCCESS
        }
        $prog.Value=10; Mark 9 "ok"

        # Final verification
        L "  Checking final cleanup..." $DIM
        $verifyIssues = @(Verify-RemovalState)
        if($verifyIssues.Count -eq 0){
            $lblSt.ForeColor=$SUCCESS; $lblSt.Text="Complete! OneDrive removal checks passed. Restart recommended."
            L "" $null; L "--------------------------------------" $SUCCESS
            L "  OneDrive removed. Final checks passed." $SUCCESS
            L "  Your files in \OneDrive\ were NOT touched." $SUCCESS
            L "--------------------------------------" $SUCCESS
        } else {
            $lblSt.ForeColor=$WARN; $lblSt.Text="Finished with warnings. Review the log before restart."
            L "" $null; L "--------------------------------------" $WARN
            L "  Removal finished, but some items still need attention:" $WARN
            foreach($issue in $verifyIssues){ L "  - $issue" $WARN }
            L "  Save the log if you want a record before restarting." $WARN
            L "--------------------------------------" $WARN
        }
        $script:done = $true
        Save-LogIfRequested
        $btn.Text="Done - Click to Restart Now"; $btn.BackColor=$SUCCESS; $btn.ForeColor=[System.Drawing.Color]::FromArgb(10,10,15); $btn.Enabled=$true
        $btnSave.Enabled=$true
    } catch {
        $lblSt.ForeColor=$DANGER
        $lblSt.Text="The tool hit an error. Review the log."
        L "" $null
        L "  ERROR: $($_.Exception.Message)" $DANGER
        if($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber){
            L "  Line: $($_.InvocationInfo.ScriptLineNumber)" $WARN
        }
        L "  Save the log and restart only after reviewing the message above." $WARN
        $btn.Enabled=$true
        $btn.Text="Run - Remove OneDrive"
        $btn.BackColor=$ACCENT
        $btnSave.Enabled=$true
        Save-LogIfRequested
        [System.Windows.Forms.MessageBox]::Show(
            "The tool hit an error, but the details are now in the log window.`n`nClick Save Log if you want to keep a copy.",
            "OneDrive Remover Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

$form.ShowDialog()|Out-Null

} catch {
    # If anything fails before the GUI loads, show a clear error
    # with exact steps to fix it rather than just the raw error
    $errMsg   = $_.Exception.Message
    $errLine  = $_.InvocationInfo.ScriptLineNumber

    $fixSteps = @"
OneDrive Remover could not start.

Error (line $errLine): $errMsg

-- Common causes & fixes ------------------------

1. Defender quarantined the script:
   - Open Windows Security - Virus & threat protection
   - Click Protection history - find this app - Allow
   - Run again

2. File blocked as downloaded:
   - Right-click OneDrive-Remover.exe - Properties
   - Check "Unblock" at the bottom - OK - Run again

3. Running from a USB drive or network share:
   - Copy the EXE to your Desktop or C:\ first
   - Then run it from there

4. Windows S Mode:
   - Settings - System - Activation - Switch out of S Mode

5. Work/school computer with locked PowerShell:
   - Contact your IT administrator

6. Skip the EXE entirely - run the PS1 directly:
   - Open PowerShell as Administrator
   - Run: Set-ExecutionPolicy Bypass -Scope Process
   - Then: & "C:\Path\To\Remove-OneDrive.ps1"
"@

    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            $fixSteps,
            "OneDrive Remover - Startup Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {
        # Last resort - console output if even MessageBox fails
        Write-Host "`nFATAL ERROR: $errMsg (line $errLine)" -ForegroundColor Red
        Write-Host "Run PowerShell as Administrator and execute this script directly." -ForegroundColor Yellow
        Read-Host "`nPress Enter to exit"
    }
}
