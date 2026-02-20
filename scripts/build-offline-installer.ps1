param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "0.1.0",
    [switch]$SkipPublish
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$projectPath = Join-Path $repoRoot "PasteWinUI\PasteWinUI.csproj"
$publishDir = Join-Path $repoRoot "PasteWinUI\bin\x64\$Configuration\net8.0-windows10.0.19041.0\$Runtime\publish"
$publishExe = Join-Path $publishDir "PasteWinUI.exe"
$workDir = Join-Path $repoRoot "installer\iexpress-work"
$outputDir = Join-Path $repoRoot "installer\output"
$targetExe = Join-Path $outputDir "PasteWinUI-Setup-$Version-offline.exe"
$sedPath = Join-Path $workDir "package.sed"
$payloadZip = Join-Path $workDir "payload.zip"
$installCmd = Join-Path $workDir "install.cmd"
$installPs1 = Join-Path $workDir "install.ps1"

if (-not $SkipPublish) {
    Write-Host "Publishing $projectPath ..."
    dotnet publish $projectPath -c $Configuration -r $Runtime --self-contained true /p:WindowsAppSDKSelfContained=true
}

if (-not (Test-Path $publishExe)) {
    throw "Publish output was not found: $publishExe"
}

if (Test-Path $workDir) {
    Remove-Item $workDir -Recurse -Force
}
New-Item -ItemType Directory -Path $workDir | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

Write-Host "Creating payload archive ..."
Compress-Archive -Path (Join-Path $publishDir "*") -DestinationPath $payloadZip -Force

$installCmdContent = @'
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
exit /b %ERRORLEVEL%
'@
Set-Content -Path $installCmd -Value $installCmdContent -Encoding ASCII

$installPs1Content = @'
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$appName = "PasteWinUI"
$appVersion = "__APP_VERSION__"
$publisher = "PasteWinUI"
$targetDir = Join-Path $env:LOCALAPPDATA "Programs\PasteWinUI"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\PasteWinUI"
$payloadZip = Join-Path $PSScriptRoot "payload.zip"
$desktopShortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "PasteWinUI.lnk"
$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$runValueName = "PasteWinUI"
$uninstallKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\PasteWinUI"
$exePath = Join-Path $targetDir "PasteWinUI.exe"

function New-AppShortcut {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [string]$Arguments = ""
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = Split-Path -Parent $TargetPath
    $shortcut.IconLocation = $TargetPath
    $shortcut.Save()
}

function Ask-YesNo {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [bool]$DefaultYes = $true
    )

    while ($true) {
        if ($DefaultYes) {
            $raw = Read-Host "$Prompt [Y/n]"
            if ([string]::IsNullOrWhiteSpace($raw)) { return $true }
        }
        else {
            $raw = Read-Host "$Prompt [y/N]"
            if ([string]::IsNullOrWhiteSpace($raw)) { return $false }
        }

        switch ($raw.Trim().ToLowerInvariant()) {
            "y" { return $true }
            "yes" { return $true }
            "n" { return $false }
            "no" { return $false }
            default { Write-Host "Please answer y or n." }
        }
    }
}

New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
Expand-Archive -Path $payloadZip -DestinationPath $targetDir -Force

New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null
$startMenuShortcut = Join-Path $startMenuDir "PasteWinUI.lnk"
New-AppShortcut -ShortcutPath $startMenuShortcut -TargetPath $exePath

$createDesktopShortcut = Ask-YesNo -Prompt "Create desktop shortcut?" -DefaultYes $true
if ($createDesktopShortcut) {
    New-AppShortcut -ShortcutPath $desktopShortcutPath -TargetPath $exePath
}

$enableAutoStart = Ask-YesNo -Prompt "Enable auto start at login?" -DefaultYes $false
if ($enableAutoStart) {
    New-Item -Path $runKeyPath -Force | Out-Null
    New-ItemProperty -Path $runKeyPath -Name $runValueName -Value ('"{0}"' -f $exePath) -PropertyType String -Force | Out-Null
}
else {
    Remove-ItemProperty -Path $runKeyPath -Name $runValueName -ErrorAction SilentlyContinue
}

$uninstallScriptPath = Join-Path $targetDir "uninstall.ps1"
$uninstallScript = @"
Set-StrictMode -Version Latest
`$ErrorActionPreference = "Stop"

`$targetDir = "__TARGET_DIR__"
`$startMenuDir = "__START_MENU_DIR__"
`$desktopShortcutPath = "__DESKTOP_SHORTCUT__"
`$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
`$runValueName = "PasteWinUI"
`$uninstallKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\PasteWinUI"

Get-Process PasteWinUI -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-ItemProperty -Path `$runKeyPath -Name `$runValueName -ErrorAction SilentlyContinue
Remove-Item -Path `$desktopShortcutPath -ErrorAction SilentlyContinue
Remove-Item -Path (Join-Path `$startMenuDir "PasteWinUI.lnk") -ErrorAction SilentlyContinue
Remove-Item -Path (Join-Path `$startMenuDir "Uninstall PasteWinUI.lnk") -ErrorAction SilentlyContinue
Remove-Item -Path `$uninstallKeyPath -Recurse -Force -ErrorAction SilentlyContinue

`$cmd = 'ping 127.0.0.1 -n 2 >nul & rmdir /s /q "{0}" & rmdir /s /q "{1}"' -f `$targetDir, `$startMenuDir
Start-Process -FilePath "cmd.exe" -ArgumentList "/c `$cmd" -WindowStyle Hidden
"@
$uninstallScript = $uninstallScript.Replace("__TARGET_DIR__", $targetDir)
$uninstallScript = $uninstallScript.Replace("__START_MENU_DIR__", $startMenuDir)
$uninstallScript = $uninstallScript.Replace("__DESKTOP_SHORTCUT__", $desktopShortcutPath)
Set-Content -Path $uninstallScriptPath -Value $uninstallScript -Encoding ASCII

$uninstallCommand = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $uninstallScriptPath
New-Item -Path $uninstallKeyPath -Force | Out-Null
New-ItemProperty -Path $uninstallKeyPath -Name "DisplayName" -Value $appName -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKeyPath -Name "DisplayVersion" -Value $appVersion -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKeyPath -Name "Publisher" -Value $publisher -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKeyPath -Name "InstallLocation" -Value $targetDir -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKeyPath -Name "DisplayIcon" -Value $exePath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKeyPath -Name "UninstallString" -Value $uninstallCommand -PropertyType String -Force | Out-Null

$uninstallShortcut = Join-Path $startMenuDir "Uninstall PasteWinUI.lnk"
New-AppShortcut -ShortcutPath $uninstallShortcut -TargetPath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$uninstallScriptPath`""

Start-Process -FilePath $exePath
'@
$installPs1Content = $installPs1Content.Replace("__APP_VERSION__", $Version)
Set-Content -Path $installPs1 -Value $installPs1Content -Encoding ASCII

$sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=$targetExe
FriendlyName=PasteWinUI Setup
AppLaunched=install.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=install.cmd
UserQuietInstCmd=install.cmd
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$workDir
[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
[Strings]
FILE0=install.cmd
FILE1=install.ps1
FILE2=payload.zip
"@
Set-Content -Path $sedPath -Value $sed -Encoding ASCII

Write-Host "Building offline installer with IExpress ..."
$process = Start-Process -FilePath iexpress.exe -ArgumentList "/N", "/Q", $sedPath -Wait -PassThru
if ($process.ExitCode -ne 0) {
    throw "IExpress failed. ExitCode=$($process.ExitCode)"
}

if (-not (Test-Path $targetExe)) {
    throw "Installer was not created: $targetExe"
}

Write-Host "Installer build completed."
Write-Host "Output: $targetExe"
