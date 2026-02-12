@echo off
setlocal
cd /d "%~dp0"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "TOOLDIR=%~dp0"

REM If user dragged files/folders onto the BAT, run immediately
if not "%~1"=="" (
  "%PS%" -NoProfile -ExecutionPolicy Bypass -File "%TOOLDIR%app\LUFS-Lens.ps1" %*
  echo.
  pause
  exit /b
)

REM Double-click: open picker, write selections to temp list file, run PS1 using @listfile
"%PS%" -NoProfile -STA -ExecutionPolicy Bypass -Command ^
  "$tool = $env:TOOLDIR; " ^
  "Set-Location -LiteralPath $tool; " ^
  "Add-Type -AssemblyName System.Windows.Forms; " ^
  "$temp = Join-Path $tool 'temp'; " ^
  "if(-not (Test-Path -LiteralPath $temp)) { New-Item -ItemType Directory -Path $temp | Out-Null } " ^
  "$d = New-Object System.Windows.Forms.OpenFileDialog; " ^
  "$d.Title = 'Select audio files for LUFS Lens'; " ^
  "$d.Filter = 'Audio Files|*.wav;*.flac;*.aif;*.aiff;*.mp3;*.m4a|All Files|*.*'; " ^
  "$d.Multiselect = $true; " ^
  "if($d.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit 2 } " ^
  "$list = Join-Path $temp 'LUFS_SelectedFiles.txt'; " ^
  "$d.FileNames | Set-Content -LiteralPath $list -Encoding Ascii; " ^
  "& (Join-Path $tool 'app\LUFS-Lens.ps1') ('@' + $list); " ^
  "exit $LASTEXITCODE"

echo.
pause
