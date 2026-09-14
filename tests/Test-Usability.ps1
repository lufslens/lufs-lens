$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$appRoot = Join-Path $project 'LUFS_Lens_v1'
if (-not (Test-Path -LiteralPath $appRoot)) { $appRoot = $project }
$ffmpegExe = Join-Path $appRoot 'ffmpeg\bin\ffmpeg.exe'
if (-not (Test-Path -LiteralPath $ffmpegExe)) { $ffmpegExe = (Get-Command ffmpeg -ErrorAction Stop).Source }
$fixture = Join-Path $project ('tests\run-' + [guid]::NewGuid().ToString('N'))
$originalPath = $env:PATH
function Clear-Host {}
function Start-Process { param($FilePath) }
function Read-Host { param($Prompt) return $global:lufsTestTargetAnswer }
$global:lufsProgressStages = @()
function Write-Progress {
  param($Id, $Activity, $Status, $PercentComplete, [switch]$Completed)
  if ($Completed) { $global:lufsProgressStages += 'Completed' }
  else { $global:lufsProgressStages += $Status }
}
try {
  New-Item -ItemType Directory -Path (Join-Path $fixture 'app') | Out-Null
  $audioDir = Join-Path $fixture ('M' + [char]0x00fc + 'sik [mix]')
  New-Item -ItemType Directory -Path $audioDir | Out-Null
  $script = Join-Path $fixture 'app\LUFS-Lens.ps1'
  Copy-Item -LiteralPath (Join-Path $appRoot 'app\LUFS-Lens.ps1') -Destination $script
  $env:PATH = (Join-Path $appRoot 'ffmpeg\bin') + ';' + $env:PATH
  $ascii = Join-Path $audioDir 'plain.wav'
  & $ffmpegExe -hide_banner -loglevel error -f lavfi -i 'sine=frequency=1000:duration=3:sample_rate=48000' -y $ascii
  if ($LASTEXITCODE -ne 0) { throw 'Fixture generation failed' }
  $accented = Join-Path $audioDir ('Gr' + [char]0x00fc + 'sse caf' + [char]0x00e9 + ' [final].wav')
  $nonLatin = Join-Path $audioDir ([string][char]0x97f3 + [char]0x697d + '.wav')
  Copy-Item -LiteralPath $ascii -Destination $accented
  Copy-Item -LiteralPath $ascii -Destination $nonLatin
  $inputs = @($ascii, $accented, $nonLatin)
  $list = Join-Path $fixture 'selected.txt'
  # Match the launcher's file-picker serialization exactly.
  $inputs | Set-Content -LiteralPath $list -Encoding UTF8
  $baseline = $null
  foreach ($mode in @('list', 'direct', 'folder')) {
    $paths = switch ($mode) {
      'list' { '@' + $list }
      'direct' { $inputs }
      'folder' { $audioDir }
    }
    $savedTarget = if ($mode -eq 'list') { -1 } else { -14 }
    $settingsPath = Join-Path $fixture 'settings.txt'
    Set-Content -LiteralPath $settingsPath -Value "TargetLUFS=$savedTarget"
    $global:lufsTestTargetAnswer = if ($mode -eq 'folder') { '-1' } else { '' }
    if ($mode -eq 'direct') { & $script -TargetLUFS -1 -Paths $paths }
    else { & $script -PromptForTarget -Paths $paths }
    if ((Get-Content -LiteralPath $settingsPath) -ne "TargetLUFS=$savedTarget") { throw 'Per-run override rewrote the saved default' }
    $csv = Get-ChildItem -LiteralPath (Join-Path $fixture 'Reports') -Filter '*.csv' | Sort-Object LastWriteTime | Select-Object -Last 1
    $rows = @(Import-Csv -LiteralPath $csv.FullName)
    if ($rows.Count -ne 3) { throw "$mode did not analyze all filenames" }
    foreach ($row in $rows) {
      if ($row.Path -notin $inputs -or $row.File -ne [IO.Path]::GetFileName($row.Path)) { throw "$mode changed a filename" }
      if ($row.Status -eq 'ERROR') { throw "$mode analysis error" }
      if ($row.TargetLUFS -ne '-1') { throw "$mode target precedence failed" }
      if ($row.Issues -notmatch 'GAIN EXCEEDS PEAK HEADROOM' -or $row.GainAdvice -notmatch 'Gain alone cannot meet both limits') { throw "$mode missing gain warning" }
      if (-not $baseline) { $baseline = $row }
      foreach ($metric in @('IntegratedLUFS', 'TruePeak_dBTP', 'SamplePeak_dBFS', 'LRA', 'SuggestedGain_dB')) {
        if ($row.$metric -eq '' -or $row.$metric -ne $baseline.$metric) { throw "$mode changed $metric" }
      }
    }
    $html = [System.Net.WebUtility]::HtmlDecode((Get-Content -Raw -LiteralPath ([IO.Path]::ChangeExtension($csv.FullName, '.html'))))
    if (-not $html.Contains('Gain alone cannot meet both limits')) { throw "$mode lost HTML gain warning" }
    foreach ($path in $inputs) {
      if (-not $html.Contains([IO.Path]::GetFileName($path))) { throw "$mode lost HTML filename" }
    }
  }
  if (@($global:lufsProgressStages | Where-Object { $_ -like '*pass 1 of 2*' }).Count -ne 9 -or
      @($global:lufsProgressStages | Where-Object { $_ -like '*pass 2 of 2*' }).Count -ne 9 -or
      @($global:lufsProgressStages | Where-Object { $_ -eq 'Completed' }).Count -ne 3) { throw 'Progress stages missing' }
  $content = Get-Content -Raw -LiteralPath $script
  if ($content -match 'Start-Sleep|Initializing loudness|Calibrating peak|Preparing loudness') { throw 'Startup delays remain' }
  Write-Output 'PASS: Windows PowerShell; ASCII, accents, non-Latin names, spaces and brackets; picker list, direct paths and folder input; identical measurements; CSV/HTML names; both progress stages; startup delays removed.'
} finally {
  $env:PATH = $originalPath
  # Only remove the unique fixture directory created by this test.
  $resolvedFixture = [IO.Path]::GetFullPath($fixture)
  $testRoot = [IO.Path]::GetFullPath($PSScriptRoot) + [IO.Path]::DirectorySeparatorChar
  if (-not $resolvedFixture.StartsWith($testRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture outside test directory' }
  if (Test-Path -LiteralPath $resolvedFixture) { Remove-Item -LiteralPath $resolvedFixture -Recurse -Force }
}
