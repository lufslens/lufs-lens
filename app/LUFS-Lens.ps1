param(
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$Paths
)

# -----------------------------
# Resolve Tool Root (because PS1 lives in .\app\ now)
# -----------------------------
# $PSScriptRoot == ...\LUFS-Lens\app
# Tool root is parent folder: ...\LUFS-Lens
$ToolRoot = Split-Path -Parent $PSScriptRoot

# -----------------------------
# Settings
# -----------------------------
$SupportedExts       = @(".wav", ".flac", ".aif", ".aiff", ".mp3", ".m4a")
$RecurseFolders      = $true

# Targets / tolerances
$TargetLUFS          = -14.0
$LUFSTolerance       = 0.5            # READY if within +/- this many LU
$MaxTruePeak         = -1.0           # READY if TP <= this (dBTP)
$AllowedSampleRates  = @(44100, 48000)

# Debug behavior
$WriteRawLoudnormDumpOnFailure = $true  # writes debug_* files only if parsing fails
$ShowDebugLines               = $false  # set true if you want loudnorm debug prints

# -----------------------------
# Splash Banner (v1.1)
# -----------------------------
Clear-Host

$bannerColor = "Magenta"
$lineColor   = "DarkMagenta"
$version     = "1.1"

Write-Host ""
Write-Host "==================================================" -ForegroundColor $lineColor
Write-Host "                    LUFS Lens                     " -ForegroundColor $bannerColor
Write-Host ("                    Version {0}                    " -f $version) -ForegroundColor $bannerColor
Write-Host "==================================================" -ForegroundColor $lineColor
Write-Host ""
Write-Host "Independent loudness analysis utility."
Write-Host "Because it sounded louder in the studio."
Write-Host ""
Write-Host "Contact: lufslens@gmail.com"
Write-Host ""
Write-Host "==================================================" -ForegroundColor $lineColor
Write-Host ""

Write-Host "Initializing loudness inspection..." -ForegroundColor Yellow
Start-Sleep -Milliseconds 300
Write-Host "Calibrating peak detectors..." -ForegroundColor Yellow
Start-Sleep -Milliseconds 300
Write-Host "Preparing loudness verdict..." -ForegroundColor Yellow
Start-Sleep -Milliseconds 300
Write-Host ""

# -----------------------------
# Helpers
# -----------------------------
function Get-AudioFilesFromPath([string]$p) {
  if (-not $p) { return @() }

  # Trim surrounding quotes and whitespace (common from BAT/drag&drop)
  $p = "$p".Trim().Trim('"')

  if (-not (Test-Path -LiteralPath $p)) { return @() }

  if (Test-Path -LiteralPath $p -PathType Container) {
    $opts = @{ Path = $p; File = $true }
    if ($RecurseFolders) { $opts.Recurse = $true }
    return Get-ChildItem @opts | Where-Object { $SupportedExts -contains $_.Extension.ToLower() }
  }

  if (Test-Path -LiteralPath $p -PathType Leaf) {
    $item = Get-Item -LiteralPath $p
    if ($SupportedExts -contains $item.Extension.ToLower()) { return @($item) }
  }

  return @()
}

function Format-Duration([double]$seconds) {
  if ($null -eq $seconds -or $seconds -le 0) { return "" }
  $ts = [TimeSpan]::FromSeconds($seconds)
  return "{0}:{1:D2}" -f [int]$ts.TotalMinutes, $ts.Seconds
}

# --- Embed logo as base64 so HTML is portable ---
function Get-Base64ImageDataUri {
  param(
    [Parameter(Mandatory = $true)][string]$ImagePath
  )

  if (-not (Test-Path -LiteralPath $ImagePath)) { return $null }

  $ext = ([IO.Path]::GetExtension($ImagePath)).ToLowerInvariant()
  $mime =
    switch ($ext) {
      ".png"  { "image/png" }
      ".jpg"  { "image/jpeg" }
      ".jpeg" { "image/jpeg" }
      ".gif"  { "image/gif" }
      ".svg"  { "image/svg+xml" }
      default { "application/octet-stream" }
    }

  try {
    $bytes = [IO.File]::ReadAllBytes($ImagePath)
    $b64   = [Convert]::ToBase64String($bytes)
    return "data:$mime;base64,$b64"
  } catch {
    return $null
  }
}

function Normalize-InputPaths([string[]]$InPaths) {
  # Returns a cleaned list of paths:
  # - supports @listfile.txt
  # - supports pipe-joined list in a single argument: "a|b|c"
  # - trims quotes/whitespace
  $out = New-Object System.Collections.Generic.List[string]

  if (-not $InPaths -or $InPaths.Count -eq 0) {
    return @()
  }

  # If a single arg begins with @, treat it as a list file
  if ($InPaths.Count -eq 1) {
    $first = "$($InPaths[0])".Trim()
    if ($first -like '@*') {
      $listFile = $first.Substring(1).Trim().Trim('"')
      if (Test-Path -LiteralPath $listFile) {
        $lines = Get-Content -LiteralPath $listFile | Where-Object { $_ -and $_.Trim() -ne "" }
        foreach ($ln in $lines) {
          $s = "$ln".Trim().Trim('"')
          if ($s) { $out.Add($s) }
        }
        return $out.ToArray()
      }
    }
  }

  foreach ($p in $InPaths) {
    if ($null -eq $p) { continue }
    $s = "$p".Trim().Trim('"')
    if (-not $s) { continue }

    # If the arg contains pipes, split into multiple paths
    if ($s -like '*|*') {
      foreach ($part in ($s -split '\|')) {
        $t = "$part".Trim().Trim('"')
        if ($t) { $out.Add($t) }
      }
    } else {
      $out.Add($s)
    }
  }

  return $out.ToArray()
}

# -----------------------------
# Locate ffmpeg/ffprobe (bundled preferred)
# -----------------------------
$ffmpeg  = Join-Path $ToolRoot "ffmpeg\bin\ffmpeg.exe"
$ffprobe = Join-Path $ToolRoot "ffmpeg\bin\ffprobe.exe"

if (-not (Test-Path $ffmpeg)) {
  $cmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($cmd) { $ffmpeg = $cmd.Source }
}
if (-not (Test-Path $ffprobe)) {
  $cmd = Get-Command ffprobe -ErrorAction SilentlyContinue
  if ($cmd) { $ffprobe = $cmd.Source }
}

if (-not (Test-Path $ffmpeg) -or -not (Test-Path $ffprobe)) {
  Write-Host "ERROR: ffmpeg/ffprobe not found." -ForegroundColor Red
  Write-Host "Fix: bundle them in .\ffmpeg\bin\ OR install FFmpeg and ensure ffmpeg/ffprobe are in PATH."
  exit 1
}

# -----------------------------
# Input paths
# -----------------------------
if (-not $Paths -or $Paths.Count -eq 0) {
  $Paths = @((Get-Location).Path)
}

$Paths = Normalize-InputPaths $Paths

# -----------------------------
# Collect files
# -----------------------------
$files = @()
foreach ($p in $Paths) { $files += Get-AudioFilesFromPath $p }
$files = $files | Sort-Object FullName -Unique

if (-not $files -or $files.Count -eq 0) {
  Write-Host "No supported audio files found." -ForegroundColor Yellow
  Write-Host "Supported: WAV, FLAC, AIFF, MP3, M4A"
  exit 0
}

# -----------------------------
# Output / temp folders (always inside tool root)
# -----------------------------
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$reportsDir = Join-Path $ToolRoot "Reports"
$tempDir    = Join-Path $ToolRoot "temp"

foreach ($d in @($reportsDir, $tempDir)) {
  if (-not (Test-Path -LiteralPath $d)) {
    New-Item -ItemType Directory -Path $d | Out-Null
  }
}

$outCsv  = Join-Path $reportsDir "loudness_report_$timestamp.csv"
$outHtml = Join-Path $reportsDir "loudness_report_$timestamp.html"

Write-Host "Analyzing $($files.Count) file(s)..."
Write-Host "Output CSV: $outCsv`n"

# -----------------------------
# Main loop
# -----------------------------
$results = foreach ($file in $files) {
  Write-Host "Analyzing $($file.Name)..."

  # ---- ffprobe: duration, SR, bit depth, channels, codec, bitrate ----
  $durationSec = $null
  $sampleRate  = $null
  $bitDepth    = $null
  $channels    = $null
  $codec       = $null
  $bitrateKbps = $null

  try {
    $probeJson = & $ffprobe -v error -select_streams a:0 `
      -show_entries format=duration,bit_rate `
      -show_entries stream=sample_rate,bits_per_sample,bits_per_raw_sample,channels,codec_name,bit_rate `
      -of json "$($file.FullName)" 2>$null

    $probe = $probeJson | ConvertFrom-Json

    $dText = [string]$probe.format.duration
    if ($dText -match '(-?\d+(\.\d+)?)') { $durationSec = [double]$Matches[1] }

    $stream = $probe.streams | Select-Object -First 1
    if ($stream) {
      if ($stream.sample_rate) { $sampleRate = [int]$stream.sample_rate }

      if ($stream.bits_per_sample) { $bitDepth = [int]$stream.bits_per_sample }
      elseif ($stream.bits_per_raw_sample) { $bitDepth = [int]$stream.bits_per_raw_sample }

      if ($stream.channels) { $channels = [int]$stream.channels }
      $codec = $stream.codec_name

      $br = $null
      if ($stream.bit_rate) { $br = [int64]$stream.bit_rate }
      elseif ($probe.format.bit_rate) { $br = [int64]$probe.format.bit_rate }

      if ($br -and $br -gt 0) { $bitrateKbps = [math]::Round($br / 1000) }
    }
  } catch { }

  $durationStr = if ($durationSec) { Format-Duration $durationSec } else { "" }

  # ---- loudnorm: Integrated LUFS, True Peak, LRA ----
  $I = $null; $TP = $null; $LRA = $null

  $targetStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.###}", $TargetLUFS)
  $tpStr     = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.###}", $MaxTruePeak)

  $loudnormFilter = "loudnorm=I=${targetStr}:TP=${tpStr}:LRA=8:print_format=json"
  if ($ShowDebugLines) { Write-Host "LOUDNORM FILTER: $loudnormFilter" }

  $oldEap = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  $logLN = & $ffmpeg -hide_banner -i "$($file.FullName)" `
    -af "$loudnormFilter" `
    -f null - 2>&1
  $ErrorActionPreference = $oldEap
  $logLN = $logLN | ForEach-Object { $_.ToString() }

  $idxHit = ($logLN | Select-String -Pattern '"input_i"\s*:' | Select-Object -First 1)

  if ($idxHit) {
    $i0 = $idxHit.LineNumber - 1

    $start = $i0
    while ($start -ge 0 -and ($logLN[$start] -notmatch '\{')) { $start-- }

    $end = $i0
    while ($end -lt $logLN.Count -and ($logLN[$end] -notmatch '\}')) { $end++ }

    if ($start -ge 0 -and $end -lt $logLN.Count) {
      $jsonText = ($logLN[$start..$end] -join "`n")
      try {
        $json = $jsonText | ConvertFrom-Json
        $I   = [double]$json.input_i
        $TP  = [double]$json.input_tp
        $LRA = [double]$json.input_lra
      } catch {
        if ($WriteRawLoudnormDumpOnFailure) {
          $dbg1 = Join-Path $reportsDir "debug_RAW_loudnorm_$($file.BaseName)_$timestamp.txt"
          $dbg2 = Join-Path $reportsDir "debug_JSON_loudnorm_$($file.BaseName)_$timestamp.txt"
          ($logLN -join "`n") | Out-File -Encoding UTF8 $dbg1
          $jsonText | Out-File -Encoding UTF8 $dbg2
        }
      }
    }
  } else {
    if ($WriteRawLoudnormDumpOnFailure) {
      $dbg1 = Join-Path $reportsDir "debug_RAW_loudnorm_$($file.BaseName)_$timestamp.txt"
      ($logLN -join "`n") | Out-File -Encoding UTF8 $dbg1
    }
  }

  $suggestedGain = if ($null -ne $I) { [math]::Round(($TargetLUFS - $I), 2) } else { $null }

  # ---- sample peak (dBFS) via astats (overall, full-file) ----
  $samplePeak = $null

  $oldEap = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'

  $logPeak = & $ffmpeg -hide_banner -i "$($file.FullName)" `
    -af "astats=metadata=1:reset=0,ametadata=print" `
    -f null - 2>&1

  $ErrorActionPreference = $oldEap
  $logPeak = $logPeak | ForEach-Object { $_.ToString() }

  $peakVals = New-Object System.Collections.Generic.List[double]
  foreach ($line in $logPeak) {
    if ($line -match 'lavfi\.astats\.Overall\.Peak_level=([-]?\d+(\.\d+)?)') {
      $peakVals.Add([double]$Matches[1])
    }
  }
  if ($peakVals.Count -gt 0) {
    $samplePeak = [math]::Round(($peakVals | Measure-Object -Maximum).Maximum, 2)
  }

  # ---- Status + Issues ----
  $issues = @()

  $analysisOk = ($null -ne $I -and $null -ne $TP -and $null -ne $LRA)
  if (-not $analysisOk) { $issues += "ANALYSIS ERROR" }

  if ($null -ne $TP -and $TP -gt $MaxTruePeak) { $issues += "TRUE PEAK HOT" }

  $srOk = ($null -eq $sampleRate -or ($AllowedSampleRates -contains $sampleRate))
  if (-not $srOk) { $issues += "SAMPLE RATE CHECK" }

  $withinLUFS = ($null -ne $I -and [math]::Abs($I - $TargetLUFS) -le $LUFSTolerance)
  if ($null -ne $I) {
    if ($I -gt ($TargetLUFS + $LUFSTolerance)) { $issues += "LUFS HIGH" }
    elseif ($I -lt ($TargetLUFS - $LUFSTolerance)) { $issues += "LUFS LOW" }
  }

  $safeTP = ($null -ne $TP -and $TP -le $MaxTruePeak)

  $status =
    if (-not $analysisOk) { "ERROR" }
    elseif ($withinLUFS -and $safeTP -and $srOk) { "READY" }
    else { "ADJUST" }

  $issuesText = if ($issues.Count -gt 0) { $issues -join "|" } else { "NONE" }

  [PSCustomObject]@{
    File             = $file.Name
    Duration         = $durationStr
    SampleRate_Hz    = $sampleRate
    Bitrate_kbps     = $bitrateKbps
    BitDepth         = $bitDepth
    Channels         = $channels
    Codec            = $codec

    IntegratedLUFS   = $I
    SuggestedGain_dB = $suggestedGain
    TruePeak_dBTP    = $TP
    SamplePeak_dBFS  = $samplePeak
    LRA              = $LRA

    Status           = $status
    Issues           = $issuesText
    Path             = $file.FullName
  }
}

# -----------------------------
# Write CSV
# -----------------------------
$results | Sort-Object File | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nDone. CSV saved to: $outCsv"

# -----------------------------
# Summary (whitespace-proof)
# -----------------------------
$total      = @($results).Count
$ready      = @(@($results) | Where-Object { ("$($_.Status)".Trim().ToUpper()) -eq "READY" }).Count
$adjust     = @(@($results) | Where-Object { ("$($_.Status)".Trim().ToUpper()) -eq "ADJUST" }).Count
$errorCount = @(@($results) | Where-Object { ("$($_.Status)".Trim().ToUpper()) -eq "ERROR" }).Count

$avgLUFS = ($results | Where-Object { $_.IntegratedLUFS -ne $null } | Measure-Object IntegratedLUFS -Average).Average
$avgLRA  = ($results | Where-Object { $_.LRA -ne $null } | Measure-Object LRA -Average).Average

Write-Host ("Summary: {0} file(s) | READY: {1} | ADJUST: {2} | ERROR: {3}" -f $total, $ready, $adjust, $errorCount)
if ($null -ne $avgLUFS) { Write-Host ("Average Integrated LUFS: {0:N2}" -f $avgLUFS) }
if ($null -ne $avgLRA)  { Write-Host ("Average LRA: {0:N2}" -f $avgLRA) }

# -----------------------------
# Write HTML report (robust)
# -----------------------------
$css = @"
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
  h1 { margin: 0 0 6px 0; }
  .meta { color: #666; margin: 0 0 16px 0; }
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { border: 1px solid #ddd; padding: 8px; vertical-align: top; }
  th { background: #f3f3f3; font-weight: 700; position: sticky; top: 0; }
  tr:nth-child(even) { background: #fafafa; }
  td.READY  { background: #e9f7ef; font-weight: 700; }
  td.ADJUST { background: #fff4e5; font-weight: 700; }
  td.ERROR  { background: #fdecea; font-weight: 700; }
  .mono { font-family: Consolas, monospace; font-size: 12px; }
  .donate { display:inline-block; padding:8px 10px; border:1px solid #ddd; border-radius:8px; text-decoration:none; color:#333; }
  .donate:hover { background:#f6f6f6; }
  .brand { display:flex; align-items:center; gap:12px; margin-bottom:10px; }
  .brand img { height:64px; width:auto; }

  .notice {
    background: #eef4ff;
    border: 1px solid #c7d8ff;
    padding: 12px 14px;
    border-radius: 8px;
    margin-bottom: 16px;
    font-size: 13px;
  }
  .notice a { font-family: Consolas, monospace; }

  .legend {
    margin-top: 18px;
    padding: 12px 14px;
    border: 1px solid #e6e6e6;
    background: #fafafa;
    border-radius: 8px;
    font-size: 13px;
  }
  .legend h3 { margin: 0 0 8px 0; font-size: 14px; }
  .legend ul { margin: 0; padding-left: 18px; }
  .legend li { margin: 4px 0; }
  .legend code { background: #f0f0f0; padding: 2px 6px; border-radius: 4px; }
</style>
"@

# Build clean objects (NO embedded HTML)
$rows = @(
  $results | Sort-Object File | ForEach-Object {
    $fileUri = "file:///" + ($_.Path -replace '\\','/')
    [PSCustomObject]@{
      File             = $_.File
      Duration         = $_.Duration
      SampleRate_Hz    = $_.SampleRate_Hz
      Bitrate_kbps     = $_.Bitrate_kbps
      BitDepth         = $_.BitDepth
      Channels         = $_.Channels
      Codec            = $_.Codec
      IntegratedLUFS   = $_.IntegratedLUFS
      SuggestedGain_dB = $_.SuggestedGain_dB
      TruePeak_dBTP    = $_.TruePeak_dBTP
      SamplePeak_dBFS  = $_.SamplePeak_dBFS
      LRA              = $_.LRA
      Status           = $_.Status
      Issues           = $_.Issues
      Path             = $fileUri
    }
  }
)

$htmlTable = $rows | ConvertTo-Html -Fragment

# Color Status cells
$htmlTable = $htmlTable `
  -replace '<td>READY</td>',  '<td class="READY">READY</td>' `
  -replace '<td>ADJUST</td>', '<td class="ADJUST">ADJUST</td>' `
  -replace '<td>ERROR</td>',  '<td class="ERROR">ERROR</td>'

# Make Path clickable (file:///...)
$htmlTable = [regex]::Replace(
  $htmlTable,
  '<td>(file:///[^<]+)</td>',
  { param($m)
    $u = $m.Groups[1].Value
    "<td><a class='mono' href='$u'>$u</a></td>"
  }
)

# -----------------------------
# Footer + quotes
# -----------------------------
$limQuotes = @(
  "Blessed are the quiet, for they shall inherit the headroom.",
  "In limiter we trust, but in LUFS we verify.",
  "Thou shalt not clip.",
  "If it's red, it's dead.",
  "Pride comes before -0.1 dBTP.",
  "Ask not what your limiter can do for you. Ask what you did to your transients.",
  "Peak performance requires peak restraint.",
  "All roads lead to -14 LUFS."
)

$randomQuote   = Get-Random -InputObject $limQuotes
$kofiDonateUrl = "https://ko-fi.com/lufslens"

$footer = @"
<hr style='margin-top:30px; opacity:0.2;'>
<p style='font-size:12px; color:#666; margin-bottom:10px;'>
  $randomQuote<br>
  Report generated by <b>LUFS Lens</b>.<br>
  Questions, feedback, or suspiciously loud masters:
  <a href='mailto:lufslens@gmail.com'>lufslens@gmail.com</a>
</p>
<p style='font-size:12px; color:#666;'>
  If this saved your mix (or your reputation), you can
  <a class='donate' href='$kofiDonateUrl' target='_blank' rel='noopener'>buy the limiter a coffee</a>.
</p>
"@

# -----------------------------
# Header with embedded logo (portable)
# -----------------------------
$logoPath    = Join-Path $ToolRoot "assets\logo\LufsLensLogo.png"
$logoDataUri = Get-Base64ImageDataUri -ImagePath $logoPath

$logoHtml = ""
if ($logoDataUri) {
  $logoHtml = "<img src='$logoDataUri' alt='LUFS Lens Logo'>"
}

$reportHeader = @"
<div class='brand'>
  $logoHtml
  <div>
    <h1>LUFS Lens - Loudness Report (v$version)</h1>
    <p class="meta">Generated: $(Get-Date) | Files: $($results.Count) | Target: $TargetLUFS LUFS | TP max: $MaxTruePeak dBTP</p>
  </div>
</div>
"@

# -----------------------------
# CSV Notice Box
# -----------------------------
$csvUri = "file:///" + ($outCsv -replace '\\','/')

$csvNotice = @"
<div class='notice'>
  <strong>CSV report also saved.</strong><br>
  <a href='$csvUri'>$outCsv</a>
</div>
"@

# -----------------------------
# Legend
# -----------------------------
$legend = @"
<div class='legend'>
  <h3>Legend</h3>
  <ul>
    <li><b>Integrated LUFS</b>: average loudness of the whole track. Target here: <code>$TargetLUFS LUFS</code> (tolerance +/-<code>$LUFSTolerance</code>).</li>
    <li><b>True Peak (dBTP)</b>: catches intersample peaks. Limit here: <code>$MaxTruePeak dBTP</code>.</li>
    <li><b>Sample Peak (dBFS)</b>: highest raw digital sample. Useful for spotting hard digital clipping; true peak is the stricter safety check.</li>
    <li><b>LRA</b> (Loudness Range): how dynamic the track is. Low LRA often means heavy compression/limiting; high LRA means more dynamics.</li>
    <li><b>Quick read</b>: <b>READY</b> = within LUFS tolerance + safe true peak + expected sample rate. <b>ADJUST</b> = check the Issues column.</li>
  </ul>
</div>
"@

$fullHtml = "<html><head><meta charset='utf-8'>$css</head><body>$reportHeader$csvNotice$htmlTable$legend$footer</body></html>"
$fullHtml | Out-File -Encoding UTF8 $outHtml

Write-Host "HTML report saved to: $outHtml"

# -----------------------------
# Auto-open HTML report
# -----------------------------
try {
  Start-Process -FilePath $outHtml
} catch {
  Write-Host "Could not auto-open report." -ForegroundColor Yellow
}
