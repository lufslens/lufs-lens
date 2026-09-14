$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Path $dist -Force | Out-Null
$archivePath = Join-Path $dist 'LUFS-Lens_v1.2.0-app-update.zip'
if (Test-Path -LiteralPath $archivePath) { throw 'Release ZIP already exists; use a new clean dist directory or review the existing build.' }
$allowlist = @(
  'LUFS Lens.bat', 'README.html', 'README.md', 'CHANGELOG.md', 'LICENSE', 'settings.txt',
  'app/LUFS-Lens.ps1', 'assets/Logo/LufsLensLogo.png', 'docs/release-1.2.0.md',
  'tests/Test-GainAdvice.ps1', 'tests/Test-Usability.ps1', 'tests/Build-Release.ps1'
)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($archivePath, [IO.Compression.ZipArchiveMode]::Create)
try {
  foreach ($relative in $allowlist) {
    $source = Join-Path $root $relative
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing release file: $relative" }
    [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $source, $relative, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
  }
} finally { $zip.Dispose() }
$zip = [IO.Compression.ZipFile]::OpenRead($archivePath)
$manifest = @()
try {
  if ($zip.Entries.Count -ne $allowlist.Count) { throw 'Unexpected ZIP member count' }
  foreach ($entry in $zip.Entries) {
    if ($entry.FullName -notin $allowlist) { throw "Unexpected ZIP member: $($entry.FullName)" }
    $stream = $entry.Open()
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { $memberHash = [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '').ToLowerInvariant() }
    finally { $stream.Dispose(); $hasher.Dispose() }
    $sourceHash = (Get-FileHash -LiteralPath (Join-Path $root $entry.FullName) -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($memberHash -ne $sourceHash) { throw "ZIP content mismatch: $($entry.FullName)" }
    $manifest += "$memberHash  $($entry.FullName)"
  }
} finally { $zip.Dispose() }
$manifest | Set-Content -LiteralPath (Join-Path $dist 'MANIFEST.sha256') -Encoding Ascii
$archiveHash = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
"$archiveHash  $([IO.Path]::GetFileName($archivePath))" | Set-Content -LiteralPath (Join-Path $dist 'SHA256SUMS.txt') -Encoding Ascii
Write-Output "PASS: $($allowlist.Count) allowlisted files, all ZIP members hash-match their sources."
Write-Output "ZIP SHA256: $archiveHash"
