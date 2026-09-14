$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'LUFS_Lens_v1\app\LUFS-Lens.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) { $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'app\LUFS-Lens.ps1' }
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$function = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GainAdvice' }, $true)
. ([scriptblock]::Create($function.Extent.Text))
$cases = @(
  @{ I=-18; TP=-2; Target=-14; Limited=$true },
  @{ I=-18; TP=-8; Target=-14; Limited=$false },
  @{ I=-18; TP=-5; Target=-14; Limited=$false },
  @{ I=-14; TP=-1; Target=-14; Limited=$false },
  @{ I=-10; TP=0; Target=-14; Limited=$false },
  @{ I=-13.5; TP=0; Target=-14; Limited=$true },
  @{ I=-14.2; TP=-1.1; Target=-14; Limited=$true }
)
foreach ($case in $cases) {
  $advice = Get-GainAdvice -IntegratedLUFS $case.I -TruePeak $case.TP -Target $case.Target -PeakLimit -1
  if ($advice.Limited -ne $case.Limited) { throw "Wrong gain advice: $($case | Out-String)" }
}
$example = Get-GainAdvice -IntegratedLUFS -18 -TruePeak -2 -Target -14 -PeakLimit -1
if ($example.Message -notmatch '\+4.00 dB' -or $example.Message -notmatch '\+1.00 dB' -or $example.Message -notmatch '2.00 dBTP') { throw 'Example advice text incorrect' }
foreach ($invalid in @($null, [double]::NaN, [double]::NegativeInfinity, [double]::PositiveInfinity)) {
  $advice = Get-GainAdvice -IntegratedLUFS $invalid -TruePeak -2 -Target -14 -PeakLimit -1
  if ($advice.Limited -or $advice.Message -notlike 'Unavailable*') { throw 'Invalid measurement advice incorrect' }
  $advice = Get-GainAdvice -IntegratedLUFS -18 -TruePeak $invalid -Target -14 -PeakLimit -1
  if ($advice.Limited -or $advice.Message -notlike 'Unavailable*') { throw 'Invalid peak advice incorrect' }
}
Write-Output 'PASS: gain warning, safe gain, exact boundary, zero gain, attenuation, insufficient attenuation, near-target peak conflict, missing/non-finite measurements.'
