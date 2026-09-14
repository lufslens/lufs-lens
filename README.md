# LUFS Lens 1.2.0

Analyze audio against your own loudness target and get local HTML and CSV reports. LUFS Lens measures integrated LUFS, true peak, sample peak, loudness range (LRA), and technical metadata. It does not modify, normalize, upload, or export audio.

![LUFS Lens](assets/Logo/LufsLensLogo.png)

## Start here

1. Download the app-update ZIP from [Releases](https://github.com/lufslens/lufs-lens/releases) and extract it into a **new folder**, for example `LUFS Lens 1.2.0`.
2. Copy the entire **`ffmpeg` folder** from your old installation into the new folder, beside `LUFS Lens.bat`. If this is your first installation, follow Setup below.
3. Double-click **LUFS Lens.bat**, select your audio, then enter a target or press **Enter** for the displayed default.
4. Follow the track count and analysis stages. The HTML report opens when the batch finishes; both reports are saved in **Reports** beside the launcher.

You can also drag files or folders onto **LUFS Lens.bat**. Folders include supported files in subfolders. Supported extensions: **WAV, FLAC, AIF, AIFF, MP3, M4A**.

Open [README.html](README.html) locally for the full user guide, examples and troubleshooting.

## Setup and updates

Use 64-bit Windows with Windows PowerShell and FFmpeg/ffprobe. Keep the app in a folder where you can save files. Administrator privileges are not normally needed.

**Try the new version in its own folder:** extract the new ZIP into an empty folder, then copy (not move) the entire `ffmpeg` folder from your old installation into it. Put it beside the new `LUFS Lens.bat`, then double-click that launcher. Your old installation and reports stay untouched, so you can return to them whenever you like. New reports will be saved in the new folder's `Reports` directory.

To keep a preferred default, edit the new `settings.txt`, or optionally copy your previous customized settings file into the new folder. Otherwise the initial default is -14 LUFS. You do not need to copy old reports or replace any files in your old installation.

**The v1.2.0 app-update ZIP does not include FFmpeg binaries.** Existing bundled installations can keep their current `ffmpeg` folder. For a fresh installation, obtain a Windows build through the [official FFmpeg download page](https://www.ffmpeg.org/download.html#build-windows) and place both executables here:

```text
LUFS Lens.bat
README.html
settings.txt
app/LUFS-Lens.ps1
assets/Logo/LufsLensLogo.png
ffmpeg/bin/ffmpeg.exe
ffmpeg/bin/ffprobe.exe
```

Alternatively, both programs may be available on PATH. The local `ffmpeg/bin` copies take priority. The tested binary was the Gyan essentials build `2025-12-14-git-3332b2db84`; no FFmpeg upgrade is included in this app update.

## Choose your default target

Open `settings.txt` next to the launcher and edit this line:

```ini
TargetLUFS=-14
```

For example, use `TargetLUFS=-9` for a -9 LUFS default. Any finite numeric value is accepted, including decimals. Decimal points and commas are supported; do not include units or quotes.

- Press Enter at startup to use the saved default.
- Enter another target to override it for that batch only. This does not rewrite the file.
- A missing or invalid settings file falls back to -14 LUFS; invalid content produces a warning.
- An explicit PowerShell `-TargetLUFS` argument takes priority over the settings file.

The initial -14 LUFS value is a comparison reference, not a universal mastering requirement. Tolerance remains **+/-0.5 LU**, the true-peak limit **-1 dBTP**, and expected sample rates **44.1 or 48 kHz**. These other checks are not configurable through `settings.txt`.

## Understand the result

| Result | Meaning |
| --- | --- |
| READY | Within target tolerance, measured true peak within the limit, sample-rate check passes, and target gain fits the measured peak headroom. |
| ADJUST | A loudness, peak, sample-rate or gain-headroom check needs attention. Read Issues and Gain Advice. |
| ERROR | Loudness analysis failed. Check the file and any debug output. |

**Suggested gain** is target LUFS minus measured integrated LUFS. **Gain Advice** flags when that gain would exceed the peak limit. For example, a track at -18 LUFS and -2 dBTP needs +4 dB to reach -14 LUFS, but only +1 dB is available before the -1 dBTP limit. The report shows `GAIN EXCEEDS PEAK HEADROOM`; gain alone cannot meet both limits.

This estimate uses rounded measurements and assumes a simple gain change. Choose a lower target or adjust the master, then remeasure the processed or encoded export. READY is not a platform approval or a guarantee of audio quality.

## Progress and reports

The console shows the track count, current measurement pass and elapsed time. Elapsed time updates at stage/track checkpoints, not continuously within a pass, and is not an estimate of remaining time.

Both original measurement passes remain: one for LUFS/true peak/LRA, one for sample peak. A comparison found matching values for correctly arranged single-pass candidates on 20 files but no speed improvement in the tested workload. See [verification and release notes](docs/release-1.2.0.md).

Reports remain locally until you delete or share them. They include filenames and full local paths; review them before sharing. `temp` can retain the last selected-file list and can be cleared after the app closes. Analysis itself makes no network requests; opening support or download links uses your browser.

## Troubleshooting and limitations

- **No supported audio found:** check the extension and that the path still exists.
- **FFmpeg/ffprobe not found:** check both executable paths above, or PATH. Extract the whole archive first.
- **Report does not open:** open the newest HTML file in `Reports` manually.
- **Analysis error:** try playing or re-exporting the file. Very short or silent files may not have finite loudness measurements. Parsing failures may create debug text in `Reports`.
- **Slow track:** the active pass may take time. Ctrl+C cancels; a cancelled batch may not produce a final report.
- **Script blocked:** follow the policy on your device or contact its administrator.
- **Multi-stream files:** use ordinary files with one audio stream. Metadata currently comes from the first stream while FFmpeg can choose another stream for measurements.

## PowerShell usage

From the app folder, without the startup prompt:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\app\LUFS-Lens.ps1" -TargetLUFS -16 -Paths "C:\Music\track.wav"
```

Add `-PromptForTarget` to request a per-run override. The normal BAT launcher always enables this prompt.

## Changes and development

See [CHANGELOG.md](CHANGELOG.md). To run the regression checks with FFmpeg and ffprobe available on PATH or in `ffmpeg/bin`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-GainAdvice.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Usability.ps1
```

The workflow is deterministic: launcher -> local file selection -> PowerShell -> local FFmpeg/ffprobe -> local HTML/CSV. It has no AI model, cloud storage, application account, paid API, or background service.

LUFS Lens is [MIT licensed](LICENSE). FFmpeg is a separate project with its own [license terms](https://www.ffmpeg.org/legal.html).

Questions and feedback: [GitHub Issues](https://github.com/lufslens/lufs-lens/issues). You can also [support LUFS Lens](https://ko-fi.com/lufslens).
