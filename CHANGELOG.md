# Changelog

## 1.2.0 — 2026-09-14

- Choose a finite LUFS target at startup, including decimal values.
- Set a persistent default in `settings.txt`; override it for one run or via PowerShell.
- Include the selected target in HTML and CSV reports.
- Show Gain Advice when reaching the target would exceed measured true-peak headroom.
- Preserve accented/non-Latin filenames from the file picker and handle brackets literally in folder paths.
- Show track count, measurement stage and elapsed time.
- Remove the artificial initializing/calibrating/preparing messages and 0.9-second startup delay.
- Expand setup, update, reporting and troubleshooting documentation.
- Keep the original two-pass measurement approach after an equivalence and timing investigation.

The new download is an app-update package. It reuses an existing FFmpeg installation and does not redistribute FFmpeg binaries.

## 1.0.0 — 2026-02-12

First public GitHub release: local loudness analysis with HTML/CSV reports and bundled FFmpeg. The recovered script displayed version 1.1 internally; no separate public 1.1 release was found.
