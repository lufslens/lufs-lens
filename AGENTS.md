# LUFS Lens maintenance

Existing public Windows utility; current phase is Learn after the v1.2.0 release. Treat release ownership and public distribution as human-controlled. Canonical purpose, scope, architecture, data flow and user guidance are in README.md and README.html. Current release gates, verification, risk assessment and operations are in docs/release-1.2.0.md.

Preserve both measurement passes unless a separate change is authorized and supported by equivalent measurements and a demonstrated benefit. Never change source audio. Keep saved-default and per-run target behavior compatible.

Run tests/Test-GainAdvice.ps1 and tests/Test-Usability.ps1 with Windows PowerShell and FFmpeg/ffprobe available locally or on PATH. Do not commit reports, selected-file lists, local paths, personal audio, credentials, generated fixtures or FFmpeg binaries. Preserve the existing MIT license. Release archives must use an explicit file allowlist.

Before publishing, resume the release record and verify its gate evidence; do not restart intake. Reddit is deferred until separate user instructions.
