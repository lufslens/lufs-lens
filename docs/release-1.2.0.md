# v1.2.0 release record

## State and decision

Phase: Launch. Maturity: Production, reflecting the existing public distribution of a local convenience utility, not a claim of certified metering. This is maintenance of the public v1.0.0 release, not a new application or hosted service.

The repository owner requested updated documentation and publication of the latest version to the existing GitHub repository on 2026-09-14. The owner controls release, rollback, risk acceptance and support. Reddit publication is explicitly deferred.

The release candidate is v1.2.0. It adds a configurable reference target, persistent default, gain-headroom advice and usability fixes. It retains both existing measurement passes. The app-update archive excludes FFmpeg binaries, audio, reports, selected-file lists, generated test fixtures and private/local analysis evidence. Existing FFmpeg can be reused; fresh setup is documented.

## Governance and boundaries

| Dimension | Exposure | Control |
| --- | --- | --- |
| Data | R2: local audio and filenames can be confidential | Audio stays local; no upload or telemetry. Reports/path lists remain user-controlled. Publish source/docs only. |
| Actions | R2: public repository/release writes | Explicit owner authorization; allowlisted contents; previous release retained for rollback. Runtime only reads audio and writes local reports. |
| Audience | R3: public download and uncontrolled inputs | Accountable repository owner controls release. Local user initiates every batch. Document limitations, use ordinary single-stream audio and device policy, retain manual fallback. |
| Criticality | R1: convenience analysis | No automatic processing or mastering decisions; compare with another meter or return to prior version. No compliance guarantee. |
| Cost | R0: local execution | No paid APIs, background loops or hosted service. CPU/disk use is bounded by the selected batch; Ctrl+C cancels. |
| Licensing | R1: standard distribution obligations | Existing MIT license retained. FFmpeg is separate and excluded from the new archive; link its official download/license pages. |

Overall R3 is driven by broad public distribution, not financial, legal or safety authority. Human oversight stays with the existing repository owner; this record is technical evidence, not independent legal/security/standards approval.

G0/G1: established project purpose and bounded improvements authorized in the maintenance conversation. G2: publication to the existing GitHub destination authorized by the owner; private data and third-party binaries excluded. G3: no maturity promotion; this resumes an already public utility. G4: owner authorization recorded; technical checks below passed, with the final archive check enforced by the packaging script before upload. G5: review issues and owner feedback after release.

Architecture, data flow and user-facing behavior are documented in README.md. The tool has no app credentials, model, cloud vendor, database, migration, hosted environment or billable runtime. Maintainer GitHub credentials stay in the local credential manager and are never packaged.

## Verification

Required release checks:

- Windows PowerShell syntax and Get-GainAdvice boundary/invalid-measurement tests.
- Real FFmpeg analysis of synthetic files with ASCII, accented and Japanese names, spaces and brackets, using selection lists, direct paths and folders.
- Saved-default Enter behavior, explicit command-line precedence and per-run prompt override, without rewriting settings.
- HTML/CSV target, gain warning, filename and measurement consistency, and both progress stages.
- Visual inspection of the updated guide and review of setup/update instructions.
- Archive allowlist, source-to-package hashes and absence of reports, audio, secrets, personal paths or executables.

On 2026-09-14 both regression scripts passed against the release checkout under Windows PowerShell. Nine synthetic analyses verified three filename variants through the three input modes, including target precedence and unchanged saved settings. The user guide was opened in a browser and visually inspected: logo, headings, reading width and gain/status explanation rendered correctly. The native file-picker UI itself was not automated; its UTF-8 selection-list handoff was tested. Documentation was reviewed against the shipped code.

The packaging script verifies an explicit file allowlist and hashes every ZIP member against its source. Its manifest and SHA-256 checksum accompany the release. No broad workspace upload is used. The published tag identifies the final source commit; the GitHub release records the publication timestamp.

## Analysis-pass decision

A separate investigation used the bundled Gyan FFmpeg build `2025-12-14-git-3332b2db84`. Two correctly arranged single-pass candidates matched the baseline's emitted LUFS, true peak, LRA, threshold and sample peak on 20 files: WAV/FLAC/AIFF/MP3/AAC, 32–192 kHz, integer/float PCM, mono/stereo/5.1, dynamic audio, silence, short clips, impulses, clipping and an existing demo track. Loudnorm was compared at its emitted two-decimal precision; sample peak at its emitted six-decimal precision.

A separate multi-stream M4A check found that the experimental split graph's explicit first-stream selection differed from the existing automatic selection. Reversing the filter order also measured normalized output instead of source sample peak. Neither arrangement is shipped.

Three rotating-order runs on a deterministic 120-second stereo 48 kHz file gave median subprocess times of 4.160 s (current two passes), 4.409 s (split), and 4.606 s (sample-peak-first chain). These exclude UI, metadata probing and PowerShell report parsing. The test does not prove equivalence on all inputs or certify baseline accuracy. With no demonstrated speed benefit, both original passes remain.

## Known limitations and acceptance

- Gain advice is an estimate from rounded values and assumes uniform gain. Users must remeasure any processed/encoded export.
- READY only describes the app's checks and is not a delivery-platform approval.
- Silent or very short audio may not yield finite loudness measurements.
- Multi-stream metadata and measurement selection may differ; ordinary single-stream audio is the documented workflow.
- Progress updates at stage/track boundaries, not continuously within a pass.
- Reports include local paths. The guide explains retention and checking them before sharing.
- Very rapid or simultaneous runs can collide on second-resolution report filenames or the last-selected-file list. Use one instance at a time.
- This release does not upgrade FFmpeg or assert that arbitrary FFmpeg builds produce identical results.

These are disclosed constraints of the owner-requested convenience-tool release. No known destructive source-audio operation or paid external runtime is introduced. Future measurement changes require separate authorization and comparative evidence.

## Operations and rollback

Support and release owner: repository maintainer. Support entry point: GitHub Issues. Feedback review: after release and before the next version; no automatic Reddit or outbound messaging.

The maintainer can mark a problematic release as a prerelease or remove its download, explain the issue, and direct users to the previous release. Retain v1.0.0 and avoid force-pushing or rewriting history. Users keep a backup of their installation and settings and can restore it. Reports/audio need no migration.

There is no server/on-call service to monitor. Local parsing failures may create debug files; users should inspect filenames/paths before sharing those. The owner triages reported accuracy, startup and filename issues before further distribution. Do not redistribute the bundled local FFmpeg executables as a new release until their corresponding source/license distribution materials are verified.
