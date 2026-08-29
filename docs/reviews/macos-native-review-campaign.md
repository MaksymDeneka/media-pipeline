# macOS native review campaign

## Frozen review packet

- Base commit: `58717a5fe1e994f51bc2b4606808c89026671a69`
- Changed or new files: 83
- Snapshot: temporary local review packet, removed after the campaign
- Packet SHA-256: `68620345B2450C134BC98F950B9BF7F7FB120BA57DE717D9D02A18864650CB47`
- Review generation 1: correctness, compatibility/security, and operations/packaging
- Reviewer constraint: read-only against the frozen packet

## Release invariants

- macOS runs the bundled native worker and never runs local PowerShell.
- Windows legacy behavior remains structurally compatible except for the accepted manifest correction.
- One worker owns a pipeline root. Legacy and native workers cannot race.
- Lane layout remains `<root>/<workspace>/<preset>` and archive/control paths stay inside the root.
- Uploads resume without mixing content, verify the actual source remotely, and delete locally only after verification.
- INI comments, order, inheritance, and case-insensitive keys remain usable.
- Long operations remain visible in the window and menu bar.
- The macOS 13 arm64 package embeds, signs, and exercises its worker and resources.
- Verification never touches the production pipeline root.

## Consolidated ledger

| ID | Severity | Finding | Source reports | Disposition |
|---|---:|---|---|---|
| MP-001 | Critical | Hardened-runtime signing omits the .NET JIT entitlement | RC-OPS-01 | Closed statically in VG4; signed Mac execution gate remains |
| MP-002 | High | Packaged SwiftPM resource bundle is copied where `Bundle.module` cannot resolve it | RC-CS-01, VG2 compatibility/operations | Closed statically in VG4; packaged Mac execution gate remains |
| MP-003 | High | No initial worker status creates a false-stopped window and unsafe root changes | RC-OPS-02, VG2 correctness | Closed in VG3: one operation gate spans stop, save, root switch, and start |
| MP-004 | High | Same-length stale chunks can replace and delete the current source | COR-001, RC-OPS-03, RC-CS-02, VG2/VG3/VG4 correctness and compatibility | Closed statically in VG7: exact verified claim identity reaches Trash and failure restoration never replaces producer output; Mac execution gate remains |
| MP-005 | High | Concurrent same-name uploads share state and can verify the wrong content | RC-CS-03, VG3/VG4 compatibility | Closed in VG4: lease spans cleanup and its stable inode is never unlinked |
| MP-006 | High | Preset/workspace traversal can escape the pipeline root | COR-002, RC-CS-04 | Closed: configuration validation and containment tests reject escaping segments |
| MP-007 | High | Legacy and native workers use different ownership locks | COR-003, RC-OPS-04, VG2 correctness | Closed in VG3: both sides canonicalize equivalent roots before hashing |
| MP-008 | High | Jobs crossing midnight disappear from Swift activity state | COR-004, RC-OPS-05, VG3/VG4/VG5/VG6/VG7 operations | Closed statically in VG7: initial history precedes the authoritative snapshot and missing/malformed snapshots publish replayed state; Mac execution gate remains |
| MP-009 | High | Retention can recursively delete fresh or active upload parts | RC-OPS-06, VG2/VG3/VG4 correctness/operations | Closed in VG4: persistent sibling lease is held through atomic claim and deletion |
| MP-010 | Medium | Failed preparation/segment operations strand temporary files | COR-008 | Closed with focused integration cleanup checks |
| MP-011 | Medium | Cancelled or crashed jobs remain running in the UI indefinitely | COR-005 | Closed statically: lifecycle events and worker reconciliation cancel stale running state; Mac execution gate remains |
| MP-012 | Medium | Size-cap retries ignore a preset `MaxWidth` override | COR-006, VG2 correctness | Closed in VG3: every fallback is clamped to preset `MaxWidth` |
| MP-013 | Medium | `TimeoutSeconds` is exposed but unused by the native worker | COR-007, RC-OPS-07, VG2/VG3/VG4/VG5/VG6 correctness | Closed in VG7: timed-out paths are never moved from stale pathname state and safely rearm through a fresh recovery window |
| MP-014 | High | Repeated Windows cutover can stop one root and start stale config on another | RC-OPS-08, RC-CS-05 | Closed in VG4 |
| ADJ-001 | Low | Activity parsing/sorting grows for the full day on the main actor | RC-OPS-09 | Partially fixed: retained state bounded; full daily parse remains tracked |
| ADJ-002 | Medium | Fresh-Mac upload setup omits rclone remote and SSH-key preflight | RC-OPS adjacent | Fixed in installer guidance and macOS runbook |
| ADJ-003 | Medium | Swift config edits discard inline comments | RC-CS-06, VG3/VG4 compatibility | Closed statically in VG4; C# parsing now also accepts a quoted value followed by a comment; Mac gate remains |
| ADJ-004 | Medium | Relative `PipelineRoot` resolves differently in UI and worker | RC-CS-07 | Closed statically: UI, worker, cutover, and restore use config-relative canonical resolution; Mac execution gate remains |
| ADJ-005 | Medium | SSH executes an assembly script from writable remote staging | RC-CS-08 | Fixed: local code is encoded; staged manifest bytes are authenticated before parsing |
| ADJ-006 | Medium | Remote chunks are retransmitted after cancellation | RC-CS-09 | Track as pre-existing optimization; not a release blocker |
| VG2-CS-01 | High | A chunk-by-chunk manifest embedded in the encoded SSH command exceeds Windows command-line limits | VG2 compatibility/security | Closed: a constant-size command authenticates a separately uploaded manifest |
| VG2-CS-02 | Medium | Windows cutover reports success without proving the new worker is healthy or restoring legacy on failure | VG2 compatibility/security, VG3/VG4 operations | Closed in VG4 to the bounded mutex-health criterion |
| VG3-OPS-01 | High | PowerShell restore resolves supported roots differently and reports success without health | VG3/VG4 operations | Closed in VG4: canonical resolution plus verified recovery in both directions |
| VG4-CS-01 | Medium | Unawaited upload-output tasks can lose or reorder terminal progress and skip Trash | VG4/VG5 compatibility and correctness | Closed statically in VG7: one locked post-exit owner drains, finishes, awaits, and validates the typed terminal stream; Mac execution gate remains |
| VG6-CS-01 | Medium | Changing `PipelineRoot` retains activity from the previous root | VG6 compatibility/security | Closed statically in VG7: standardized root identity clears all cursor/job state before loading a different root; Mac execution gate remains |
| ADJ-007 | Low | Stable zero-byte upload lease files accumulate by content identity | VG3 design checkpoint | Accepted safety tradeoff; parts are removed and lock files remain reusable |

## Verification generation 2

All three reviewers reread the repaired live tree without changing it. Correctness kept
MP-003, MP-004, MP-007, MP-009, MP-012, and MP-013 open with concrete interleavings or
configuration counterexamples. Compatibility/security kept MP-002 and MP-004 open and found
VG2-CS-01 and VG2-CS-02. Operations independently confirmed MP-002 and MP-009. Where reviewers
disagreed, the root accepted the stricter report because its failure scenario was reproducible
or directly supported by the source.

## Verification generation 3 and design checkpoint

VG3 closed the packaging, lifecycle-gate, canonical-root, resize, manifest, and most traversal
findings. It kept MP-004, MP-005, MP-008, MP-009, MP-013, ADJ-003, and VG2-CS-02 open and added
VG3-OPS-01. Because two consecutive verification generations found material ownership gaps, the
campaign stopped patching for a design checkpoint. The selected model uses persistent lock inodes,
cleanup under ownership, a worker-authored active-job snapshot, platform-native Trash on macOS,
sticky terminal readiness state, and health-verified migration in both directions.

## Verification generation 4

VG4 closed package placement, signing, upload lease, retention, cutover, and rollback findings. It
found that historical activity replay could override an empty authoritative snapshot, that the
macOS worker restored a verified claim to a replaceable pathname before Trash, that terminal
progress delivery was not awaited, and that rejected readiness state could leak to a later file
reusing the same path. The root accepted the stricter reports where reviewers disagreed. Repairs
now overlay the active snapshot after history, retain the exact verified claim for Trash, serialize
and await upload output, restore a failed Trash handoff without replacing new producer output, and
clear missing observations per lane. These changes require VG5.

## Verification generation 5

VG5 closed the exact-claim Trash ownership model but found three stricter lifecycle gaps. Today's
history was still replayed after the authoritative snapshot, the first readiness repair still
relied on a nonportable creation timestamp and eventually moved a pathname, and an EOF callback
could finish the stream before the post-exit drain yielded its terminal bytes. The next repair
replays all initial history before the snapshot, never moves a rejected pathname, rearms only after
a fresh stable/exclusive recovery window, and gives stream completion solely to the locked
post-exit path. A no-trailing-newline mock-worker test now exercises terminal delivery. These
changes require VG6.

## Verification generation 6

VG6 closed the source-ownership, timeout, and terminal-stream findings. Operations found that a
missing or malformed active snapshot skipped the final publish after current-day replay, and
compatibility/security found that an ActivityStore reused across a root change retained the old
root's terminal jobs. The repair publishes replayed history even when no snapshot is available and
resets every activity cursor and job collection when the standardized pipeline root changes. The
existing rollover test and a new root-isolation test cover both cases. These changes require VG7.

## Verification generation 7

All three independent lenses reported the repaired tree clean. MP-008 and VG6-CS-01 closed after
reviewers confirmed that missing or malformed snapshots publish current history, valid snapshots
remain authoritative, date rollover preserves jobs, and root changes clear all prior-root activity.
Spot checks kept MP-004, MP-013, VG4-CS-01, package/signing, upload leases, authenticated assembly,
configuration preservation, and migration/rollback closed. No accepted static finding remains open.

## Verification evidence

This section is updated after every fix and verification generation. A finding is closed only with a focused regression or a direct package/platform assertion.

Current local evidence:

- Core self-test: pass, including equivalent-root mutex identity, source-identity deletion, quoted-comment parsing, replacement-path readiness, authenticated constant-size remote command, and atomic retention claims.
- Integration self-test: pass with FFmpeg/ExifTool, including an inherited wide fallback, sticky locked-input timeout, missing-path identity reset, failed temp cleanup, and initial status liveness.
- The generated remote assembly command executes successfully under Windows PowerShell and verifies the assembled SHA-256.
- Native worker Release build: zero warnings and zero errors.
- Windows tray isolated Release build: zero warnings and zero errors.
- Windows tray upload/config/archive/lane self-tests: pass.
- Legacy watcher smoke test: pass.
- Isolated same-corpus legacy capture versus native compare: parity pass for MP4, MOV, JPG, PNG, HEIC, flat/per-source/per-set grouping, manifests, and segmentation.
- Self-contained publish: pass for `win-x64` and `osx-arm64`.
- PowerShell parser and `git diff --check`: pass.
- Swift source tests now include same-day and prior-day orphan suppression, missing-snapshot
  rollover, root isolation, exact-claim recovery, quoted-comment editing, and twenty terminal
  deliveries without a trailing newline. Execution remains part of the Mac gate below.
- Tray lifecycle UI automation: all steps passed except the synthetic final tray-menu keyboard quit action. The normal test target was already running and the harness is not isolated, so it will not be rerun or treated as macOS release evidence.
- Swift tests, signed app launch, entitlement assertion, and packaged worker execution remain macOS-only and are encoded in `Build-MacApp.sh` and `Test-MacPackage.sh`.

## Release decision

The static campaign is clean and the implementation is ready for Apple-silicon validation. Final
macOS release readiness requires the signed package checks on an Apple-silicon Mac; they cannot be
executed on this Windows host.
