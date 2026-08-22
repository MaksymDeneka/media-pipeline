# Pipeline model

Target design for collapsing the nine hardcoded pipelines into one configurable pipeline.

## Why

The nine pipelines are not nine different processes. They are one process with different
orchestration around it, and the code says so:

- `videoclean` is fifteen lines of logging around `Process-VideoFile -CopyCount 1`.
- `Convert-SetImageVariant` and `Convert-SetBatchImageVariant` are 66-line functions that
  differ only in the function name, one parameter rename, which extension helper they call,
  and three log strings.
- `Convert-VideoVariant` and `Convert-SetVideoVariant` differ only in that the set version
  omits the size cap, and `New-VideoEncoderArguments` already treats a zero bitrate ceiling
  as "no ceiling".
- The randomized-crop block is copy-pasted verbatim four times.
- The batch-settling block is copy-pasted verbatim twice.

The duplication has already caused drift. Three output-extension mappers disagree, so the
same `.heic` file becomes `.png` in the `default` and `sets` lanes and `.jpg` in
`images`, `imageclean` and `setbatch`.

## The model

One pipeline. A **preset** is a named bundle of options. A **workspace** is unchanged: it
routes output by category (`LC`, `MD`, `YL`, `PL`, `general`).

Folder layout is unchanged, which means existing folders keep working with no migration:

```
<PipelineRoot>\<preset>\<workspace>\{input,output,original,failed}
```

### Media type is detected, not configured

The single most important change. A preset carries a **separate copy count per media type**:

```ini
[preset standard]
VideoCopies = 20
ImageCopies = 100
```

That is what makes one inbox viable. Drop a mixed folder of videos and photos into one place
and each file gets the treatment its type calls for, with no lane to choose.

### Options

Every option inherits from the global config sections unless the preset overrides it, so a
typical preset is three lines rather than forty.

| Option | Values | Default | Replaces |
|---|---|---|---|
| `VideoCopies` | integer, 0 disables video | 1 | the per-lane copy-count keys |
| `ImageCopies` | integer, 0 disables images | 1 | same |
| `Grouping` | `Flat`, `PerSource`, `PerSet` | `Flat` | the difference between images, sets and setbatch |
| `SetCount` | integer, only when `Grouping = PerSet` | 1 | `SetBatchCount`, `AssetStoreSetCount` |
| `Batch` | `PerFile`, `PerGroup` | `PerFile` | the setbatch/assetstore settle rule |
| `Segment` | `false`, `true` | `false` | the long lane |
| `Manifest` | `false`, `true` | `false` | the assetstore lane |
| `Normalize` | `false`, `true` | `true` | the convert lane |
| `OnFailure` | `PreservePartial`, `DeleteFiles`, `DeleteContainer` | by grouping | four inconsistent rollback behaviors |
| `Parallel` | `OverFiles`, `OverVariants`, `Sequential` | `OverFiles` | per-lane hand-wiring |

Quality settings (`MaxWidth`, `SizeCapMB`, `Crf`, `NvencCq`, `AmfQp`, `AudioBitrate`,
`JpegQuality`, `PngCompressionLevel`, trim range, crop range) are all overridable per preset
and otherwise inherit the global value.

### How today's nine map onto it

| Today | Preset options |
|---|---|
| `default` | `VideoCopies=20`, `ImageCopies=20` |
| `videoclean` | `VideoCopies=1`, `ImageCopies=0` |
| `imageclean` | `VideoCopies=0`, `ImageCopies=1` |
| `images` | `VideoCopies=0`, `ImageCopies=100` |
| `sets` | `Grouping=PerSource`, copies 10 |
| `setbatch` | `Grouping=PerSet`, `SetCount=10`, `Batch=PerGroup` |
| `assetstore` | as setbatch, plus `Manifest=true` |
| `long` | `Segment=true`, `VideoCopies=3` per segment |
| `convert` | folded into `Normalize`, no longer a preset |

### `convert` stops being a destination

The `convert` lane never encodes anything. It is a stream copy with extension-driven dispatch,
and its pass-through branch breaks the contract every other lane keeps: the input file becomes
the output, so nothing is preserved in `original\`.

The `long` lane already pre-remuxes `.mov` before segmenting, using `convert`'s own helper. So
normalization is better modeled as an **input stage every preset gets**: a `.mov` or `.heic`
dropped into any inbox is normalized first, then processed. The pass-through branch is dropped,
because under `Normalize` an already-supported file simply proceeds to processing.

## Agreed behavior changes

Three deliberate changes, confirmed with the repository owner:

1. **Images get real variation.** The `default` lane currently produces image "variants" by
   calling `Copy-Item` N times, yielding N byte-identical files with different random names.
   Unified, images use the `images` lane technique: a small randomized crop scaled back to the
   original dimensions, so every copy actually differs.
2. **`.jpg` is the standard.** HEIC and PNG-family inputs converge on `.jpg` output everywhere,
   ending the three-way mapper disagreement.
3. **`convert` is absorbed** into the normalization stage as described above.

## Observability

The watcher writes a flat text log with no lane tag, no workspace field, and no job
correlation. With parallel runspaces, per-variant lines from different files interleave with
nothing tying them back to a source. That is unusable for a progress UI.

Alongside the existing log, the watcher gains an append-only event stream at
`<PipelineRoot>\logs\events-YYYYMMDD.jsonl`, one JSON object per line, written through the
same mutex the logger already uses so parallel runspaces stay safe:

```json
{"ts":"2026-08-22T11:20:53.412Z","seq":1841,"ev":"job.variant","preset":"standard","ws":"LC","jobId":"a1b2c3d4","file":"IMG_20260819_154233.heic","n":47,"total":100}
```

Events: `watcher.start`, `watcher.stop`, `watcher.heartbeat`, `job.detected`, `job.start`,
`job.variant`, `job.done`, `job.failed`, `archive`, `retention`.

`jobId` is what makes parallel progress attributable, which the current log cannot do.

## Control

There is no IPC of any kind today. Restarting means `Stop-Process -Force`, which orphans an
in-flight ffmpeg, leaves partial outputs, and strands the input file in `input\` with no
failed-move.

The poll loop gains a check of `<PipelineRoot>\control\` each tick:

| File | Effect |
|---|---|
| `stop` | finish the current file, then exit cleanly and release the mutex |
| `pause` | stop picking up new work, keep running |
| `pause.<preset>` | pause one preset across all workspaces |
| `pause.<preset>.<workspace>` | pause one lane |

Current state is mirrored to `<PipelineRoot>\status\watcher.json` so a UI can read it without
guessing. Retrying a failed file needs no control surface at all: moving it from `failed\`
back to `input\` is already the queue API.

## Verification

`tools\Test-PipelineParity.ps1` fingerprints every lane's output structurally: counts,
extensions, folder depth, pixel dimensions, bucketed durations, and source disposition.
Output names come from a crypto RNG and can never be reproduced, so names are ignored.

Capture a baseline before a change, compare after. The refactor is done in stages, and every
stage that is not one of the three agreed behavior changes must report `PARITY OK`.
