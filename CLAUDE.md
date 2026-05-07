# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository shape

This repo is **not** a typical app project — there is no `package.json`, no build system, no test runner, no linter. Each `Booking engine v*.html` is a fully self-contained single-page app (HTML + inlined CSS + inlined JS + inlined `ExcelJS` and `pdfjs-dist` libraries). The "dev loop" is opening the HTML file directly in a Chromium-based browser.

Versions are kept **side-by-side**, not overwritten. `Booking engine v5.html` (~4705 lines, ~2.5 MB) is the current canonical app; `Booking engine v4.html` (~4221 lines) and `Booking engine v3.3 (offline).html` are the build pipeline inputs and stay at the repo root because `make-v4.sh` / `make-v5.sh` reference root-relative paths. Pure-history files (`Booking engine v3.2.html`, `V2 Booking engine.html`, `v3-bookings.html`, `Booking sheets unified V1.html`, plus the `Booking engine v4 copy.html` / `Booking engine v4 - Ops preview.html` scratch files) live under `archive/` and should generally not be edited unless explicitly requested. Sample PDFs for the inhouse-meal-plan parser live in `samples/`.

## Building v5 from v4 (and v4 from v3.3)

Each version is built by applying one patch on top of the previous one:

```sh
sh make-v4.sh   # v3.3 → v4: autofill + PDF management
sh make-v5.sh   # v4  → v5: multi-user sync
```

`make-v4.sh` `cp`s `Booking engine v3.3 (offline).html` to `Booking engine v4.html` and applies `patches/0001-autofill-and-pdf-management.patch`. `make-v5.sh` does the same for `Booking engine v4.html` → `Booking engine v5.html` with `patches/0002-multi-user-sync.patch`. Both scripts rewrite the patch's filename paths on the fly. See `patches/README.md` for context — the patches exist because some agent tooling couldn't transmit the 2.5 MB+ files through a single tool call. If you make further changes to v5, edit `Booking engine v5.html` directly and commit it as a binary blob; only re-run `make-v5.sh` when starting from a clean v4. The patch is a snapshot of the original v4→v5 transition; subsequent v5 changes (tab-close safety, etc.) live only in the committed `Booking engine v5.html`.

## Reading the HTML files

The HTML files are too large to `Read` whole-file. Use line offsets (numbers below are for v5):

- Lines 1–10: `<head>`, fonts, meta.
- Lines 11–58: inlined ExcelJS@4.4.0 — **skip**.
- Lines 59–83: inlined pdfjs-dist@3.11.174 (main bundle) — **skip**.
- Lines 84–1056: `<style>` block.
- Lines 1059–1146: `<body>` markup (header with `#status-pill` + `#saving-indicator` + `#btn-save`, welcome card, `#workspace`, `#modal-mount`, `#toast-mount`, hidden `#file-input` / `#pdf-input`).
- Lines 1147–4702: the application `<script>`. The pdf.js **worker** is inlined as a string literal in the first ~50 lines of this script; ignore that block when grepping. After it, top-level `const STATE = { … }` (line 1159) and named sections delimited by `/* === … === */` banners.

Use `grep -n '/\* ============================================================' "Booking engine v5.html"` to jump between sections, then read the line below each banner for the section name. Current section banners (line → name):

| Line | Section |
|------|---------|
| 1148 | Reef Desk — unified booking interface (ExcelJS edition) |
| 1204 | Date helpers |
| 1257 | Cell helpers |
| 1291 | Layout detection |
| 1438 | Booking I/O |
| 1593 | Meta sheets (hidden) — booking-level + sheet-level extras |
| 1812 | File load |
| 1914 | Save / reload |
| 2059 | **Sync (multi-user)** |
| 2414 | File source |
| 2539 | View shell |
| 2596 | Date strip |
| 2645 | Main view dispatch |
| 2658 | Agenda view (selected date) |
| 2983 | People (guests) view |
| 3072 | Search results (across all dates / guests) |
| 3129 | Modals |
| 3689 | Inhouse PDF parsing |
| 4192 | Inhouse reports modal |
| 4308 | Stats dashboard |
| 4596 | Wire-up |

(One additional banner at line 925 lives inside `<style>` — "Operational tracking — status pills, banners, capacity override, saving indicator".)

## Architecture

**Single global `STATE`** (line 1159) holds everything: `mode` (`'fs'` File System Access API vs `'fp'` manual file picker), `view` (`'agenda' | 'people' | 'stats'`), `dirHandle`, `files[]`, `bookings[]`, `excursions[]`, `excFilter`, `search`, `selectedDate` (ISO), `showCancelled`, `inhouseGuests[]`, `inhouseReports[]`, `userId` (operator name, persisted at `localStorage['reefdesk:userId']`), `saving` (Set of filenames currently mid-save). Renders are pull-based: mutate `STATE`, then call `renderView()` / `renderDateStrip()` / `updateStatus()` / `updateSavingIndicator()`.

**One Excel file = one excursion; one sheet = one day.** `detectLayout(ws)` scans the first ~20 rows for a row containing both a "room" and "guest" header (regex map in `HEADER_KEYS`), then walks back up to find the excursion title and capacity. Files whose layout can't be detected are skipped. `~$` lock files and the two hidden meta sheets are filtered out at load time.

**Hidden meta sheets** (`_BookingMeta`, `_SheetMeta`) carry data the original spreadsheet format doesn't have a column for: per-row `status` / `cancelReason` / `attended` / `paid` / postpone log, and per-sheet `activityCancelled` / `customCapacity`. Constants `META_BOOKING_SHEET`, `META_SHEET_SHEET`, `BOOKING_META_COLS`, `SHEET_META_COLS`, `BOOKING_META_DEFAULTS`, `SHEET_META_DEFAULTS` define the schema. `readMetaSheets(wb)` builds the in-memory maps on load; `flushMetaSheets(fileEntry)` writes them back before `wb.xlsx.writeBuffer()` runs in `saveAll()`. Always go through `getBookingMeta` / `setBookingMeta` / `getSheetMeta` / `setSheetMeta` so dirty-tracking stays correct.

**Two file source modes** share the same `loadFiles(specs)` shape:

- `'fs'` — `pickDirectory()` → `showDirectoryPicker({ mode: 'readwrite' })` → `openDirectory(handle)` (line 2416). `saveAll()` writes via `fileHandle.createWritable()`. `openDirectory` also sweeps stale `~$*.savelock.*` files left by crashed tabs and, after `loadFiles`, calls `offerCrashRecovery()` to surface any persisted pendingOps. The "Open shared folder" button hides itself when `showDirectoryPicker` is missing.
- `'fp'` — `pickFiles()` uses `<input type="file" multiple accept=".xlsx">`. `saveAll()` falls back to a download blob (one per dirty file). `reload`, sync, and crash recovery are all no-ops in this mode.

**Inhouse PDF parsing** (TRML / TRMD daily meal-plan reports) uses pdf.js text items with x/y coordinates rather than naive line-extraction. Header detection (`classifyInhouseHeader`) recognises a wide vocab (room/cabin/villa/suite/unit, name/guest/booker, arrival/checkin, departure/checkout, plus neighbours like type/ta/meal/rate so column-range splitting has tight boundaries). Falls back to filename-based date parsing (`parseFilenameDate`) when the report header doesn't carry one — for ambiguous DDMM/MMDD blocks it picks the date closer to today. `lookupInhouse(room, dateISO)` deliberately falls back to the **most recent record for the room** when no date is given or the booking date is outside the guest's stay; this drives the autofill in the New Booking modal even when the loaded reports are stale, and stale/out-of-range cases are surfaced in the modal status hint via `tryAutofill()`.

## Multi-user sync (v5 only)

`saveAll()` in v4 overwrote each .xlsx wholesale, so concurrent saves on the SMB share clobbered each other. v5 adds a `/* === Sync (multi-user) === */` banner at line 2059 (between `Save / reload` and `File source`) that makes saves merge against the current disk version and survive tab close. `'fs'` mode only — `'fp'` is sync-disabled by design. v4 retains the unsafe single-user behaviour.

- **Op-log replay.** `writeBooking` (line 1474), `setBookingMeta` (line 1751), `setSheetMeta` (line 1781) each push a structured op into `f.pendingOps` next to their existing `f.dirty = true`. `bookingTempId` threads a create through later meta ops on the same row (cancel-after-create); `txnId` keeps multi-op flows like `postponeBooking` atomic at replay. A module-level `replaying` flag suppresses `recordOp` and the `rebuildBookings()` call inside `writeBooking`. `replayOps(fileEntry, ops)` is shared by the conflict path (`saveAll` after a remote-mtime change) and the crash-recovery path (`offerCrashRecovery` after reopening the folder).
- **Save flow** (`saveAll`, line 1916). Snapshot `const ops = f.pendingOps.slice()` *before* any await, add `f.name` to `STATE.saving` and call `updateSavingIndicator()`, acquire lock, stat `lastModified`. Equal to `f.lastSyncedMtime` → fast path; else `reloadFile(f)` and replay `ops` grouped by `txnId`, reallocating `intendedRow` when a new booking's row is now occupied. Re-stat **after** `close()`. Splice the snapshotted prefix off `f.pendingOps` — never reassign `[]`, or ops queued during the await are lost — then `persistPendingOps(f)` to clear the localStorage mirror. The `finally` block removes from `STATE.saving`, calls `updateSavingIndicator()`, and releases the lock (also removing from `_heldLocks`).
- **Locks.** `~$<filename>.savelock.<sessionId>` in `STATE.dirHandle`. Two-phase: write candidate, list `~$<filename>.savelock.*`, drop entries whose own `lastModified` is older than 12s (judge by lock-file mtime, not the JSON `ts` — SMB clocks drift), smallest sessionId wins. 10s heartbeat while held; release via `dirHandle.removeEntry(name)`. Held locks are also tracked in a module-level `_heldLocks` Set so the `pagehide` handler can drain them best-effort (fire-and-forget removeEntry) and peers don't wait the full 12s after a clean tab close. The lock is **advisory** — `acquireLock` returning null does not block the save; the mtime check + replayOps is the real safety net. The existing `~$` load filter in `openDirectory` already excludes lock files.
- **Background poll.** 8s `setInterval` while visible (`_pollTimer`); paused on `document.hidden`; skips `STATE.saving` files; guards on `dirHandle.queryPermission({mode:'readwrite'}) === 'granted'` (post-sleep downgrade aborts the tick silently). Remote mtime advanced on a non-dirty file → silent `reloadFile(f)` + "Updated by …" toast; on a dirty file → `f.remoteChanged = true` so the next save forces the conflict branch. View state lives on `STATE`, so re-renders preserve agenda position.
- **`reloadFile(fileEntry)`** (line 1867) — extracted from `loadFiles`'s per-file body, reused by poller and conflict path; does not touch view state.
- **User identity.** `STATE.userId` prompted once via an in-page `el()` modal (not browser `prompt()`), persisted at `localStorage['reefdesk:userId']`, embedded in lock contents and "Updated by …" toasts. `loadUserId` runs at workspace bootstrap; `ensureUserId` prompts only on first run.
- **Crash recovery.** Each `recordOp` (line 2095) mirrors the file's full op queue under `localStorage['reefdesk:pendingOps:<dirHandle.name>:<filename>']` via `persistPendingOps`. The splice-after-save and the Discard path both clear it. `openDirectory` calls `offerCrashRecovery()` (line 2452) after `loadFiles` to surface a Restore/Discard modal listing each affected file and op count; Restore runs `replayOps` against the freshly-loaded workbook and re-marks dirty so a normal Save persists. `'fp'` mode skips persistence (no stable folder identity, so `_opsKey` returns null).
- **Save-in-flight UI.** A `#saving-indicator` "Saving — please don't close" pill in the header (CSS at line ~931) is toggled by `updateSavingIndicator()` (line 2589) on every `STATE.saving.add/delete`. The `beforeunload` handler also blocks on `STATE.saving.size > 0`, not just dirty files — a tab killed mid-write can corrupt the .xlsx, so the prompt is worth showing.

## Conventions

- Don't introduce a build system, framework, or package manager. The whole point is that this app runs by double-clicking the HTML.
- Don't extract helpers into separate files. Keep code in the `<script>` block, grouped under the existing `/* === … === */` section banners.
- DOM construction uses the `el(tag, attrs, ...children)` helper near the top of the script (line 1181, handles `class`, `html`, `style`-as-object, and `on*` listeners). New UI should use it rather than ad-hoc `document.createElement` chains.
- New modals: open with `openModal(node)` and close with `closeModal()`. There is no modal stack — opening a second modal overwrites the first; the userId modal and the recovery modal are intentionally non-concurrent in practice.
- Anything that may need to survive a tab close should round-trip through `localStorage` under the `reefdesk:` key prefix (see `reefdesk:userId`, `reefdesk:pendingOps:…`).
- Never edit the inlined ExcelJS / pdf.js / pdf.worker blocks by hand. If they need updating, replace the whole block from the upstream npm tarball.
- Folder access requires a Chromium browser opened directly (not an iframe). State this in user-facing copy when relevant — the welcome card already does.
- Keyboard shortcuts (defined in the Wire-up section, line 4596): Cmd/Ctrl+S save, Cmd/Ctrl+N new booking, Cmd/Ctrl+K focus search, Esc closes modal or clears search, Left/Right arrows move the selected date. `beforeunload` warns when any file is dirty **or** a save is in flight; `pagehide` drains `_heldLocks` best-effort.

## Git workflow

Working branch for this session: `claude/booking-engine-safety-features-4vcMQ`. Develop, commit, and push to that branch. Don't push to `main`. (The branch name varies per session — check the system prompt at the top of each new session before committing.)

Note that `Booking engine v4.html` and `v5.html` are each ~2.5 MB; some agent tooling can't transmit them through a single tool call. If a push fails for size/proxy reasons, capture the change as a patch under `patches/` and document it in `patches/README.md` (this is exactly how v4 itself was delivered — see `.branch-notes.md` and `patches/0001-autofill-and-pdf-management.patch`). Prefer `Edit` over `Write` for v5 modifications so only the diff is transmitted.
