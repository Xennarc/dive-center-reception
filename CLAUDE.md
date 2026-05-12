# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository shape

No package manager, no build system, no test runner for the apps themselves — each `Booking engine v*.html` is a self-contained single-page app (HTML + inlined CSS + inlined JS + inlined `ExcelJS`, `JSZip`, and `pdfjs-dist`). Dev loop = open the file in a Chromium browser.

Versions live side-by-side. **`Booking engine v5.html`** (~5295 lines, ~2.7 MB) is canonical. `Booking engine v4.html` and `Booking engine v3.3 (offline).html` stay at the repo root because `make-v4.sh` / `make-v5.sh` reference them. All pure-history files (`v3.2`, `V2`, `v3-bookings`, `Booking sheets unified V1`, scratch v4 copies) live in `archive/`; don't edit them. Sample PDFs in `samples/`.

Sibling tool **`sheet-shaper.html`** (~1419 lines) — see its own section below.

Node-side: `package.json` declares `npm test` → `node test.js`, devDeps `@xmldom/xmldom`, `exceljs`, `jszip`. `benchmark.js` is a one-off perf microbench (safe to ignore). `.gitignore` covers `node_modules`, `package-lock.json`, `.claude/worktrees/`.

The `patches/` directory + `make-v*.sh` scripts exist only as a delivery fallback for when agent tooling can't transmit the 2.5 MB+ HTML in one tool call — see `patches/README.md`. v5 is edited directly; don't try to rebuild it from a patch unless you're starting from a clean v4.

## Reading the HTML files

`Booking engine v5.html` is too large to `Read` whole-file. Use line offsets:

- **1–11** — `<head>`, fonts, meta.
- **12–59** — inlined ExcelJS@4.4.0 — **skip**.
- **60–75** — inlined JSZip@3.10.1 — **skip**. (Used by Save preservation.)
- **76–100** — inlined pdfjs-dist@3.11.174 — **skip**.
- **101–1136** — `<style>` block (one in-style banner at **965**: "Operational tracking — status pills, banners, capacity override, saving indicator").
- **1139–1250** — `<body>` markup (`#status-pill`, `#saving-indicator`, `#btn-save`, `#workspace`, `#modal-mount`, `#toast-mount`, hidden `#file-input` / `#pdf-input`).
- **1251–5292** — application `<script>`. The pdf.js **worker** is inlined as a string literal in the first ~50 lines of this script; ignore that block when grepping. `const STATE = { … }` at **1273**.

`grep -n '/\* ====' "Booking engine v5.html"` jumps between sections. Current banners:

| Line | Section |
|------|---------|
| 1252 | Reef Desk — unified booking interface |
| 1317 | Date helpers |
| 1370 | Cell helpers |
| 1404 | Layout detection |
| 1551 | Booking I/O |
| 1706 | Meta sheets (hidden) |
| 1925 | File load |
| 2027 | Save / reload |
| **2217** | **Save preservation** |
| 2585 | Sync (multi-user) |
| 2940 | File source |
| 3065 | View shell |
| 3122 | Date strip |
| 3171 | Main view dispatch |
| 3184 | Agenda view (selected date) |
| 3521 | People (guests) view |
| 3614 | Search results |
| 3691 | Modals |
| 4266 | Inhouse PDF parsing |
| 4780 | Inhouse reports modal |
| 4896 | Stats dashboard |
| 5184 | Wire-up |

## Architecture

**Single global `STATE`** (line 1273) holds everything: `mode` (`'fs'` File System Access API vs `'fp'` manual file picker), `view` (`'agenda' | 'people' | 'stats'`), `dirHandle`, `files[]`, `bookings[]`, `excursions[]`, `excFilter`, `search`, `selectedDate` (ISO), `showCancelled`, `inhouseGuests[]`, `inhouseReports[]`, `userId` (operator name, persisted at `localStorage['reefdesk:userId']`), `saving` (Set of filenames currently mid-save). Renders are pull-based: mutate `STATE`, then call `renderView()` / `renderDateStrip()` / `updateStatus()` / `updateSavingIndicator()`.

**One Excel file = one excursion; one sheet = one day.** `detectLayout(ws)` scans the first ~20 rows for a row with both "room" and "guest" headers (`HEADER_KEYS`), then walks back up to find the excursion title and capacity. Unrecognised files, `~$` lock files, and the two hidden meta sheets are filtered at load time.

**Hidden meta sheets** (`_BookingMeta`, `_SheetMeta`) carry data the spreadsheet format lacks columns for: per-row `status` / `cancelReason` / `attended` / `paid` / postpone log; per-sheet `activityCancelled` / `customCapacity`. Schema constants: `META_BOOKING_SHEET`, `META_SHEET_SHEET`, `BOOKING_META_COLS`, `SHEET_META_COLS`, `BOOKING_META_DEFAULTS`, `SHEET_META_DEFAULTS`. `readMetaSheets(wb)` builds the in-memory maps on load; `flushMetaSheets(fileEntry)` writes them back before `wb.xlsx.writeBuffer()` in `saveAll()`. Always go through `getBookingMeta` / `setBookingMeta` / `getSheetMeta` / `setSheetMeta` so dirty-tracking stays correct.

**Two file-source modes** share the same `loadFiles(specs)` shape:

- `'fs'` — `showDirectoryPicker({ mode: 'readwrite' })` → `openDirectory(handle)`. `saveAll()` writes via `fileHandle.createWritable()`. `openDirectory` sweeps stale `~$*.savelock.*` files and calls `offerCrashRecovery()` after `loadFiles`.
- `'fp'` — `<input type="file" multiple accept=".xlsx">`. `saveAll()` falls back to a download blob. `reload`, sync, and crash recovery are no-ops in this mode.

### Save preservation (line 2217)

ExcelJS's `writeBuffer()` round-trips the workbook through its own model and drops parts it doesn't understand (drawings, external links, threaded comments, conditional formatting, printer settings, charts, tables, …). Excel then refuses the resulting `.xlsx` with the "We found a problem with some content" recovery dialog. `applyPreservation()` (constants `PRESERVED_PART_PREFIXES`, `PRESERVED_COMMENTS_RE`; helper `isPreservedPath`) unzips both the original `.xlsx` and ExcelJS's output, splices preserved parts back verbatim, merges `[Content_Types].xml` and `*.rels`, and runs the merge-orphan fix (strip stray `<c>` inside a `<mergeCell>` that isn't the anchor — Excel rejects those too). This is why JSZip is inlined alongside ExcelJS. `?diag=1` dumps the pre-preservation diff. Save order is fixed: **ExcelJS write → preserve → file write** — don't reorder.

### Multi-user sync (line 2585; `'fs'` mode only)

v4 overwrote each `.xlsx` wholesale, so two operators on an SMB share would clobber each other. v5 fixes this with:

- **Op-log replay.** `writeBooking` / `setBookingMeta` / `setSheetMeta` push structured ops into `f.pendingOps` (alongside `f.dirty = true`). On save, if disk `lastModified` advanced since load, `reloadFile(f)` and `replayOps` replay the ops onto the fresh workbook before write. `bookingTempId` threads create-then-meta flows; `txnId` keeps multi-op flows atomic.
- **Advisory locks.** `~$<filename>.savelock.<sessionId>` in `dirHandle`. Two-phase acquire, 12 s stale TTL judged by lock-file mtime, 10 s heartbeat, drained best-effort by `pagehide` from a module-level `_heldLocks` Set. Lock is advisory — the mtime check + `replayOps` is the real safety net.
- **Background poll.** 8 s `setInterval` while visible (`document.hidden` pauses it; permission downgrade aborts the tick silently). Remote mtime advanced on a non-dirty file → silent `reloadFile` + "Updated by …" toast; on a dirty file → `f.remoteChanged = true` so the next save forces the conflict branch.
- **User identity.** Prompted once via an in-page `el()` modal, persisted at `localStorage['reefdesk:userId']`, embedded in lock contents and toasts.
- **Crash recovery.** Each op also mirrors to `localStorage['reefdesk:pendingOps:<dirHandle.name>:<filename>']`. `openDirectory` calls `offerCrashRecovery()` after load; Restore re-runs `replayOps` against the fresh workbook. `'fp'` mode skips persistence.
- **Save-in-flight UI.** A `#saving-indicator` pill is toggled by `updateSavingIndicator()` on every `STATE.saving.add/delete`. `beforeunload` blocks on `STATE.saving.size > 0` (not just dirty files), because a tab killed mid-write can corrupt the `.xlsx`.

### Inhouse PDF parsing (line 4266)

TRML / TRMD daily meal-plan reports are parsed via pdf.js text items with x/y coordinates rather than naive line extraction (`classifyInhouseHeader` recognises a wide vocab; falls back to `parseFilenameDate` when the header lacks a date). `lookupInhouse(room, dateISO)` deliberately falls back to the **most recent record for the room** when no date is given or the booking date is outside the guest's stay — this is what drives the autofill in the New Booking modal even with stale reports; stale/out-of-range cases surface in the modal status hint via `tryAutofill()`.

## Sheet Shaper (`sheet-shaper.html`)

Standalone Reef Desk companion tool, **not a v6**. Drop a folder of empty excursion masters, pick which weekdays each runs, generate clean dated workbooks for a target month — output feeds back into the booking engine.

- Own `STATE` (~line 633: `mode`, `dirHandle`, `outDirHandle`, `masters`, `targetMonth`, `format`), own `el()` helper, own CSS. localStorage key `sheetshaper.daypatterns.v1` (no `reefdesk:` prefix). `Booking engine v5.html` has no reference to it; treat it as a sibling, not a sub-page.
- Generator entry point: `generateAll()` (~line 1305) — a **surgical XML-level workbook clone** that bypasses ExcelJS's write path entirely. This avoids the same recovery-dialog corruption that drove the Save-preservation pipeline in v5.
- **Load-bearing markers:** `/* === DATE HELPERS BEGIN === */ … END === */` (~672–723) and `/* === SURGICAL CLONE BEGIN === */ … END === */` (~962–1286). `test.js` extracts these blocks by regex and runs them in a Node `vm` sandbox — **do not remove or rename them.**
- Only JSZip 3.10.1 is inlined; no ExcelJS, no pdf.js. Update by replacing the whole minified block from the upstream npm tarball.

## Tests

`npm test` (= `node test.js`) is the only automated check in the repo. Requires `npm install` first (`node_modules/` and `package-lock.json` are gitignored). devDeps: `@xmldom/xmldom`, `exceljs`, `jszip`.

The test extracts the two BEGIN/END blocks from `sheet-shaper.html`, runs them under `vm.runInContext` with `@xmldom/xmldom` standing in for the browser DOM, builds a synthetic Reef-Desk-shaped master via ExcelJS, runs the generator, then asserts the output zip is structurally clean: correct sheet count, `[Content_Types].xml` and workbook rels updated, no `xmlns=""` anywhere (the W3C-correct browser serialiser emits this for any element appended without inheriting the parent's default namespace — that's the regression PR #29 fixed; the test guards by both static-scanning the extracted source for `createElement(` and string-checking the output), no `calcChain`, no value-carrying `<c>` inside a non-anchor merge cell, and `{{date}}` substitution applied. Optionally runs `xmllint --noout` on every XML part and a `soffice --convert-to xlsx` round-trip — both skipped silently when not on PATH (don't add hard installs).

`benchmark.js` is unrelated; ignore unless touching the unique-rooms code path.

## Conventions

- No build system, no framework, no package manager for the apps. Double-click the HTML; that's the dev loop.
- Don't extract helpers into separate files. Keep code in the `<script>` block, grouped under the existing `/* === … === */` banners.
- DOM construction uses `el(tag, attrs, ...children)` (line 1181 in v5; supports `class`, `html`, `style`-as-object, `on*` listeners). Use it rather than ad-hoc `document.createElement` chains.
- New modals: `openModal(node)` / `closeModal()`. There is no modal stack — opening a second overwrites the first.
- State that must survive a tab close round-trips through `localStorage` under the `reefdesk:` prefix (or `sheetshaper.` for Sheet Shaper).
- Never edit the inlined ExcelJS / JSZip / pdf.js / pdf.worker blocks by hand. Replace the whole block from upstream if updating.
- Folder access requires a Chromium browser opened directly (not an iframe). The welcome card already says so where it matters.
- Keyboard shortcuts (Wire-up, line 5184): Cmd/Ctrl+S save, Cmd/Ctrl+N new booking, Cmd/Ctrl+K focus search, Esc closes modal or clears search, Left/Right arrows move the selected date. `beforeunload` warns when any file is dirty **or** a save is in flight; `pagehide` drains `_heldLocks` best-effort.
- Prefer `Edit` over `Write` for v5 modifications so only the diff is transmitted (the file is ~2.7 MB).

## Git workflow

Each session is on its own working branch (check the system prompt at the top of the session). Develop, commit, and push to that branch. **Don't push to `main`.**

If a push fails for size/proxy reasons on a v4/v5 HTML edit, capture the change as a patch under `patches/` and document it in `patches/README.md` — that's exactly how v4 itself was delivered (see `patches/0001-autofill-and-pdf-management.patch`). For changes that only touch v5 going forward, edit `Booking engine v5.html` directly and commit it as a binary blob; only re-run `make-v5.sh` when starting from a clean v4.
