# Patches for the booking engine

Each `Booking engine v*.html` is a 2.5 MB+ self-contained file. Some
agent tooling can't transmit it through a single tool call, and earlier
sessions hit `503 Could not verify PR ownership` on direct pushes of
the full file. As a fallback, behaviour changes are captured as
patches against the previous version, plus a `make-v<n>.sh` script
that rebuilds the new version from a clean copy of the previous one.

## `0001-autofill-and-pdf-management.patch` — v3.3 → v4

Run `sh make-v4.sh` from the repo root to produce `Booking engine v4.html`.

What the patch does in v4 vs. v3.3:

- **Autofill (Add booking modal)** — typing a room number autofills
  guest name and departure date from any matching inhouse PDF record,
  falling back to the most recent record when the booking date is
  outside the guest's stay; fields stay blank when no record matches.
- **Inhouse Reports modal** — a "Loaded PDFs" header with file count
  and Clear-all action, a labelled Remove button per row, and an
  always-visible Add PDFs button alongside the drop zone.

## `0002-multi-user-sync.patch` — v4 → v5

Run `sh make-v5.sh` from the repo root to produce `Booking engine v5.html`.

What the patch does in v5 vs. v4:

- **Op-log replay on save.** `writeBooking`, `setBookingMeta`, and
  `setSheetMeta` each push a structured op into `f.pendingOps`. On
  save, if the file's `lastModified` advanced since load, `reloadFile`
  pulls the disk version and the queued ops replay onto it before the
  workbook is written back — so concurrent saves from 2–3 users on the
  shared SMB folder merge instead of clobbering each other.
- **Cooperative lock files.** `~$<filename>.savelock.<sessionId>` in
  the dirHandle, two-phase acquire (write candidate → list → smallest
  non-stale `sessionId` wins), 30 s stale TTL judged by the lock
  file's own mtime, 10 s heartbeat while held.
- **Background poll.** 8 s `setInterval` while the workspace is
  visible; silent `reloadFile` + toast when a peer's save lands on a
  non-dirty file, `f.remoteChanged = true` flag when the user has
  unsaved edits so the next save forces the conflict path.
- **In-page user-id prompt.** Persisted at
  `localStorage['reefdesk:userId']`, embedded in lock contents and
  "Updated by …" toasts.
