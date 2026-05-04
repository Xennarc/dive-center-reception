# Pending patches for Booking engine v4 (offline)

The Claude Code session that authored these changes could not push the
2.5 MB rewrite of the booking engine HTML directly: the agent's git
proxy returned `503 Could not verify PR ownership` for every push, and
the full file is too large to transmit through `mcp__github__push_files`
as a single tool argument.

The behaviour change is captured here as `0001-autofill-and-pdf-management.patch`.
Run `make-v4.sh` from the repo root to produce `Booking engine v4.html`:

```sh
sh make-v4.sh
git add "Booking engine v4.html"
git commit -m "Add Booking engine v4 (offline)"
git push origin main
```

What the patch does in v4 vs. v3.3:

- **Autofill (Add booking modal)** — typing a room number autofills
  guest name and departure date from any matching inhouse PDF record,
  falling back to the most recent record when the booking date is
  outside the guest's stay; fields stay blank when no record matches.
- **Inhouse Reports modal** — a "Loaded PDFs" header with file count
  and Clear-all action, a labelled Remove button per row, and an
  always-visible Add PDFs button alongside the drop zone.
