# Pending patches

The push of the actual `Booking engine v3.3 (offline).html` change failed inside this Claude Code session — the agent's git proxy returned `503 Could not verify PR ownership` on every retry, and the 2.5 MB rewrite was too large to transmit through `mcp__github__push_files` as a single tool argument.

Apply locally instead:

```sh
git fetch origin claude/update-booking-engine-v3.3-CzPwU
git checkout claude/update-booking-engine-v3.3-CzPwU
git am patches/0001-autofill-and-pdf-management.patch
git push origin claude/update-booking-engine-v3.3-CzPwU
```

The patch updates `Booking engine v3.3 (offline).html` to:

- Autofill guest name + departure date as soon as a room number is entered in the Add booking modal; fall back to the latest record for that room when the booking date is outside the guest's stay; leave fields blank when no PDF record matches.
- Show a Loaded PDFs section in the Reports modal with a file count, Clear-all action, labelled Remove button per row, and an always-visible Add PDFs button.
