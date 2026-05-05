#!/usr/bin/env sh
# Build "Booking engine v5.html" by applying the multi-user sync patch
# on top of "Booking engine v4.html".
#
# Run from the repo root:
#   sh make-v5.sh
set -eu

SRC="Booking engine v4.html"
DST="Booking engine v5.html"
PATCH="patches/0002-multi-user-sync.patch"

if [ ! -f "$SRC" ]; then
  echo "missing: $SRC" >&2
  exit 1
fi
if [ ! -f "$PATCH" ]; then
  echo "missing: $PATCH" >&2
  exit 1
fi

cp "$SRC" "$DST"
# Reuse the v4 patch hunks against the new file by rewriting the paths.
TMP=$(mktemp)
sed "s|Booking engine v4.html|$DST|g" "$PATCH" > "$TMP"
git apply --whitespace=nowarn "$TMP" || patch -p1 < "$TMP"
rm -f "$TMP"

echo "Wrote $DST"
