#!/usr/bin/env sh
# Build "Booking engine v4.html" by applying the autofill + PDF management
# patch on top of "Booking engine v3.3 (offline).html".
#
# Run from the repo root:
#   sh make-v4.sh
set -eu

SRC="Booking engine v3.3 (offline).html"
DST="Booking engine v4.html"
PATCH="patches/0001-autofill-and-pdf-management.patch"

if [ ! -f "$SRC" ]; then
  echo "missing: $SRC" >&2
  exit 1
fi
if [ ! -f "$PATCH" ]; then
  echo "missing: $PATCH" >&2
  exit 1
fi

cp "$SRC" "$DST"
# Reuse the v3.3 patch hunks against the new file by rewriting the paths.
TMP=$(mktemp)
sed "s|Booking engine v3.3 (offline).html|$DST|g" "$PATCH" > "$TMP"
git apply --whitespace=nowarn "$TMP" || patch -p1 < "$TMP"
rm -f "$TMP"

echo "Wrote $DST"
