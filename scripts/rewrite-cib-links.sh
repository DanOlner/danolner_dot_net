#!/usr/bin/env bash
#
# Rewrite Covered in Bees archive links in the Quarto source files.
#
#   https://coveredinbees.org.archived.website/node/354.html
#     -> https://danolner.github.io/cib_archive/node/354.html
#
# Only touches hand-written source (.qmd / .md). Build output and the old
# site snapshot are skipped -- those regenerate or are frozen history.
#
# Usage:
#   scripts/rewrite-cib-links.sh --dry-run   # show what would change
#   scripts/rewrite-cib-links.sh             # apply
#
# Idempotent: running twice is a no-op. Undo with `git checkout -- _sections`.

set -euo pipefail

OLD_HOST="coveredinbees.org.archived.website"
NEW_ROOT="https://danolner.github.io/cib_archive"

# Regex-safe form of the host (dots escaped), for grep.
OLD_ESC=$(printf '%s' "$OLD_HOST" | sed 's/\./\\./g')

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

DRY_RUN=0
[[ "${1:-}" == "--dry-run" || "${1:-}" == "-n" ]] && DRY_RUN=1

# Directories that hold generated or archived copies, not source.
find_sources() {
  find . \( -path ./docs -o -path ./_site -o -path ./old_danolner_dot_net_site \
             -o -path ./.quarto -o -path ./_freeze -o -path ./.git \
             -o -name '*_files' \) -prune -o \
       -type f \( -name '*.qmd' -o -name '*.md' \) -print \
  | sort
}

total=0
changed_files=0

while IFS= read -r f; do
  [ -n "$f" ] || continue
  n=$( { grep -oE "https?://$OLD_ESC" "$f" 2>/dev/null || true; } | wc -l | tr -d ' ')
  [ "$n" -eq 0 ] && continue

  total=$(( total + n ))
  changed_files=$(( changed_files + 1 ))
  printf '%4s  %s\n' "$n" "${f#./}"

  if [[ "$DRY_RUN" -eq 0 ]]; then
    perl -pi -e "s{https?://\Q${OLD_HOST}\E/?}{${NEW_ROOT}/}g" "$f"
  fi
done <<EOF
$(find_sources)
EOF

echo
if [[ "$changed_files" -eq 0 ]]; then
  echo "No '${OLD_HOST}' links found -- nothing to do."
elif [[ "$DRY_RUN" -eq 1 ]]; then
  echo "DRY RUN: would rewrite ${total} link(s) across ${changed_files} file(s)."
  echo "Re-run without --dry-run to apply."
else
  echo "Rewrote ${total} link(s) across ${changed_files} file(s)."
  echo "Undo with: git checkout -- _sections"
fi

# Flag any stragglers outside the source set, for information only.
if [ "$DRY_RUN" -eq 0 ]; then
strays=$(grep -rl "$OLD_HOST" --include='*.qmd' --include='*.md' --include='*.html' . 2>/dev/null \
         | grep -vE '^\./(old_danolner_dot_net_site|\.git)/' | grep -v '^\./docs/' || true)
if [[ -n "$strays" ]]; then
  echo
  echo "Note -- '${OLD_HOST}' still appears in (generated output, regenerates on render):"
  echo "$strays" | sed 's|^\./|  |'
fi
fi
