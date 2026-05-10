#!/usr/bin/env bash
# Usage: ./fetch_transcript.sh <youtube_url_or_id>
# Outputs: video.en.srt and video.txt (plain text, no timestamps) in the cwd.
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "Usage: $0 <youtube_url_or_id>" >&2
  exit 1
fi

URL="$1"

command -v yt-dlp >/dev/null || { echo "Installing yt-dlp..."; pip install --user yt-dlp; }

yt-dlp \
  --write-auto-sub --write-sub \
  --sub-lang en --sub-format "vtt/srv1/best" \
  --skip-download \
  --convert-subs srt \
  -o "video.%(ext)s" \
  "$URL"

SRT=$(ls video*.srt | head -1)
echo "Got: $SRT"

python3 - "$SRT" <<'PY'
import re, sys, pathlib
src = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
# Drop index lines, timestamp lines, and blank lines.
lines = []
for line in src.splitlines():
    if not line.strip():
        continue
    if re.fullmatch(r"\d+", line.strip()):
        continue
    if "-->" in line:
        continue
    # Strip inline tags like <c> </c> and {\an8} positioning.
    line = re.sub(r"<[^>]+>", "", line)
    line = re.sub(r"\{[^}]+\}", "", line)
    lines.append(line.strip())
# Deduplicate consecutive identical lines (common in auto-captions).
clean = []
for l in lines:
    if not clean or clean[-1] != l:
        clean.append(l)
out = " ".join(clean)
pathlib.Path("video.txt").write_text(out + "\n", encoding="utf-8")
print(f"Wrote video.txt ({len(out)} chars, {len(clean)} caption lines)")
PY
