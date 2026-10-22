#!/bin/sh
# docx-js가 넣는 비표준 <w:highlightCs/> 제거
set -e
IN="$1"; OUT="$2"; T=$(mktemp -d -p "$(dirname "$OUT")")
unzip -q "$IN" -d "$T"
sed -i 's#<w:highlightCs [^>]*/>##g' "$T/word/document.xml"
(cd "$T" && rm -f "$OUT.tmp" && zip -qXr "$OLDPWD/$OUT.tmp" .)
mv "$OUT.tmp" "$OUT"; rm -rf "$T"
