#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
INPUT="$ROOT/input/board-memo-input.docx"
OUTPUT="$ROOT/candidate/board-memo-output.docx"
MAP="$ROOT/candidate/replacements.json"
TMP="$ROOT/candidate/tmp"

cp "$INPUT" "$OUTPUT"

deword replace "$OUTPUT" --map "$MAP"

mkdir -p "$TMP"

deword xml "$OUTPUT" -p word/document.xml > "$TMP/document.xml"
deword xml "$OUTPUT" -p word/footnotes.xml > "$TMP/footnotes.xml"

python3 - <<'PY'
import re
from pathlib import Path

base = Path("benchmarks/board-memo-refresh/candidate/tmp")
doc_path = base / "document.xml"
foot_path = base / "footnotes.xml"
doc = doc_path.read_text()
foot = foot_path.read_text()

sentence = "The vendor readiness assessment informed the staffing plan for the first wave of distributor onboarding."


def last_run_start_before(text: str, pos: int) -> int:
    starts = [m.start() for m in re.finditer(r'<w:r(?:\s|>)', text)]
    starts = [s for s in starts if s < pos]
    if not starts:
        raise ValueError("no run start found")
    return starts[-1]

marker = '<w:footnoteReference w:id="1"></w:footnoteReference>'
mi = doc.find(marker)
if mi == -1:
    raise SystemExit("No existing footnoteReference id=1 found")
rs = last_run_start_before(doc, mi)
re_end = doc.find('</w:r>', mi)
if re_end == -1:
    raise SystemExit("Could not isolate reference run end")
ref_run = doc[rs:re_end + len('</w:r>')].replace('w:id="1"', 'w:id="2"', 1)

text_marker = f'<w:t>{sentence}</w:t>'
ti = doc.find(text_marker)
if ti == -1:
    raise SystemExit("Assessment sentence text node not found")
rs2 = last_run_start_before(doc, ti)
re2 = doc.find('</w:r>', ti)
if re2 == -1:
    raise SystemExit("Could not isolate assessment run end")
run_block = doc[rs2:re2 + len('</w:r>')]

m = re.search(r'^<w:r(?P<attrs>[^>]*)>\s*<w:rPr>(?P<rpr>[\s\S]*?)</w:rPr>\s*<w:t>' + re.escape(sentence) + r'</w:t>\s*</w:r>$', run_block, re.S)
if not m:
    raise SystemExit("Assessment run did not match expected shape")
attrs = m.group('attrs')
rpr = m.group('rpr')
new_run_block = (
    f'<w:r{attrs}><w:rPr>{rpr}</w:rPr><w:t>The vendor readiness assessment</w:t></w:r>'
    f'{ref_run}'
    f'<w:r{attrs}><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve"> informed the staffing plan for the first wave of distributor onboarding.</w:t></w:r>'
)
doc = doc[:rs2] + new_run_block + doc[re2 + len('</w:r>'):]

if '<w:footnote w:id="2">' not in foot:
    fm = re.search(r'<w:footnote w:id="1">[\s\S]*?</w:footnote>', foot, re.S)
    if not fm:
        raise SystemExit("Could not find footnote 1 block")
    foot2 = fm.group(0)
    foot2 = foot2.replace('w:id="1"', 'w:id="2"', 1)
    foot2 = re.sub(r'w14:paraId="[0-9A-F]{8}"', 'w14:paraId="A1B2C3D4"', foot2, count=1)
    foot2 = re.sub(r'w14:textId="[0-9A-F]{8}"', 'w14:textId="1A2B3C4D"', foot2, count=1)
    foot2 = foot2.replace(
        'Forecast model updated March 30, 2026 using regional sell-through data.',
        'Assessment drafted April 2, 2026 from confirmed distributor staffing assumptions.'
    )
    foot = foot.replace('</w:footnotes>', foot2 + '\n</w:footnotes>')

doc_path.write_text(doc)
foot_path.write_text(foot)
PY

deword xml "$OUTPUT" --set word/document.xml -i "$TMP/document.xml"
deword xml "$OUTPUT" --set word/footnotes.xml -i "$TMP/footnotes.xml"

echo "Built $OUTPUT"
