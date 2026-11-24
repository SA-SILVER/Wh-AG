"""Sets the app version in every place that needs it, from one command.
Usage: python3 tools/release.py 1.16 "first note" "second note" ...
Writes: app.js (APP_VERSION), sw.js (cache name), version.json (version, date, notes).
"""
import sys, re, json, datetime, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
v, notes = sys.argv[1], sys.argv[2:]
assert re.fullmatch(r'\d+\.\d+', v), 'version must look like 1.16'
app = root / 'app.js'; s = app.read_text(encoding='utf-8')
s, n = re.subn(r"const APP_VERSION = '[^']*';", f"const APP_VERSION = '{v}';", s); assert n == 1
app.write_text(s, encoding='utf-8')
sw = root / 'sw.js'; s = sw.read_text(encoding='utf-8')
s, n = re.subn(r"const VERSION = 'kesef-[^']*';", f"const VERSION = 'kesef-{v}';", s); assert n == 1
sw.write_text(s, encoding='utf-8')
(root / 'version.json').write_text(json.dumps({'version': v, 'date': datetime.date.today().isoformat(), 'notes': notes}, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
print('version', v)
