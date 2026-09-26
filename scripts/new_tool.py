"""Adds a new tool: copies the blueprint in tools/_template and gives it a tile on the home page.

Usage:    python scripts/new_tool.py <id> "<Name>" ["<one-line description>"]
Example:  python scripts/new_tool.py csv-cleaner "CSV Cleaner" "Tidy up messy CSV files."

The tool starts as "Coming soon". When it works, set its "status" to "live" in assets/tools.json.
"""
import html
import json
import pathlib
import re
import shutil
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
TOOLS = ROOT / 'tools'
TEMPLATE = TOOLS / '_template'
REGISTRY = ROOT / 'assets' / 'tools.json'
ID = re.compile(r'^[a-z0-9][a-z0-9-]*$')


def fail(msg):
    print('ERROR', msg)
    sys.exit(1)


def main():
    if len(sys.argv) not in (3, 4):
        print(__doc__.strip())
        sys.exit(1)
    tid, name = sys.argv[1], sys.argv[2].strip()
    blurb = sys.argv[3].strip() if len(sys.argv) == 4 else 'A new tool is on the way.'
    if not ID.match(tid):
        fail(f'"{tid}" is not a valid id. Use lowercase letters, numbers and dashes, e.g. csv-cleaner')
    if not name:
        fail('the tool needs a name')
    dest = TOOLS / tid
    if dest.exists():
        fail(f'tools/{tid}/ already exists')
    data = json.loads(REGISTRY.read_text(encoding='utf-8'))
    if any(t.get('id') == tid for t in data['tools']):
        fail(f'"{tid}" is already in assets/tools.json')

    shutil.copytree(TEMPLATE, dest)
    for f in dest.rglob('*'):
        if f.is_file() and f.suffix in ('.html', '.css', '.js'):
            quote = html.escape if f.suffix == '.html' else str
            text = f.read_text(encoding='utf-8')
            text = text.replace('{{ID}}', tid).replace('{{NAME}}', quote(name)).replace('{{BLURB}}', quote(blurb))
            f.write_text(text, encoding='utf-8', newline='\n')

    data['tools'].append({'id': tid, 'name': name, 'blurb': blurb, 'status': 'soon', 'icon': 'plus'})
    REGISTRY.write_text(json.dumps(data, indent=2, ensure_ascii=False) + '\n', encoding='utf-8', newline='\n')
    print(f'Created tools/{tid}/ and added "{name}" to the home page.')
    print(f'Preview: http://localhost:8000/tools/{tid}/')
    sys.exit(subprocess.call([sys.executable, str(ROOT / 'scripts' / 'check.py')]))


if __name__ == '__main__':
    main()
