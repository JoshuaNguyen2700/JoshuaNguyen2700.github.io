"""Checks the site's wiring before you publish, so a new tool can't quietly break another.

Run:  python scripts/check.py      (exits with code 1 if anything is broken)

Checks:
  - assets/tools.json: valid ids, no duplicates, names, statuses, icons, and a page for every tool
  - every tools/<id>/ folder is registered (warning only)
  - no tool reaches into another tool's folder; shared code belongs on the bus (/assets/)
  - every local link, script, stylesheet and redirect in the site's pages points at a real file
  - tools that use the bus link /assets/bus.css and load /assets/core/theme.js
"""
import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
TOOLS = ROOT / 'tools'
TEMPLATE = TOOLS / '_template'
REGISTRY = ROOT / 'assets' / 'tools.json'
ICONS = ROOT / 'assets' / 'core' / 'icons.js'
BUS = '/assets/bus.css'
THEME = '/assets/core/theme.js'
ID = re.compile(r'^[a-z0-9][a-z0-9-]*$')
STATUSES = {'live', 'soon'}
LINKS = [re.compile(r'''\b(?:href|src)\s*=\s*["']([^"'#?]*)''', re.I),
         re.compile(r'''content\s*=\s*["']\s*\d+\s*;\s*url=([^"'#?]+)''', re.I)]
EXTERNAL = re.compile(r'^(?:[a-z][a-z0-9+.-]*:|//)', re.I)
TEXT_FILES = {'.html', '.css', '.js', '.json'}

errors, warnings = [], []


def rel(p):
    return p.relative_to(ROOT).as_posix()


def load_registry():
    try:
        data = json.loads(REGISTRY.read_text(encoding='utf-8'))
    except (OSError, ValueError) as e:
        errors.append(f'{rel(REGISTRY)}: cannot be read ({e})')
        return []
    tools = data.get('tools') if isinstance(data, dict) else None
    if not isinstance(tools, list):
        errors.append(f'{rel(REGISTRY)}: needs a "tools" list')
        return []
    return tools


def check_registry(entries):
    icons = set(re.findall(r'^\s*(\w+)\s*:', ICONS.read_text(encoding='utf-8'), re.M))
    seen, tools = set(), []
    for n, t in enumerate(entries, 1):
        tid = t.get('id') if isinstance(t, dict) else None
        where = f'tools.json entry {n} ({tid or "no id"})'
        if not isinstance(tid, str) or not ID.match(tid):
            errors.append(f'{where}: "id" must use only lowercase letters, numbers and dashes')
            continue
        if tid in seen:
            errors.append(f'{where}: duplicate id')
            continue
        seen.add(tid)
        if not isinstance(t.get('name'), str) or not t['name'].strip():
            errors.append(f'{where}: needs a "name"')
        if t.get('status') not in STATUSES:
            errors.append(f'{where}: "status" must be one of: {", ".join(sorted(STATUSES))}')
        icon = t.get('icon', 'plus')
        if not isinstance(icon, str):
            errors.append(f'{where}: "icon" must be text')
        elif icon.lower().endswith('.svg'):
            if not (TOOLS / tid / icon).is_file():
                errors.append(f'{where}: icon file tools/{tid}/{icon} not found')
        elif icon not in icons:
            errors.append(f'{where}: unknown icon "{icon}" (built-in: {", ".join(sorted(icons))})')
        if not (TOOLS / tid / 'index.html').is_file():
            errors.append(f'{where}: tools/{tid}/index.html is missing')
        tools.append(t)
    return tools


def check_unregistered(ids):
    for d in sorted(TOOLS.iterdir()):
        if d.is_dir() and not d.name.startswith('_') and d.name not in ids:
            warnings.append(f'tools/{d.name}/ is not in tools.json, so it has no tile on the home page')


def check_isolation(ids):
    for tid in ids:
        others = [o for o in ids if o != tid]
        if not others:
            continue
        crossing = re.compile(r'(?:/tools/|\.\./)(' + '|'.join(map(re.escape, others)) + r')/')
        for f in (TOOLS / tid).rglob('*'):
            if f.is_file() and f.suffix in TEXT_FILES:
                m = crossing.search(f.read_text(encoding='utf-8', errors='ignore'))
                if m:
                    errors.append(f'{rel(f)}: reaches into another tool (tools/{m.group(1)}/). Share code through /assets/ instead.')


def pages(tools):
    # Standalone tools are sealed single-file apps: only their entry page is checked.
    sealed = {t['id'] for t in tools if t.get('standalone')}
    for f in sorted(ROOT.rglob('*.html')):
        parts = f.relative_to(ROOT).parts
        if parts[0] == '.git':
            continue
        if len(parts) > 2 and parts[0] == 'tools' and parts[1] in sealed and f.name != 'index.html':
            continue
        yield f


def check_links(tools):
    for f in pages(tools):
        text = f.read_text(encoding='utf-8', errors='ignore')
        for pattern in LINKS:
            for link in pattern.findall(text):
                link = link.strip()
                if not link or EXTERNAL.match(link):
                    continue
                target = ROOT / link.lstrip('/') if link.startswith('/') else f.parent / link
                if link.endswith('/') or target.is_dir():
                    target = target / 'index.html'
                if not target.is_file():
                    errors.append(f'{rel(f)}: broken link "{link}"')


def check_bus(tools):
    for t in tools:
        page = TOOLS / t['id'] / 'index.html'
        if t.get('standalone') or not page.is_file():
            continue
        text = page.read_text(encoding='utf-8')
        if BUS not in text:
            warnings.append(f'{rel(page)}: does not link {BUS}, so it will not get the shared look')
        if THEME not in text:
            warnings.append(f'{rel(page)}: does not load {THEME}, so a saved light/dark choice will not apply')


def main():
    tools = check_registry(load_registry())
    ids = [t['id'] for t in tools]
    check_unregistered(set(ids))
    check_isolation(ids)
    check_links(tools)
    check_bus(tools)
    if not (TEMPLATE / 'index.html').is_file():
        errors.append('tools/_template/index.html is missing (new_tool.py copies it)')
    for w in warnings:
        print('WARNING ', w)
    for e in errors:
        print('ERROR   ', e)
    if errors:
        print(f'\n{len(errors)} problem(s) found. Fix them before publishing.')
        sys.exit(1)
    print(f'OK: {len(tools)} tool(s) registered; links and wiring look good.')


if __name__ == '__main__':
    main()
