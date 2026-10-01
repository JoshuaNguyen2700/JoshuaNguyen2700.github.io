# Joshua's GitHub website

A home page of tools. Each tool is a self-contained module that plugs into a shared "main bus",
like the resource highway in Factorio: shared things flow down the bus, and each tool taps off
what it needs without touching the others.

```
                ┌──────────────────── assets/  (THE MAIN BUS) ────────────────────┐
                │ bus.css        shared look: tokens → base → components          │
                │ core/shell.js  shared nav bar + footer                           │
                │ core/*.js      registry loader, tiles, icons, helpers            │
                │ tools.json     the list of tools (one entry per tool)            │
                └─────┬──────────────┬──────────────┬──────────────┬──────────────┘
                      │              │              │              │
                 index.html     tools/sla/     tools/otp/      tools/<new>/
                 (home page)    (sealed app)   (bus tool)      (next tool)
```

## Add a tool

```powershell
python scripts/new_tool.py csv-cleaner "CSV Cleaner" "Tidy up messy CSV files."
```

This copies the blueprint in `tools/_template/` to `tools/csv-cleaner/`, adds a tile to the home
page, and runs the checker. Build the tool in its folder: `index.html`, `tool.css`, `tool.js`.
When it works, change its `"status"` from `"soon"` to `"live"` in `assets/tools.json`.

## Rules of the bus

1. **A tool owns only its own folder.** Everything for `tools/<id>/` lives inside it.
2. **Tools never reach into each other.** Anything two tools need goes on the bus (`assets/`).
3. **Bus paths are absolute, tool paths are relative.** `/assets/bus.css`, but `tool.css`.
4. **The bus only grows.** Add new tokens, classes, icons and helpers freely; don't rename or
   remove existing ones, because other tools may use them.
5. **Tool styles always win.** Bus styles sit in CSS layers, so a tool's own CSS overrides them
   without affecting any other page.
6. **Sealed tools are allowed.** A tool marked `"standalone": true` (like the SLA Dashboard, which
   also works offline as a single file) doesn't use the bus at all.

## Scripts

| Command | What it does |
|---|---|
| `python scripts/dev.py` | Live preview at http://localhost:8000; pages reload when files change |
| `python scripts/check.py` | Checks the registry, links and tool isolation. Run before publishing |
| `python scripts/new_tool.py ...` | Adds a new tool from the blueprint |
| `tools\sla\build.ps1` | Rebuilds the SLA Dashboard's single shareable file |

Put a sample Salesforce export in `tools/sla/samples/` and the preview loads it into the dashboard
after every reload. `samples/` folders are never committed.

## Publish

```powershell
python scripts/check.py
git add -A
git commit -m "Describe the change"
git push
```

GitHub Pages updates the live site about a minute after the push.
