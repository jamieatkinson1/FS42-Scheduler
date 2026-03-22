# &#x1F4FA; FS42 Scheduler

FS42 Scheduler is a lightweight, single-user planning utility for building linear channel schedules for FieldStation42 (FS42). It is designed for visual scheduling, operational checks, and export-ready handoff without the complexity of a full traffic, rights, or collaboration system.

## &#x1F464; Who It Is For

- Individual channel schedulers and programmers
- FS42 users who want a local planning board
- Small operations teams that need a simple, single-user workflow

## &#x1F3AF; What It Solves

- Plans channels in a visual timeline instead of a spreadsheet
- Keeps programmes, commercial items, bumpers, promos, and continuity material organised
- Surfaces conflicts and export issues before handoff
- Produces FS42-oriented per-channel JSON exports

## &#x2728; Key Features

- Scheduler-first timeline with drag, drop, snap, and resize
- Channel management with groups, colours, and taglines
- Lane and category planning for clearer operational layout
- Operational flags: watershed restricted, prime time, and must-run
- Validation and export-readiness feedback
- Table view for review and audit work
- Export profiles for `Internal scheduler JSON` and `FS42 station config`
- CSV export for spreadsheet review
- Local `localStorage` persistence
- Duplicate item and keyboard nudge shortcuts

## &#x26A0;&#xFE0F; Current Scope & Limitations

- Single-user only
- No collaboration, permissions, traffic, rights, or approvals workflow
- No backend, server, build step, or framework dependency
- Export output is FS42-oriented, but should still be checked against your ingest expectations

## &#x25B6;&#xFE0F; Run Locally

No install step is required.

1. Open `index.html` in a browser.
2. Or serve the folder locally with:

```powershell
python -m http.server 8000
```

3. Open [http://localhost:8000](http://localhost:8000).

## &#x1F9ED; Basic Workflow

1. Define channels.
2. Add items.
3. Place and edit items in the timeline.
4. Review the plan in table view.
5. Export when the schedule is ready for FS42.

## &#x1F4E4; Export Profiles

- `Internal scheduler JSON` is the more permissive export option for working drafts or local review.
- `FS42 station config` exports one FieldStation42-compatible `station_conf` JSON file per channel and blocks export when the generated config is invalid.
- Commercial items are planned in the workspace but excluded from the per-channel schedule payload.
- If a channel contains commercial material, the exported station config keeps the commercial settings available rather than silently rewriting them.
- `Use multi-logo mode` controls whether the exporter writes a string `multi_logo` profile name or leaves it empty.

## &#x1F9E0; Suggested Operating Pattern

1. Set up channels first.
2. Use consistent lane categories for programmes, commercials, and continuity work.
3. Keep `Export code` values stable for anything that needs to survive ingest or re-import.
4. Use `Block family` for related blocks, strands, or franchise groupings.
5. Review conflicts and readiness warnings before exporting.
6. Export one JSON file per channel for FieldStation42.

## &#x1F4C1; Repository Layout

- `index.html` - app shell and UI
- `styles.css` - presentation and layout
- `app.js` - scheduler logic, validation, and exports
- `examples/` - place sample export files here
- `README.md` - project overview and workflow
- `LICENSE` - project licence

## &#x2139;&#xFE0F; Notes

This project is intended as a practical local planning tool. It helps you build clear channel schedules for FS42, but it is not a replacement for a full multi-user broadcast traffic platform.

If you want to share sample output, place example JSON files in `examples/` so the repo stays organised without adding runtime complexity.
