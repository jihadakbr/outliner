# Outliner

A lightweight, browser-based tool for building hierarchical outlines and table of contents. Built for planning presentations, structuring documents, and organizing ideas visually.

**Live demo:** _add your GitHub Pages URL here after deployment_

## Features

- **Hierarchical structure** up to 10 levels deep (Chapter → Section → Subsection → Topic → ...) or you can adjust the maximum depth yourself.
- **Outline numbering** auto-computed (1, 1.1, 1.2.3, ...)
- **Drag and drop** to reorder or re-nest items
- **Rich text** with bold and italic formatting
- **Multi-line text** per item (Enter inserts a new line)
- **Expand and collapse** branches for focus
- **Auto-save** to browser localStorage, persists across refreshes
- **Import and export** as JSON
- **Export as PNG** for sharing or embedding in slides
- **Keyboard shortcuts** for fast outlining
- **Zero backend**, pure HTML/CSS/JS, works offline once loaded

## Quick start

1. Clone or download this repo.
2. Open `index.html` in any modern browser.
3. Start typing. Everything auto-saves locally.

No build step, no dependencies to install. The only external resource is `html2canvas` loaded from a CDN for PNG export.

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Enter` | New line inside the current item |
| `Ctrl` + `Enter` | Add a new item at the same level |
| `Tab` | Add a sub-item (child) |
| `Shift` + `Tab` | Collapse or expand the current branch |
| `Ctrl` + `B` | Bold selected text |
| `Ctrl` + `I` | Italic selected text |
| `Ctrl` + `Del` | Delete the current item |

## Drag and drop

Hover any row to reveal the drag handle on the left. Drag an item onto another row to move it:

- **Top area** of target: drop as a sibling above
- **Middle area** of target: drop as a child (nest inside)
- **Bottom area** of target: drop as a sibling below

Drops that would exceed the max-depth setting are refused automatically.

## Project structure

```
Outliner/
├── index.html        Markup and entry point
├── css/
│   └── styles.css    All styling
├── js/
│   └── app.js        All application logic
├── README.md
└── .gitignore
```

## Data storage

All work is saved to `localStorage` under the key `toc-builder.v2`. Nothing leaves your browser. To back up or share an outline, use **Save** to download a JSON file, or **PNG** to export an image.

## Browser support

Tested on recent versions of Chrome, Edge, and Firefox. Requires a browser with `contenteditable`, `localStorage`, and HTML5 drag-and-drop support.

## License

MIT
