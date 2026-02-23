# part-finder

# Part Finder

A fast, lightweight Windows desktop utility for searching engineering part files across multiple network locations.

Built with PowerShell + WinForms for internal engineering workflows.

---

## 🚀 Overview

Part Finder is designed to quickly locate drawing files (DWG, PDF, etc.) across multiple shared directories.

Instead of manually browsing network drives, Part Finder lets you:

- Search across multiple folders simultaneously
- Toggle folders on/off per search
- View key file metadata instantly
- Persist window layout and preferences
- Maintain versioned releases

---

## 🔍 Features

### Multi-Folder Search
- Add unlimited folders (UNC or mapped drives)
- Enable/disable folders using checkboxes (v1.6+)
- Search runs across all checked folders

### Search Modes
- Filename contains
- (Future expansion-ready for additional modes)

### Search Controls
- Include subfolders toggle
- Max results limiter
- Ignore file extensions (e.g. `.log`, `.bak`)

---

## 📁 Folder Management

- Add folders via Browse or paste UNC path
- Remove selected folders
- Clear all folders
- Toggle folders on/off without removing them
- Folder states persist between sessions

---

## 🧾 File Metadata Panel

Displays:

- Name
- Full path
- Type
- Size
- Date Modified
- Date Created
- Owner (e.g. `VALMONT\tj729593`)

Metadata updates instantly when selecting a result.

---

## 🖥 UI Features

- Resizable window
- Persistent window size and position (v1.5+)
- Persistent splitter position
- Clean layout with fixed metadata panel
- Word-wrapped dynamic folder hints

---

## 💾 Settings Persistence

Stored in: part_finder_settings.json

Persists:
- Folder list + enabled state
- Search preferences
- Window size / position
- Splitter layout

Backward compatible with older settings format.

---

## 🏷 Version History Highlights

- **v1.4** — Layout refinements, metadata improvements
- **v1.5** — Owner field added, persistent window sizing
- **v1.6** — Checkbox-enabled folders

---

## 🛠 Tech Stack

- PowerShell
- WinForms
- JSON configuration
- Git version control

---

## 🎯 Intended Use

This tool was built to support engineering drawing workflows in a shared network environment where:

- Drawings are distributed across multiple project directories
- Fast lookup is critical
- File ownership and metadata matter

---

## 📦 Installation

Run the appropriate version folder: v1.x/Part-Finder.ps1

Or use the provided desktop shortcut.

---

## 🔮 Planned Improvements

- File preview pane
- Content search inside PDFs / text
- Check All / Uncheck All folder controls
- Optional dark mode

---

## 👨‍💻 Author

Chris Rohlfing
Built for internal engineering automation.

---
