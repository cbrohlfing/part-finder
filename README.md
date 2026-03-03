# Part Finder

A fast, lightweight Windows desktop utility for searching engineering part files across multiple network locations.

Built with PowerShell + WinForms for internal engineering workflows.

---

## 🚀 Download Latest Release

Always download the newest production version from:

**GitHub → Releases → Latest**

[![Download Latest Release](https://img.shields.io/badge/Download-Latest%20Release-brightgreen?style=for-the-badge)](https://github.com/cbrohlfing/part-finder/releases/latest)
[![Latest Version](https://img.shields.io/github/v/release/cbrohlfing/part-finder?display_name=tag&style=for-the-badge)](https://github.com/cbrohlfing/part-finder/releases/latest)


Download:

PartFinder_vX.Y.zip

*(Replace X.Y with the most recent version number.)*

---

## 📦 Installation (Coworker Setup)

### 1️⃣ Extract the ZIP

Extract `PartFinder_vX.Y.zip` to:

C:\Scripts\PartFinder

After extraction, your folder structure should look like:

C:\Scripts\PartFinder
├── run-latest.ps1
└── vX.Y
  └── Part-Finder.ps1

---

### 2️⃣ Create Desktop Shortcut

Create a shortcut with this target:

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "C:\Scripts\PartFinder\run-latest.ps1"

This ensures:

- Execution policy does not block the script
- The window launches cleanly
- The newest installed version runs automatically

---

### 3️⃣ Updating to a New Version

When a new version is released:

1. Download the newest ZIP from GitHub Releases.
2. Extract it into:

C:\Scripts\PartFinder

3. A new version folder (e.g. `v1.8`) will be added.
4. No shortcut changes are required.

`run-latest.ps1` automatically launches the highest installed version.

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

Stored locally in:

part_finder_settings.json

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
- **v1.7** — Professional Dev → Release workflow and automated packaging

---

## 🛠 Tech Stack

- PowerShell
- WinForms
- JSON configuration
- Git version control
- GitHub Releases distribution

---

## 🎯 Intended Use

Built to support engineering drawing workflows in shared network environments where:

- Drawings are distributed across multiple project directories
- Fast lookup is critical
- File ownership and metadata matter

---

## 🔮 Planned Improvements

- File preview pane
- Content search inside PDFs / text
- Check All / Uncheck All folder controls
- Optional dark mode

---

## 👨‍💻 Author

Chris Rohlfing
Engineering Automation
