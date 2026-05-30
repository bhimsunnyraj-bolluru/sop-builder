# SOP Builder

**SOP Builder** is a Windows desktop app for capturing SAP GUI screens, building step-by-step Standard Operating Procedures (SOPs), annotating screenshots, and exporting polished Word documents. Built with Electron for the **Business Core Solutions** team.

Think of it as a lightweight Scribe-style tool focused on SAP process documentation.

---

## Features

| Feature | Description |
|---------|-------------|
| **Screenshot capture** | Capture the active SAP GUI window while you work |
| **SAP context detection** | Auto-fills step titles from window name, transaction code, and status bar (when SAP Scripting is enabled) |
| **Global hotkey** | Configurable shortcut (default `Alt+C`) to capture without clicking the app |
| **Compact mode** | Shrinks to a floating toolbar so SAP stays in focus |
| **Annotation tools** | Box, arrow, highlight, numbered callouts, and text labels |
| **Step management** | Drag to reorder, edit descriptions, delete steps |
| **Save / Open** | Projects saved as JSON (`.json`) under `data/sops/` |
| **DOCX export** | Exports a Word document with metadata and annotated screenshots |
| **SOP metadata** | Author, version, review date (defaults to today), and title |

---

## Two ways to run SOP Builder

| Who | How |
|-----|-----|
| **Most team members** | Install the `.exe` (no Node.js needed) — see [Install from EXE](#install-from-exe-recommended) |
| **Developers** | Clone the repo and run `npm start` — see [Developer setup](#developer-setup) |

---

## Install from EXE (recommended)

You do **not** need Node.js or `npm start` for daily use. Someone on the team builds the installer once; everyone else double-clicks the app.

### For end users

1. Get the installer from your team lead or from the [GitHub Releases](https://github.com/bhimsunnyraj-bolluru/sop-builder/releases) page:
   - **`SOP Builder Setup 1.0.0.exe`** — installs like any Windows app (Start menu + desktop shortcut)
   - **`SOP Builder 1.0.0.exe`** — portable version; no install, runs from any folder
2. Run the installer or portable `.exe`.
3. Launch **SOP Builder** from the Start menu, desktop, or the portable file.

Your SOP projects, screenshots, and exports are stored **next to the installed app** (or next to the portable `.exe`) in `data/`, `screenshots/`, and `exports/`.

### For whoever builds the EXE (once per release)

On a Windows machine with Node.js installed, from the project folder:

```powershell
npm install
npm run dist
```

Output appears in the `dist/` folder:

| File | Purpose |
|------|---------|
| `dist/SOP Builder Setup 1.0.0.exe` | NSIS installer for teammates |
| `dist/SOP Builder 1.0.0.exe` | Portable single-file exe |
| `dist/win-unpacked/` | Unpacked app folder (for testing) |

Share the **Setup** or **portable** exe with the team (Teams, SharePoint, GitHub Releases, etc.).

Portable build only:

```powershell
npm run dist:portable
```

---

## Developer setup

Use this if you are changing the code or running from source.

### Requirements

- **Windows 10 or 11**
- **Node.js 18 LTS or newer** — [Download Node.js](https://nodejs.org/)
- **SAP GUI for Windows** (for capturing SAP screens)
- **Optional:** SAP GUI Scripting enabled for richer step titles

### 1. Install Node.js

1. Go to [https://nodejs.org/](https://nodejs.org/)
2. Download and install the **LTS** version
3. Open **PowerShell** or **Command Prompt** and verify:

```powershell
node --version
npm --version
```

Both commands should print version numbers (e.g. `v20.x.x` and `10.x.x`).

### 2. Clone the repository

```powershell
git clone https://github.com/bhimsunnyraj-bolluru/sop-builder.git
cd sop-builder
```

If your organization uses a different GitHub org or repo URL, use that URL instead.

### 3. Install dependencies

From the project folder:

```powershell
npm install
```

This downloads Electron and other packages into `node_modules/`. It may take a few minutes the first time.

### 4. First-run setup (optional)

Copy the example settings file if you want a local config (the app creates defaults automatically on first run):

```powershell
copy data\settings.example.json data\settings.json
```

Open **Settings** (⚙️) in the app to change the capture hotkey if `Alt+C` conflicts with another tool.

### 5. Run from source

```powershell
npm start
```

The SOP Builder window opens. Keep it running while you document SAP transactions.

To build a distributable `.exe` for teammates, see [Install from EXE](#install-from-exe-recommended).

---

## Quick Start Workflow

1. Fill in **Author**, **Version**, **Review Date**, and **SOP Title** (e.g. `VA01 – Create Sales Order`).
2. Click **🔳 Compact** to minimize the app to a small toolbar (top-right of the screen).
3. Work in SAP GUI as usual.
4. Press **`Alt+C`** (or click **📷 Capture**) to capture a step.
5. Annotate the screenshot in the full-screen editor, then click **Done** (or **Skip**).
6. Repeat for each step in the process.
7. **💾 Save** your project as JSON so you can resume later.
8. **📘 Export** when ready to generate the Word document in the `exports/` folder.

---

## Project Structure

```
sop-builder/
├── main.js                 # Electron main process (window, capture, hotkeys)
├── paths.js                # App folders (dev vs packaged .exe)
├── package.json
├── src/
│   ├── index.html          # UI layout
│   ├── App.js              # App logic, steps, save/load
│   ├── annotator.js        # Canvas annotation tools
│   ├── exporter.js         # DOCX export
│   ├── sapgui-context.js   # SAP window / scripting context
│   ├── screenshot.js       # Screen capture
│   ├── config.js           # User settings
│   └── theme.css           # UI styling
├── scripts/
│   └── get-sapgui-context.ps1   # PowerShell helper for SAP context
├── data/
│   ├── settings.json       # Local settings (not in git — use settings.example.json)
│   └── sops/               # Saved SOP projects (.json)
├── screenshots/            # Captured PNGs (per project)
└── exports/                # Generated .docx files
```

---

## SAP GUI Scripting (Optional)

For automatic transaction codes and status bar text in step titles:

1. In SAP GUI: **Options → Accessibility & Scripting → Scripting**
2. Enable **Enable scripting**
3. Confirm security prompts when connecting

Without scripting, step titles still use the SAP window title from Windows APIs.

---

## Configuration

| Setting | Default | Location |
|---------|---------|----------|
| Capture hotkey | `Alt+C` | Settings modal or `data/settings.json` |
| Last opened project | — | Stored in `data/settings.json` |

Example `data/settings.json`:

```json
{
  "captureHotkey": "Alt+C",
  "lastProjectPath": ""
}
```

---

## Saving and Sharing SOPs

- **Project file:** Save as `.json` via **💾 Save**. Share this file with teammates; they can **📂 Open** it in their copy of SOP Builder.
- **Screenshots:** Stored locally in `screenshots/` with paths referenced in the JSON. When sharing projects, zip the JSON **and** the referenced screenshot files, or save/export from the same machine.
- **Word export:** **📘 Export** writes to `exports/<SOP Title>.docx` — suitable for review and distribution.

---

## Troubleshooting

| Issue | What to try |
|-------|-------------|
| App won't start | Run `npm install` again; ensure Node.js 18+ is installed |
| Capture shows wrong window | Focus SAP GUI before capturing; use compact mode |
| Hotkey doesn't work | Change hotkey in Settings; avoid conflicts with SAP/other tools |
| Empty step title | Enable SAP GUI Scripting, or type the description before capture |
| Export fails | Ensure each step has a valid screenshot path on disk |

---

## Development

```powershell
npm start          # run from source
npm run dist       # build Windows installer + portable exe
```

Main technologies: **Electron**, **electron-builder**, **docx**, **screenshot-desktop**, **SortableJS**.

---

## License

What changed
Packaging is now configured with electron-builder. I built the Windows installers on your machine:

File	Use
dist\SOP Builder Setup 1.0.0.exe
Installer — teammates double-click, no Node.js
dist\SOP Builder 1.0.0.exe
Portable — single file, no install
dist\win-unpacked\SOP Builder.exe
Unpacked app (good for quick testing)
For your team (no npm start)
Share SOP Builder Setup 1.0.0.exe or the portable exe via Teams/SharePoint. They install or run it like any Windows app.

Windows may show a SmartScreen warning because the app isn't code-signed yet — that's normal for internal tools. Click More info → Run anyway.

For you (when you change code)
Build a fresh exe once per release:

npm install
npm run dist
Output goes to the dist/ folder.

Why npm start existed before
npm start = developer mode (needs Node.js + source code)
npm run dist = packages Electron + your app into a standalone .exe for everyone else
The README is updated with both paths. Data folders (data/, screenshots/, exports/) are created next to the installed exe, so projects persist between runs.

Try double-clicking dist\SOP Builder 1.0.0.exe — it should open without Node.js. If you want, I can push these packaging changes to GitHub and attach the exe as a Release for the team to download.
