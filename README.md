# SOP Builder

**SOP Builder** is a Windows desktop app for capturing SAP GUI screens (or any other application), building step-by-step Standard Operating Procedures (SOPs), annotating screenshots, and exporting polished Word documents. Built with Electron for the **Business Core Solutions** team.

Think of it as a lightweight Scribe-style tool focused on SAP process documentation—with optional capture for browsers, Office, Calculator, and other Windows apps.

---

## Features

| Feature | Description |
|---------|-------------|
| **SAP GUI capture** | Focus SAP, crop to the SAP window, read fields via VBScript (SAP GUI Scripting), human-readable labels from tooltips |
| **Non-SAP capture** | Document any active window—Chrome, Edge, Calculator, Notepad, etc.—full desktop screenshot and window title only |
| **Capture mode** | Choose **SAP GUI** or **Non-SAP** in SOP Details (one row with Author, Version, Date) |
| **Status bar steps** | After Save/Post, SAP success messages become the step description (no field list) |
| **Session recorder** | `Ctrl+Shift+S` captures screenshot + SAP field deltas with **Set**-style descriptions |
| **Global hotkeys** | `Alt+C` capture; `Ctrl+Shift+S` session step (configurable in Settings) |
| **Compact mode** | Shrinks to a floating toolbar so the target app stays in focus |
| **Annotation tools** | Box, arrow, highlight, numbered callouts, and text labels |
| **Step management** | Drag to reorder, edit descriptions, delete steps |
| **Save / Open** | Projects saved as JSON (`.json`) under `data/sops/` |
| **DOCX export** | Word document with metadata; multi-field steps break onto separate lines |
| **Branding & templates** | Reusable company logo, name, accent color, and footer; cover page, Document Control table, headers/footers with page numbers — applied to every export |
| **Document templates** | **Standard** (cover + headers/footers), **Detailed** (adds Revision History), **Minimal** (header only, no cover) |
| **SOP metadata** | Author, version, review date (defaults to today), title, plus optional Document ID, Department, and Classification |

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
- **SAP GUI for Windows** (for SAP capture mode)
- **Optional:** SAP GUI Scripting enabled for field capture and richer titles

### 1. Install Node.js

1. Go to [https://nodejs.org/](https://nodejs.org/)
2. Download and install the **LTS** version
3. Open **PowerShell** or **Command Prompt** and verify:

```powershell
node --version
npm --version
```

### 2. Clone the repository

```powershell
git clone https://github.com/bhimsunnyraj-bolluru/sop-builder.git
cd sop-builder
```

### 3. Install dependencies

```powershell
npm install
```

### 4. First-run setup (optional)

```powershell
copy data\settings.example.json data\settings.json
```

Open **Settings** (⚙️) for hotkeys and max SAP fields per capture.

### 5. Run from source

```powershell
npm start
```

---

## Quick Start Workflow

1. Fill in **Author**, **Version**, **Review Date**, and **SOP Title**.
2. Under the same row, choose **SAP GUI** or **Non-SAP** (saved automatically).
3. Click **🔳 Compact** to minimize the app to a small toolbar.
4. **SAP GUI:** work in SAP; the app brings SAP to the front when you capture from SOP Builder.  
   **Non-SAP:** click the app you want (e.g. Chrome), then capture—SAP is not focused or read.
5. Press **`Alt+C`** or **`Ctrl+Shift+S`** (session steps with SAP field deltas in SAP GUI mode only).
6. Annotate the screenshot, then click **Done** (or **Skip**).
7. **💾 Save** the project JSON; **📘 Export** for Word in `exports/`.

---

## Capture modes (SAP GUI vs Non-SAP)

| Mode | Select in SOP Details | Behavior |
|------|------------------------|----------|
| **SAP GUI** | SAP GUI | Brings SAP to the front; crops screenshot to SAP window; captures up to N fields (Settings); status bar success text only on save steps |
| **Non-SAP** | Non-SAP | Uses whichever window is active after SOP hides; **full desktop** screenshot; step title = window title; **no** SAP fields or scripting |

Switch modes when your SOP mixes SAP and other apps (e.g. Fiori in Chrome vs VA01 in SAP GUI).

**Tips**

- **Non-SAP:** focus Chrome/Edge/Calculator **before** `Alt+C`—do not click Capture in SOP while that app is in the background unless it is the window behind SOP.
- **SAP GUI:** capture right after **Save** while the green status message is visible for a clean status-bar step.
- Field steps use descriptions like `Set Sales Organization: P100` (semicolon-separated in the UI; separate lines in Word export).

---

## Project Structure

```
SapSOPBuilder_Full/
├── main.js                 # Electron main process (capture, hotkeys, IPC)
├── paths.js                # App folders (dev vs packaged .exe)
├── package.json
├── src/
│   ├── index.html          # UI layout
│   ├── App.js              # Steps, save/load, capture UI
│   ├── annotator.js        # Canvas annotation tools
│   ├── exporter.js         # DOCX export
│   ├── sapgui-context.js   # SAP / foreground window context
│   ├── screenshot.js       # Screen capture + optional crop
│   ├── config.js           # User settings
│   ├── captureLog.js       # Capture diagnostics log
│   ├── focus-sap.js        # Bring SAP window to front
│   ├── sap-powershell.js   # PowerShell executable helper
│   └── modules/
│       ├── sap/
│       │   ├── snapshotEngine.js   # VBScript SAP snapshot
│       │   ├── deltaEngine.js      # Compare snapshots
│       │   └── labelResolver.js    # Field labels from capture
│       └── recording/
│           └── sessionRecorder.js  # Session step pipeline
├── scripts/
│   ├── get-sapgui-context.ps1      # Window context (-CaptureMode Sap | NonSap)
│   ├── capture-sap-snapshot.vbs    # Primary SAP field capture (VBScript)
│   ├── capture-sap-snapshot.ps1    # PowerShell fallback
│   ├── focus-sap-window.ps1
│   └── crop-screenshot.ps1
├── data/
│   ├── settings.json       # Local (gitignored; see settings.example.json)
│   ├── snapshot-last.json  # Last capture debug (gitignored)
│   ├── capture.log         # Capture log (gitignored)
│   └── sops/               # Saved SOP projects
├── screenshots/
└── exports/
```

---

## SOP Session Recorder (`Ctrl+Shift+S`)

Available in **SAP GUI** mode only.

| Hotkey | Mode | What it does |
|--------|------|----------------|
| **`Alt+C`** | Capture | Screenshot + title (SAP fields or window title) |
| **`Ctrl+Shift+S`** | Session | Same + compare to previous SAP snapshot for field changes |

Session steps use **Set** phrasing, e.g. `Set Sales Document Type: OR; Set Sales Organization: P100`.

The **first** session step is a baseline. **New SOP** or opening a saved project resets the baseline. Switching from SAP GUI to Non-SAP clears the SAP field baseline.

---

## SAP context detection (SAP GUI mode)

| Source | What you get |
|--------|----------------|
| **VBScript snapshot** | Field values, labels (DefaultTooltip), transaction, screen |
| **PowerShell / Win32** | SAP window title, rect for crop, status bar (UI Automation) |
| **Success status bar** | Step description = status message only (e.g. `Standard Order 6677 has been saved.`) |

### Enable SAP GUI Scripting

**On your PC:** SAP GUI → **Options → Accessibility & Scripting → Scripting** → enable scripting.

**On the server:** `sapgui/user_scripting` and user authorizations (ask Basis). VBScript capture needs a working scripting engine; if PowerShell `GetScriptingEngine` fails on your PC, VBScript may still work via `cscript`.

---

## Configuration

| Setting | Default | Where |
|---------|---------|--------|
| Capture target | SAP GUI | **SOP Details** row (SAP GUI / Non-SAP radios) |
| Capture hotkey | `Alt+C` | Settings (⚙️) |
| Session hotkey | `Ctrl+Shift+S` | Settings |
| Max SAP fields per capture | `10` | Settings (1–50) |
| Last opened project | — | `data/settings.json` |

Example `data/settings.json`:

```json
{
  "captureHotkey": "Alt+C",
  "sessionHotkey": "CommandOrControl+Shift+S",
  "captureTarget": "sap",
  "maxCaptureFields": 10,
  "lastProjectPath": "",
  "branding": {
    "companyName": "Business Core Solutions",
    "logoPath": "",
    "accentColor": "#F97316",
    "footerText": "Confidential — for internal use only",
    "template": "standard"
  }
}
```

Diagnostics after each capture: **Settings → Open data folder** → `capture.log`, `snapshot-last.json`.

---

## Branding & Templates

Set your company branding **once** in **Settings (⚙️) → Branding & Template**; it is reused on every Word export.

| Setting | What it does |
|---------|----------------|
| **Company name** | Shown on the cover page, header, and footer (default: *Business Core Solutions*) |
| **Document template** | **Standard** = cover page + headers/footers + Document Control table; **Detailed** = adds a Revision History table; **Minimal** = simple header only, no cover page |
| **Accent color** | Color for the title rule, section headings, and header line |
| **Footer / confidentiality text** | Left side of every page footer (page number is on the right) |
| **Logo** | PNG/JPG/GIF/BMP; copied into `data/branding/` so it travels with your data folder. Appears centered on the cover page |

Per-SOP fields in **SOP Details** — **Document ID**, **Department**, and **Classification** — populate the cover page's **Document Control** table when filled in. They are saved inside each project's `.json`.

> Branding lives in `data/settings.json` under a `branding` key, so the whole team can share one branded template. Department is remembered when you start a New SOP.

---

## Saving and Sharing SOPs

- **Project file:** **💾 Save** as `.json`; share with teammates via **📂 Open**.
- **Screenshots:** Local `screenshots/` paths in JSON—zip JSON + images when moving machines.
- **Word export:** **📘 Export** → `exports/<SOP Title>.docx`.

---

## Troubleshooting

| Issue | What to try |
|-------|-------------|
| App won't start | `npm install`; Node.js 18+; check terminal for errors |
| Wrong title (Chrome vs SAP) | Set **SAP GUI** or **Non-SAP** in SOP Details; for Non-SAP, focus the target app before capture |
| SAP shows Chrome title | Switch to **SAP GUI**; capture from SOP or focus SAP before `Alt+C` |
| Non-SAP shows SAP fields | Switch to **Non-SAP**; do not use session recorder for browser-only steps |
| Empty SAP fields | Enable scripting; keep SAP focused; increase max fields in Settings |
| Capture timed out | Keep SAP focused; reduce open sessions; check `data/snapshot-last.json` |
| Hotkey conflict | Change hotkey in Settings |
| Export fails | Ensure screenshot files exist on disk |

---

## Development

```powershell
npm start          # run from source
npm run dist       # build Windows installer + portable exe
```

Main technologies: **Electron**, **electron-builder**, **docx**, **screenshot-desktop**, **SortableJS**, **SAP GUI Scripting** (VBScript + PowerShell).

---

## License

Internal use — Business Core Solutions. Contact your team lead for distribution and usage guidelines.
