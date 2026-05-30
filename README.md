# SOP Builder

**SOP Builder** is a Windows desktop app for capturing SAP GUI screens, building step-by-step Standard Operating Procedures (SOPs), annotating screenshots, and exporting polished Word documents. Built with Electron for the **Business Core Solutions** team.

Think of it as a lightweight Scribe-style tool focused on SAP process documentation.

---

## Features

| Feature | Description |
|---------|-------------|
| **Screenshot capture** | Capture the active SAP GUI window while you work |
| **SAP context detection** | Auto-fills step titles from SAP window name, status bar (UI Automation/OCR), and transaction code when SAP Scripting is enabled |
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

## SAP context detection

When you capture a step, SOP Builder reads SAP metadata and fills the step description automatically.

| Source | What you get | When |
|--------|----------------|------|
| **Win32** | SAP window title (e.g. `Create Standard Order: Overview`) | Always |
| **UI Automation** | Status bar (e.g. `Standard Order 6676 has been saved.`) | Usually — no SAP Scripting required |
| **Screenshot OCR** | Status bar text from the captured image | Fallback if UI Automation misses it |
| **SAP Scripting** | Window title + **t-code** (e.g. `VA01`) + **system** + status bar | SAP GUI + server scripting enabled |

Example step title with scripting enabled:

`Create Sales Documents (VA01) [TS4] — Standard Order 4500012345 has been saved`

The text after `—` comes from the **SAP GUI status bar** (bottom of the screen), including messages like a newly created sales order number after you click **Save** in VA01.

**Workflow for VA01 save step:** click **Save** in SAP → wait for the status bar message → press **`Alt+C`** (or **Capture**) while SAP is still showing that message.

After capture, the status bar shows what was detected (e.g. `UI Automation: window title, status bar` or `SAP Scripting: window title, t-code VA01, status bar`). Each step shows a meta line under the description (e.g. `status · status detected` or `VA01 · scripting`).

### Enable SAP GUI Scripting

**On your PC (SAP Logon / SAP GUI):**

1. Open **Options → Accessibility & Scripting → Scripting**
2. Check **Enable scripting**
3. Set security to allow scripting (confirm prompts when SOP Builder captures)

**On the SAP system (often required — ask your SAP Basis team):**

- Parameter `sapgui/user_scripting` must be `TRUE` on the application server
- Your user may need scripting allowed in transaction **RZ11** / user profile settings

Without server-side scripting, you still get the **window title only** (`source: win32` in saved projects).

### Tips

- Click the SAP screen you want **before** pressing `Alt+C` or **Capture**
- Use **Compact** mode so SOP Builder hides before the screenshot
- Type your own text in **Step description** before capture to override auto-detection

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
| Empty step title | Focus SAP before capture; enable SAP GUI Scripting for t-code/status bar; or type Step description first |
| Only window title, no status bar | Capture right after Save while the green status message is visible; restart app for latest detection |
| Only window title, no t-code | SAP Scripting not enabled on client or server — see [SAP context detection](#sap-context-detection) |
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

Internal use — Business Core Solutions. Contact your team lead for distribution and usage guidelines.
