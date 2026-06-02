const {app,BrowserWindow,ipcMain,globalShortcut,dialog} = require("electron");
const path=require("path");
const fs=require("fs");
const {takeScreenshot}=require("./src/screenshot");
const {mergeContextFromSnapshotAndPs,getSapGuiContext,isSapForegroundContext,isSapSuccessStatusBar,formatStepDescription,describeCaptureContext}=require("./src/sapgui-context");
const {loadSettings,saveSettings,DEFAULT_SETTINGS,normalizeCaptureTarget}=require("./src/config");
const {getProjectRoot,getScreenshotsDir,getSopsDir,getDataDir,ensureDir}=require("./paths");
const {captureSnapshot,emptySnapshot}=require("./src/modules/sap/snapshotEngine");
const {recordStep,resetSessionRecorder}=require("./src/modules/recording/sessionRecorder");
const {appendCaptureLog,getLogPath,getSnapshotDebugPath}=require("./src/captureLog");
const {focusSapWindow}=require("./src/focus-sap");

let mainWindow;
let _normalBounds = null;
let _registeredHotkey = null;
let _registeredSessionHotkey = null;
let _isCompact = false;
let _wasCompactBeforeAnnotate = false;

function createWindow(){
 mainWindow=new BrowserWindow({
  width:1400,
  height:900,
  show:false,
  title:"SOP Builder",
  backgroundColor:'#08060D',
  webPreferences:{nodeIntegration:true,contextIsolation:false}
 });
 mainWindow.loadFile("src/index.html");
 mainWindow.once("ready-to-show", ()=>{ mainWindow.show(); });
}

async function runFullCapture(){
 if(!mainWindow) throw new Error("Main window not available");
 const t0 = Date.now();
 let phase = t0;
 const timings = {};
 const screenshotsDir=ensureDir(getScreenshotsDir());
 const settings = loadSettings(getProjectRoot());
 const maxFields = settings.maxCaptureFields ?? DEFAULT_SETTINGS.maxCaptureFields;
 const captureTarget = normalizeCaptureTarget(settings.captureTarget);

 const { wasVisible, wasFocused } = await prepareCaptureWindow();
 timings.prepareMs = Date.now() - phase; phase = Date.now();

 const captureSap = captureTarget === "sap";
 if (captureSap) {
  await focusSapWindow();
  await new Promise((r) => setTimeout(r, 150));
 }

 const psContext = await getSapGuiContext({ captureTarget });
 timings.contextMs = Date.now() - phase; phase = Date.now();

 let snapshot = emptySnapshot();
 if (captureSap) {
  snapshot = await captureSnapshot({ maxFields });
 } else {
  snapshot = { ...emptySnapshot(), captureSource: "win32-foreground" };
 }
 timings.snapshotMs = Date.now() - phase; phase = Date.now();

 if(snapshot.debug && snapshot.debug.discoverMs != null){
  timings.vbsDiscoverMs = snapshot.debug.discoverMs;
 }

 const context = mergeContextFromSnapshotAndPs(snapshot, psContext, captureTarget);
 const statusBarOnly = isSapSuccessStatusBar(context.statusBar);
 if (statusBarOnly) {
  snapshot = { ...snapshot, controls: [] };
 }
 const cropRect = captureSap
  ? (context.sapWindowRect || snapshot.sapWindowRect || null)
  : null;

 const file=path.join(screenshotsDir, `${Date.now()}.png`);
 await takeScreenshot(file, cropRect);
 timings.screenshotMs = Date.now() - phase; phase = Date.now();
 const cropped = !!cropRect;

 timings.ocrMs = 0;

 await restoreCaptureWindow(wasVisible, wasFocused);
 timings.restoreMs = Date.now() - phase; phase = Date.now();

 const step = recordStep({ screenshot: file, snapshot, context });
 timings.recordMs = Date.now() - phase;
 timings.totalMs = Date.now() - t0;
 const controlCount = statusBarOnly ? 0 : (Array.isArray(snapshot.controls) ? snapshot.controls.length : 0);
 const changeCount = Array.isArray(step.changes) ? step.changes.length : 0;

 setImmediate(() => {
  appendCaptureLog({
   event: "capture-complete",
   contextSource: context.source || "none",
   captureSource: snapshot.captureSource || "",
   controlCount,
   changeCount,
   cropped,
   maxCaptureFields: maxFields,
   captureTarget,
   statusBar: context.statusBar || undefined,
   statusBarOnly,
   snapshotError: snapshot.error || undefined,
   debug: snapshot.debug || undefined,
   timings,
  });
 });

 return {
  file,
  title: step.description,
  context,
  contextDetail: describeCaptureContext(context),
  step,
  snapshotControlCount: controlCount,
  snapshotError: snapshot.error || null,
  logPath: getLogPath(),
  snapshotDebugPath: getSnapshotDebugPath(),
  dataDir: getDataDir(),
  timings,
 };
}

async function runCapture(){
 return runFullCapture();
}

async function runSessionCapture(){
 return runFullCapture();
}

async function prepareCaptureWindow(){
 if(!mainWindow) throw new Error("Main window not available");
 const wasVisible = mainWindow.isVisible();
 const wasFocused = mainWindow.isFocused();
 if(wasVisible){
  try{ mainWindow.hide(); }catch(e){}
 }
 await new Promise((r) => setTimeout(r, 200));
 return { wasVisible, wasFocused };
}

async function restoreCaptureWindow(wasVisible, wasFocused){
 if(!mainWindow || !wasVisible) return;
 try{ mainWindow.show(); if(!wasFocused) mainWindow.blur(); else mainWindow.focus(); }catch(e){}
}

function registerSessionHotkey(accelerator){
 if(_registeredSessionHotkey){
  try{ globalShortcut.unregister(_registeredSessionHotkey); }catch(e){}
  _registeredSessionHotkey = null;
 }
 if(!accelerator) return { ok:false, error:"No session hotkey specified" };
 if(!globalShortcut.isRegistered(accelerator)){
  const ok = globalShortcut.register(accelerator, async ()=>{
   if(!mainWindow) return;
   try{
    const result = await runSessionCapture();
    mainWindow.webContents.send("session-step-complete", result);
   }catch(err){
    mainWindow.webContents.send("session-step-error", err.message || String(err));
   }
  });
  if(ok){
   _registeredSessionHotkey = accelerator;
   return { ok:true, hotkey: accelerator };
  }
  return { ok:false, error:`Could not register session hotkey "${accelerator}".` };
 }
 return { ok:false, error:`"${accelerator}" is already registered.` };
}

function registerCaptureHotkey(accelerator){
 if(_registeredHotkey){
  try{ globalShortcut.unregister(_registeredHotkey); }catch(e){}
  _registeredHotkey = null;
 }
 if(!accelerator) return { ok:false, error:"No hotkey specified" };
 if(!globalShortcut.isRegistered(accelerator)){
  const ok = globalShortcut.register(accelerator, async ()=>{
   if(!mainWindow) return;
   try{
    const result = await runCapture();
    mainWindow.webContents.send("capture-complete", result);
   }catch(err){
    mainWindow.webContents.send("capture-error", err.message || String(err));
   }
  });
  if(ok){
   _registeredHotkey = accelerator;
   return { ok:true, hotkey: accelerator };
  }
  return { ok:false, error:`Could not register "${accelerator}". It may be in use by another app.` };
 }
 return { ok:false, error:`"${accelerator}" is already registered.` };
}

function applyCompactMode(enable){
 if(!mainWindow) throw new Error("Main window not available");
 _isCompact = !!enable;
 if(enable){
  try{ _normalBounds = mainWindow.getBounds(); }catch(e){ _normalBounds = null; }
  mainWindow.setAlwaysOnTop(true, 'floating');
  mainWindow.setMenuBarVisibility(false);
  mainWindow.setResizable(false);
  mainWindow.setMinimumSize(200, 50);
  const compactWidth = 480;
  const compactHeight = 110;
  mainWindow.setSize(compactWidth, compactHeight);
  try{
   const {screen} = require('electron');
   const display = screen.getPrimaryDisplay();
   const area = display.workArea;
   const marginRight = 16;
   const marginTop = 52;
   const posX = Math.round(area.x + area.width - compactWidth - marginRight);
   const posY = Math.round(area.y + marginTop);
   mainWindow.setPosition(Math.max(area.x, posX), Math.max(area.y, posY));
  }catch(e){/* ignore */}
 } else {
  mainWindow.setAlwaysOnTop(false);
  mainWindow.setMenuBarVisibility(true);
  mainWindow.setResizable(true);
  mainWindow.setMinimumSize(800, 600);
  if(_normalBounds){
   try{ mainWindow.setBounds(_normalBounds); }catch(e){/* ignore */}
   _normalBounds = null;
  } else {
   mainWindow.setSize(1400, 900);
   mainWindow.center();
  }
 }
 return _isCompact;
}

function enterAnnotateMode(){
 if(!mainWindow) return;
 _wasCompactBeforeAnnotate = _isCompact;
 mainWindow.setAlwaysOnTop(false);
 mainWindow.setMenuBarVisibility(true);
 mainWindow.setResizable(true);
 mainWindow.setMinimumSize(900, 650);
 try{
  const {screen} = require('electron');
  const display = screen.getPrimaryDisplay();
  const {width, height} = display.workAreaSize;
  mainWindow.setBounds({
   x: Math.round(width * 0.03),
   y: Math.round(height * 0.03),
   width: Math.round(width * 0.94),
   height: Math.round(height * 0.94),
  });
 }catch(e){
  mainWindow.maximize();
 }
 mainWindow.show();
 mainWindow.focus();
}

function exitAnnotateMode(){
 if(!mainWindow) return;
 if(_wasCompactBeforeAnnotate){
  applyCompactMode(true);
 } else {
  mainWindow.setMinimumSize(800, 600);
  if(!mainWindow.isMaximized()){
   mainWindow.setSize(1400, 900);
   mainWindow.center();
  }
 }
 _wasCompactBeforeAnnotate = false;
}

ipcMain.handle("capture-sapgui", async () => runCapture());
ipcMain.handle("capture-session-step", async () => runSessionCapture());
ipcMain.handle("reset-session-recorder", async () => { resetSessionRecorder(); return true; });
ipcMain.handle("get-settings", async () => loadSettings(getProjectRoot()));
ipcMain.handle("save-settings", async (_event, settings) => {
 // Merge over the currently-saved settings so a partial save (e.g. just the
 // hotkey) never wipes other persisted values like captureTarget or branding.
 const current = loadSettings(getProjectRoot());
 const incoming = settings || {};
 const next = { ...current, ...incoming };
 if (incoming.branding) next.branding = { ...current.branding, ...incoming.branding };
 const merged = saveSettings(getProjectRoot(), next);
 const reg = registerCaptureHotkey(merged.captureHotkey);
 const sessionReg = registerSessionHotkey(merged.sessionHotkey || "CommandOrControl+Shift+S");
 return { settings: merged, hotkey: reg, sessionHotkey: sessionReg };
});

ipcMain.handle("pick-logo-file", async () => {
 if(!mainWindow) throw new Error("Main window not available");
 const result = await dialog.showOpenDialog(mainWindow, {
  title: "Choose company logo",
  filters: [{ name: "Images", extensions: ["png", "jpg", "jpeg", "gif", "bmp"] }],
  properties: ["openFile"],
 });
 if(result.canceled || !result.filePaths || !result.filePaths.length) return { canceled: true };
 const src = result.filePaths[0];
 // Copy the logo into data/branding so it travels with the data folder.
 try{
  const brandingDir = ensureDir(path.join(getDataDir(), "branding"));
  const ext = (path.extname(src) || ".png").toLowerCase();
  const dest = path.join(brandingDir, "logo" + ext);
  fs.copyFileSync(src, dest);
  return { ok: true, path: dest };
 }catch(e){
  // Fall back to referencing the original location if the copy fails.
  return { ok: true, path: src };
 }
});
ipcMain.handle("set-compact-mode", async (_event, enable) => applyCompactMode(enable));
ipcMain.handle("enter-annotate-mode", async () => { enterAnnotateMode(); return true; });
ipcMain.handle("exit-annotate-mode", async () => { exitAnnotateMode(); return { compact: _isCompact }; });

function sanitizeFileName(name){
 return (name || "SOP").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").slice(0, 80);
}

function getSopsDirEnsured(){
 return ensureDir(getSopsDir());
}

ipcMain.handle("save-project-dialog", async (_event, payload) => {
 if(!mainWindow) throw new Error("Main window not available");
 const project = payload && payload.project ? payload.project : payload;
 const suggested = sanitizeFileName(project && project.title) + ".json";
 const existingPath = payload && payload.filePath ? payload.filePath : null;

 if(existingPath && fs.existsSync(path.dirname(existingPath))){
  fs.writeFileSync(existingPath, JSON.stringify(project, null, 2));
  const settings = loadSettings(getProjectRoot());
  saveSettings(getProjectRoot(), { ...settings, lastProjectPath: existingPath });
  return { ok: true, filePath: existingPath, fileName: path.basename(existingPath) };
 }

 const result = await dialog.showSaveDialog(mainWindow, {
  title: "Save SOP Project",
  defaultPath: path.join(getSopsDirEnsured(), suggested),
  filters: [{ name: "SOP Project", extensions: ["json"] }],
 });
 if(result.canceled || !result.filePath) return { ok: false, canceled: true };

 let filePath = result.filePath;
 if(!filePath.toLowerCase().endsWith(".json")) filePath += ".json";
 fs.writeFileSync(filePath, JSON.stringify(project, null, 2));
 const settings = loadSettings(getProjectRoot());
 saveSettings(getProjectRoot(), { ...settings, lastProjectPath: filePath });
 return { ok: true, filePath, fileName: path.basename(filePath) };
});

ipcMain.handle("load-project-dialog", async () => {
 if(!mainWindow) throw new Error("Main window not available");
 const settings = loadSettings(getProjectRoot());
 const defaultPath = (settings.lastProjectPath && fs.existsSync(settings.lastProjectPath))
  ? path.dirname(settings.lastProjectPath)
  : getSopsDirEnsured();

 const result = await dialog.showOpenDialog(mainWindow, {
  title: "Open SOP Project",
  defaultPath,
  filters: [{ name: "SOP Project", extensions: ["json"] }],
  properties: ["openFile"],
 });
 if(result.canceled || !result.filePaths || !result.filePaths.length){
  return { ok: false, canceled: true };
 }

 const filePath = result.filePaths[0];
 const project = JSON.parse(fs.readFileSync(filePath, "utf8"));
 saveSettings(getProjectRoot(), { ...settings, lastProjectPath: filePath });
 return { ok: true, filePath, fileName: path.basename(filePath), project };
});

ipcMain.handle("load-last-project", async () => {
 const settings = loadSettings(getProjectRoot());
 const filePath = settings.lastProjectPath;
 if(!filePath || !fs.existsSync(filePath)) return { ok: false };
 const project = JSON.parse(fs.readFileSync(filePath, "utf8"));
 return { ok: true, filePath, fileName: path.basename(filePath), project };
});

ipcMain.handle("get-data-paths", async () => ({
 dataDir: getDataDir(),
 projectRoot: getProjectRoot(),
 snapshotDebugPath: getSnapshotDebugPath(),
 logPath: getLogPath(),
}));

ipcMain.handle("open-data-folder", async () => {
 const { shell } = require("electron");
 ensureDir(getDataDir());
 await shell.openPath(getDataDir());
 return { ok: true, path: getDataDir() };
});

app.whenReady().then(()=>{
 ensureDir(getDataDir());
 createWindow();
 const settings = loadSettings(getProjectRoot());
 registerCaptureHotkey(settings.captureHotkey);
 registerSessionHotkey(settings.sessionHotkey || "CommandOrControl+Shift+S");
});
app.on("will-quit", ()=>{ globalShortcut.unregisterAll(); });
app.on("window-all-closed", () => {
 if(process.platform !== "darwin") app.quit();
});
