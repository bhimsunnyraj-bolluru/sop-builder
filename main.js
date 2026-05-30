const {app,BrowserWindow,ipcMain,globalShortcut,dialog} = require("electron");
const path=require("path");
const fs=require("fs");
const {takeScreenshot}=require("./src/screenshot");
const {getSapGuiContext,formatStepDescription}=require("./src/sapgui-context");
const {loadSettings,saveSettings}=require("./src/config");

let mainWindow;
let _normalBounds = null;
let _registeredHotkey = null;
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

async function runCapture(){
 if(!mainWindow) throw new Error("Main window not available");
 const screenshotsDir=path.join(__dirname, "screenshots");
 if(!fs.existsSync(screenshotsDir)) fs.mkdirSync(screenshotsDir,{recursive:true});
 const wasVisible = mainWindow.isVisible();
 const wasFocused = mainWindow.isFocused();
 if(wasVisible && wasFocused){
  try{ mainWindow.hide(); }catch(e){}
  await new Promise((resolve) => setTimeout(resolve, 500));
 }
 const context = await getSapGuiContext();
 const file=path.join(screenshotsDir, `${Date.now()}.png`);
 await takeScreenshot(file);
 if(wasVisible){
  try{ mainWindow.show(); if(!wasFocused) mainWindow.blur(); else mainWindow.focus(); }catch(e){}
 }
 const title = formatStepDescription(context);
 return { file, title, context };
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
ipcMain.handle("get-settings", async () => loadSettings(__dirname));
ipcMain.handle("save-settings", async (_event, settings) => {
 const merged = saveSettings(__dirname, settings || {});
 const reg = registerCaptureHotkey(merged.captureHotkey);
 return { settings: merged, hotkey: reg };
});
ipcMain.handle("set-compact-mode", async (_event, enable) => applyCompactMode(enable));
ipcMain.handle("enter-annotate-mode", async () => { enterAnnotateMode(); return true; });
ipcMain.handle("exit-annotate-mode", async () => { exitAnnotateMode(); return { compact: _isCompact }; });

function sanitizeFileName(name){
 return (name || "SOP").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").slice(0, 80);
}

function getSopsDir(){
 const dir = path.join(__dirname, "data", "sops");
 if(!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
 return dir;
}

ipcMain.handle("save-project-dialog", async (_event, payload) => {
 if(!mainWindow) throw new Error("Main window not available");
 const project = payload && payload.project ? payload.project : payload;
 const suggested = sanitizeFileName(project && project.title) + ".json";
 const existingPath = payload && payload.filePath ? payload.filePath : null;

 if(existingPath && fs.existsSync(path.dirname(existingPath))){
  fs.writeFileSync(existingPath, JSON.stringify(project, null, 2));
  const settings = loadSettings(__dirname);
  saveSettings(__dirname, { ...settings, lastProjectPath: existingPath });
  return { ok: true, filePath: existingPath, fileName: path.basename(existingPath) };
 }

 const result = await dialog.showSaveDialog(mainWindow, {
  title: "Save SOP Project",
  defaultPath: path.join(getSopsDir(), suggested),
  filters: [{ name: "SOP Project", extensions: ["json"] }],
 });
 if(result.canceled || !result.filePath) return { ok: false, canceled: true };

 let filePath = result.filePath;
 if(!filePath.toLowerCase().endsWith(".json")) filePath += ".json";
 fs.writeFileSync(filePath, JSON.stringify(project, null, 2));
 const settings = loadSettings(__dirname);
 saveSettings(__dirname, { ...settings, lastProjectPath: filePath });
 return { ok: true, filePath, fileName: path.basename(filePath) };
});

ipcMain.handle("load-project-dialog", async () => {
 if(!mainWindow) throw new Error("Main window not available");
 const settings = loadSettings(__dirname);
 const defaultPath = (settings.lastProjectPath && fs.existsSync(settings.lastProjectPath))
  ? path.dirname(settings.lastProjectPath)
  : getSopsDir();

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
 saveSettings(__dirname, { ...settings, lastProjectPath: filePath });
 return { ok: true, filePath, fileName: path.basename(filePath), project };
});

ipcMain.handle("load-last-project", async () => {
 const settings = loadSettings(__dirname);
 const filePath = settings.lastProjectPath;
 if(!filePath || !fs.existsSync(filePath)) return { ok: false };
 const project = JSON.parse(fs.readFileSync(filePath, "utf8"));
 return { ok: true, filePath, fileName: path.basename(filePath), project };
});

app.whenReady().then(()=>{
 createWindow();
 const settings = loadSettings(__dirname);
 registerCaptureHotkey(settings.captureHotkey);
});
app.on("will-quit", ()=>{ globalShortcut.unregisterAll(); });
app.on("window-all-closed", () => {
 if(process.platform !== "darwin") app.quit();
});
