const {ipcRenderer}=require("electron");
const fs=require("fs");
const path=require("path");
const { pathToFileURL } = require('url');
const { getDataDir, getExportsDir, ensureDir } = require("../paths");
const {
	initAnnotator,
	openAnnotator,
	closeAnnotator: closeAnnotatorModal,
	saveAnnotation,
	undoAnnotation,
	clearAnnotations,
} = require("./annotator");

let _sortableLib = null;
let _confirmCallback = null;

let project={
	title:"",
	author:"",
	version:"1.0",
	reviewDate: todayDateIso(),
	steps:[]
};
let _compactMode = false;
let _sortable = null;
let _captureHotkey = "Alt+C";
let _pendingAnnotateIndex = -1;
let _inAnnotateFlow = false;
let _currentProjectPath = "";

function setStatus(message, type="info"){
 const status=document.getElementById("status");
 if(!status) return;
 status.textContent=message;
 status.className = "status-" + (type === "error" ? "error" : type === "success" ? "success" : type === "loading" ? "loading" : "info");
}

function todayDateIso(){
 const d = new Date();
 const pad = (n)=> String(n).padStart(2, "0");
 return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}

function syncProjectFromUI(){
 project.title = document.getElementById("title").value.trim();
 project.author = document.getElementById("author").value.trim();
 project.version = document.getElementById("version").value.trim() || "1.0";
 project.reviewDate = document.getElementById("reviewDate").value || todayDateIso();
}

function syncUIToProject(){
 document.getElementById("title").value = project.title || "";
 document.getElementById("author").value = project.author || "";
 document.getElementById("version").value = project.version || "1.0";
 if(!project.reviewDate) project.reviewDate = todayDateIso();
 document.getElementById("reviewDate").value = project.reviewDate;
}

function updateCaptureButtonTitle(){
 const label = `Capture screenshot (${displayHotkey(_captureHotkey)})`;
 ["captureBtn","captureBtnMain"].forEach((id)=>{
  const btn = document.getElementById(id);
  if(btn) btn.title = label;
 });
}

function displayHotkey(accelerator){
 return (accelerator || "").replace(/CommandOrControl/g, "Ctrl");
}

function getSortable(){
 if(!_sortableLib) _sortableLib = require("sortablejs");
 return _sortableLib;
}

function setBusy(busy){
 document.body.classList.toggle("ui-busy", !!busy);
}

function showConfirm(message, onOk, title="Confirm"){
 _confirmCallback = onOk;
 const modal = document.getElementById("confirmModal");
 const msg = document.getElementById("confirmMessage");
 const ttl = document.getElementById("confirmTitle");
 if(msg) msg.textContent = message;
 if(ttl) ttl.textContent = title;
 if(modal) modal.classList.add("open");
}

function closeConfirm(){
 _confirmCallback = null;
 document.getElementById("confirmModal")?.classList.remove("open");
}

function runConfirmOk(){
 const cb = _confirmCallback;
 closeConfirm();
 if(typeof cb === "function") cb();
}

function destroySortable(){
 if(_sortable){
  try{ _sortable.destroy(); }catch(e){}
  _sortable = null;
 }
}

function clearStepThumbnails(container){
 container.querySelectorAll("img.thumb").forEach((img)=>{ img.removeAttribute("src"); });
}

function render(){
 const d=document.getElementById("steps");
 if(!d) return;

 clearStepThumbnails(d);
 destroySortable();
 d.replaceChildren();

 if(!project.steps.length){
  const empty = document.createElement("div");
  empty.className = "empty-steps";
  empty.textContent = "No steps yet — use Capture or your hotkey while working in SAP.";
  d.appendChild(empty);
  return;
 }

 const frag = document.createDocumentFragment();
 project.steps.forEach((s,i)=>{
  const item = document.createElement("div");
  item.className = "step-item";
  item.dataset.index = String(i);
  item.innerHTML = `
    <div class="handle" title="Drag">≡</div>
    <div class="num">${i+1}</div>
    <div class="thumb-container"></div>
    <div class="desc">${escapeHtml(s.description||"")}</div>
    <div class="step-actions">
      <button type="button" data-step-action="annotate" data-index="${i}" title="Annotate screenshot">🖊️</button>
      <button type="button" data-step-action="edit" data-index="${i}" title="Edit description">✏️</button>
      <button type="button" data-step-action="delete" data-index="${i}" title="Delete">🗑️</button>
    </div>`;
  if(s.image){
   const thumbDiv = item.querySelector(".thumb-container");
   const img = document.createElement("img");
   img.className = "thumb";
   img.title = "Click to annotate";
   img.loading = "lazy";
   img.decoding = "async";
   img.dataset.index = String(i);
   img.dataset.stepAction = "annotate";
   try{ img.src = pathToFileURL(path.resolve(s.image)).href; }catch(e){ img.src = s.image; }
   thumbDiv.appendChild(img);
  }
  frag.appendChild(item);
 });
 d.appendChild(frag);

 try{
  _sortable = getSortable().create(d, {
   handle: ".handle",
   animation: 120,
   onEnd: (evt)=>{
    const from = evt.oldIndex; const to = evt.newIndex;
    if(from===to) return;
    const item = project.steps.splice(from,1)[0];
    project.steps.splice(to,0,item);
    requestAnimationFrame(render);
   },
  });
 }catch(e){}
}

function renderDeferred(){
 requestAnimationFrame(()=> render());
}

function escapeHtml(str){ return (str||'').replace(/[&<>"]/g, (c)=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' })[c]); }

function startEdit(index){
  const d=document.getElementById('steps');
  const item = d.querySelector(`[data-index="${index}"]`);
  if(!item) return;
  const descDiv = item.querySelector('.desc');
  const current = project.steps[index].description || '';
  descDiv.innerHTML = '';
  const ta = document.createElement('textarea');
  ta.id = `edit-${index}`;
  ta.value = current;
  descDiv.appendChild(ta);
  const actions = item.querySelector('.step-actions');
  actions.innerHTML = '';
  const saveBtn = document.createElement('button'); saveBtn.type='button'; saveBtn.textContent = 'Save'; saveBtn.onclick = ()=>saveEdit(index);
  const cancelBtn = document.createElement('button'); cancelBtn.type='button'; cancelBtn.textContent = 'Cancel'; cancelBtn.onclick = ()=>cancelEdit(index);
  actions.appendChild(saveBtn); actions.appendChild(cancelBtn);
  setTimeout(()=>{
    try{ ta.focus(); ta.select(); }catch(e){}
    ta.addEventListener('keydown', (ev)=>{
      if(ev.key === 'Escape') { ev.preventDefault(); cancelEdit(index); }
      if((ev.ctrlKey || ev.metaKey) && ev.key === 'Enter'){ ev.preventDefault(); saveEdit(index); }
    });
  }, 10);
}

function saveEdit(index){
  const ta = document.getElementById(`edit-${index}`);
  if(!ta) return;
  project.steps[index].description = ta.value.trim();
  render();
}

function cancelEdit(index){ render(); }

function deleteStep(index){
  if(index<0 || index>=project.steps.length) return;
  project.steps.splice(index,1);
  render();
}

async function annotateStep(index){
 const step = project.steps[index];
 if(!step || !step.image) return;
 await startAnnotateFlow(step.image, index, ()=>{
  render();
  setStatus("Annotations saved.");
 });
}

async function startAnnotateFlow(imagePath, stepIndex, onSaved){
 _inAnnotateFlow = true;
 try{
  await ipcRenderer.invoke("enter-annotate-mode");
  document.body.classList.add("annotating");
  const step = project.steps[stepIndex];
  openAnnotator(imagePath, stepIndex, ()=>{
   if(typeof onSaved === "function") onSaved();
  }, {
   stepNumber: stepIndex + 1,
   title: (step && step.description) || "",
  });
 }catch(err){
  _inAnnotateFlow = false;
  setStatus("Could not open annotator: "+err.message, "error");
 }
}

async function finishAnnotateFlow(){
 document.body.classList.remove("annotating");
 closeAnnotatorModal();
 _pendingAnnotateIndex = -1;
 _inAnnotateFlow = false;
 try{
  const result = await ipcRenderer.invoke("exit-annotate-mode");
  _compactMode = !!(result && result.compact);
  document.body.classList.toggle("compact", _compactMode);
  const btn = document.getElementById("compactBtn");
  if(btn) btn.innerHTML = _compactMode ? "⤴" : "🔳";
 }catch(err){
  setStatus("Could not restore window: "+err.message, "error");
 }
}

function onAnnotatorSaved(stepIndex){
 render();
 const step = project.steps[stepIndex];
 const source = step && step.sapContext ? step.sapContext.source : "unknown";
 const desc = step ? step.description : "";
 setStatus(`Step saved${desc ? ': "' + desc + '"' : ""}${source !== "unknown" ? " ("+source+")" : ""}.`);
}

async function captureStep(){
 const inputDesc=document.getElementById("stepDesc").value.trim();
 setStatus("Capturing SAPGUI screenshot...");
 try{
  const res = await ipcRenderer.invoke("capture-sapgui");
  applyCaptureResult(res, inputDesc, true);
 }catch(err){
  setStatus("Capture failed: "+err.message, "error");
 }
}

function applyCaptureResult(res, inputDesc="", openAnnotatorAfter=false){
 const file = (res && res.file) ? res.file : res;
 const desc = inputDesc || (res && res.title) || "";
 syncProjectFromUI();
 const step = {description:desc, image:file};
 if(res && res.context) step.sapContext = res.context;
 project.steps.push(step);
 document.getElementById("stepDesc").value="";
 const stepIndex = project.steps.length - 1;
 render();

 const source = (res && res.context && res.context.source) ? res.context.source : "unknown";
 const detail = desc ? `"${desc}"` : "no title detected";

 if(openAnnotatorAfter && file){
  _pendingAnnotateIndex = stepIndex;
  setStatus(`Step captured (${source}): ${detail}. Opening full-screen annotator...`);
  startAnnotateFlow(file, stepIndex, ()=> onAnnotatorSaved(stepIndex));
 } else {
  setStatus(`Step captured (${source}): ${detail}.`);
 }
}

function newProject(){
 if(project.steps.length > 0){
  showConfirm(
   "Start a new SOP? Save your current work first if you want to keep it.",
   ()=> applyNewProject(),
   "New SOP"
  );
  return;
 }
 applyNewProject();
}

function applyNewProject(){
 const keepAuthor = project.author || document.getElementById("author").value.trim();
 project = {
  title:"",
  author: keepAuthor,
  version:"1.0",
  reviewDate: todayDateIso(),
  steps:[]
 };
 _currentProjectPath = "";
 document.getElementById("title").value = "";
 document.getElementById("stepDesc").value = "";
 document.getElementById("version").value = "1.0";
 document.getElementById("reviewDate").value = todayDateIso();
 if(keepAuthor) document.getElementById("author").value = keepAuthor;
 setStatus("New SOP started.", "success");
 renderDeferred();
}

async function saveProject(){
 syncProjectFromUI();
 try{
  const result = await ipcRenderer.invoke("save-project-dialog", {
   project,
   filePath: _currentProjectPath || null,
  });
  if(result.canceled) return;
  if(!result.ok) throw new Error("Save failed");
  _currentProjectPath = result.filePath;
  setStatus(`Saved: ${result.fileName}`, "success");
 }catch(err){
  setStatus("Save failed: "+err.message, "error");
 }
}

async function loadProject(){
 try{
  const result = await ipcRenderer.invoke("load-project-dialog");
  if(result.canceled) return;
  if(!result.ok || !result.project) throw new Error("Load failed");
  project = result.project;
  if(!project.version) project.version = "1.0";
  if(!project.reviewDate) project.reviewDate = todayDateIso();
  if(!project.steps) project.steps = [];
  _currentProjectPath = result.filePath;
  syncUIToProject();
  renderDeferred();
  setStatus(`Loaded: ${result.fileName}`, "success");
 }catch(err){
  setStatus("Load failed: "+err.message, "error");
 }
}

async function loadLastProjectOnBoot(){
 setStatus("Loading project...", "loading");
 try{
  const result = await ipcRenderer.invoke("load-last-project");
  if(result.ok && result.project){
   project = result.project;
   if(!project.version) project.version = "1.0";
   if(!project.reviewDate) project.reviewDate = todayDateIso();
   if(!project.steps) project.steps = [];
   _currentProjectPath = result.filePath;
   syncUIToProject();
   renderDeferred();
   setStatus(`Resumed: ${result.fileName}`, "success");
   return;
  }
  const legacyPath = path.join(getDataDir(), "project.json");
  if(fs.existsSync(legacyPath)){
   project = JSON.parse(fs.readFileSync(legacyPath, "utf8"));
   if(!project.version) project.version = "1.0";
   if(!project.reviewDate) project.reviewDate = todayDateIso();
   if(!project.steps) project.steps = [];
   syncUIToProject();
   renderDeferred();
   setStatus("Loaded previous work. Save to keep a named copy.", "info");
   return;
  }
  renderDeferred();
  syncUIToProject();
  setStatus("Ready — start a new SOP or open an existing one.", "info");
 }catch(err){
  setStatus("Could not resume last project: "+err.message, "error");
 }
}

async function exportDoc(){
 syncProjectFromUI();
 setBusy(true);
 setStatus("Exporting Word document...", "loading");
 try{
  ensureDir(getExportsDir());
  const { exportWord } = require("./exporter");
  const outPath = await exportWord(project);
  setStatus(`DOCX exported: ${outPath}`, "success");
 }catch(err){
  setStatus("Export failed: "+err.message, "error");
 }finally{
  setBusy(false);
 }
}

async function toggleCompact(){
 _compactMode = !_compactMode;
 try{
  await ipcRenderer.invoke('set-compact-mode', _compactMode);
  document.body.classList.toggle('compact', _compactMode);
  ["compactBtn","compactBtnMain"].forEach((id)=>{
   const btn = document.getElementById(id);
   if(!btn) return;
   if(id === "compactBtn") btn.innerHTML = _compactMode ? '⤴' : '🔳';
   else btn.textContent = _compactMode ? '⤴ Restore' : '🔳 Compact';
  });
  setStatus(_compactMode ? 'Compact mode enabled.' : 'Restored window.', 'success');
 }catch(err){
  setStatus('Failed to toggle compact mode: '+err.message, 'error');
 }
}

function normalizeKeyPart(key){
 if(key === " ") return "Space";
 if(/^F\d{1,2}$/.test(key)) return key;
 if(key.length === 1) return key.toUpperCase();
 return key;
}

function setupHotkeyRecorder(){
 const input = document.getElementById("hotkeyInput");
 if(!input) return;
 input.addEventListener("keydown", (e)=>{
  e.preventDefault();
  e.stopPropagation();
  const parts = [];
  if(e.ctrlKey || e.metaKey) parts.push("CommandOrControl");
  if(e.altKey) parts.push("Alt");
  if(e.shiftKey) parts.push("Shift");
  const key = normalizeKeyPart(e.key);
  if(["Control","Alt","Shift","Meta","Command"].includes(key)) return;
  parts.push(key);
  input.value = parts.join("+");
 });
}

function applyHotkeyPreset(accelerator){
 const input = document.getElementById("hotkeyInput");
 if(input) input.value = accelerator;
}

async function openSettings(){
 try{
  const settings = await ipcRenderer.invoke("get-settings");
  _captureHotkey = settings.captureHotkey || "Alt+C";
  const input = document.getElementById("hotkeyInput");
  if(input) input.value = _captureHotkey;
  document.getElementById("hotkeyStatus").textContent = "Current: " + displayHotkey(_captureHotkey);
  document.getElementById("settingsModal").classList.add("open");
 }catch(err){
  setStatus("Could not load settings: "+err.message, "error");
 }
}

function closeSettings(){
 document.getElementById("settingsModal").classList.remove("open");
}

async function saveSettingsUi(){
 const input = document.getElementById("hotkeyInput");
 const hotkey = (input && input.value.trim()) || "Alt+C";
 try{
  const result = await ipcRenderer.invoke("save-settings", { captureHotkey: hotkey });
  if(result.hotkey && !result.hotkey.ok){
   document.getElementById("hotkeyStatus").textContent = result.hotkey.error;
   document.getElementById("hotkeyStatus").style.color = "#a00";
   return;
  }
  _captureHotkey = result.settings.captureHotkey;
  updateCaptureButtonTitle();
  closeSettings();
  setStatus("Settings saved. Capture hotkey: " + displayHotkey(_captureHotkey));
 }catch(err){
  setStatus("Failed to save settings: "+err.message, "error");
 }
}

async function loadInitialSettings(){
 try{
  const settings = await ipcRenderer.invoke("get-settings");
  _captureHotkey = settings.captureHotkey || "Alt+C";
  updateCaptureButtonTitle();
 }catch(e){}
}

async function closeAnnotateUi(){
 closeAnnotatorModal();
 if(_pendingAnnotateIndex >= 0){
  const step = project.steps[_pendingAnnotateIndex];
  const source = step && step.sapContext ? step.sapContext.source : "unknown";
  const desc = step ? step.description : "";
  setStatus(`Step captured (${source})${desc ? ': "' + desc + '"' : ""}. Annotation skipped.`);
 }
 await finishAnnotateFlow();
}

async function saveAnnotateUi(){
 saveAnnotation();
 await finishAnnotateFlow();
}

function closeAnnotator(){
 closeAnnotateUi();
}

function setupEventListeners(){
 document.querySelector(".compact-bar")?.addEventListener("click", handleToolbarClick);
 document.querySelector(".app-main .toolbar")?.addEventListener("click", handleToolbarClick);
}

function handleToolbarClick(e){
  if(document.body.classList.contains("ui-busy")) return;
  const btn = e.target.closest("[data-action]");
  if(!btn) return;
  e.preventDefault();
  const action = btn.dataset.action;
  if(action === "compact") toggleCompact();
  else if(action === "new") newProject();
  else if(action === "capture") captureStep();
  else if(action === "save") saveProject();
  else if(action === "load") loadProject();
  else if(action === "export") exportDoc();
  else if(action === "settings") openSettings();
}

function setupOtherEventListeners(){

 document.getElementById("steps")?.addEventListener("click", (e)=>{
  const el = e.target.closest("[data-step-action]");
  if(!el) return;
  const index = parseInt(el.dataset.index, 10);
  if(Number.isNaN(index)) return;
  const action = el.dataset.stepAction;
  if(action === "annotate") annotateStep(index);
  else if(action === "edit") startEdit(index);
  else if(action === "delete") deleteStep(index);
 });

 document.getElementById("settingsModal")?.addEventListener("click", (e)=>{
  if(e.target.id === "settingsModal") closeSettings();
  const btn = e.target.closest("[data-action]");
  if(!btn) return;
  if(btn.dataset.action === "settings-cancel") closeSettings();
  if(btn.dataset.action === "settings-save") saveSettingsUi();
 });

 document.getElementById("confirmModal")?.addEventListener("click", (e)=>{
  if(e.target.id === "confirmModal") closeConfirm();
  const btn = e.target.closest("[data-action]");
  if(!btn) return;
  if(btn.dataset.action === "confirm-cancel") closeConfirm();
  if(btn.dataset.action === "confirm-ok") runConfirmOk();
 });

 document.querySelector(".hotkey-presets")?.addEventListener("click", (e)=>{
  const btn = e.target.closest("[data-hotkey]");
  if(btn) applyHotkeyPreset(btn.dataset.hotkey);
 });

 document.getElementById("annotatorModal")?.addEventListener("click", (e)=>{
  if(e.target.id === "annotatorModal") return;
  const btn = e.target.closest("[data-action]");
  if(!btn) return;
  const action = btn.dataset.action;
  if(action === "annot-skip") closeAnnotateUi();
  else if(action === "annot-save") saveAnnotateUi();
  else if(action === "annot-undo") undoAnnotation();
  else if(action === "annot-clear") clearAnnotations();
 });

 if(ipcRenderer){
  ipcRenderer.on("capture-complete", (_event, res)=> applyCaptureResult(res, "", true));
  ipcRenderer.on("capture-error", (_event, message)=> setStatus("Capture failed: "+message, "error"));
 }
}

function boot(){
 try{
  setupEventListeners();
  setupOtherEventListeners();
  initAnnotator();
  setupHotkeyRecorder();
  syncUIToProject();
  renderDeferred();
  setStatus("Starting...", "loading");
  loadInitialSettings();
  setTimeout(()=>{ loadLastProjectOnBoot(); }, 0);
 }catch(err){
  console.error(err);
  setStatus("Startup error: "+err.message, "error");
 }
}

if(document.readyState === "loading"){
 document.addEventListener("DOMContentLoaded", boot);
} else {
 boot();
}
