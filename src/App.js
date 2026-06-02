const {ipcRenderer}=require("electron");
const fs=require("fs");
const path=require("path");
const { pathToFileURL } = require('url');
const { getDataDir, getExportsDir, ensureDir } = require("../paths");
const { formatStepDescription, describeCaptureContext, isSapSuccessStatusBar, isSapForegroundContext } = require("./sapgui-context");
const { formatChangePhrase, descriptionListsChanges } = require("./modules/recording/sessionRecorder");
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
	documentId:"",
	department:"",
	classification:"",
	steps:[]
};
let _brandingLogoPath = "";
let _compactMode = false;
let _sortable = null;
let _captureHotkey = "Alt+C";
let _sessionHotkey = "Ctrl+Shift+S";
let _captureTarget = "sap";
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

function valueOf(id){
 const el = document.getElementById(id);
 return el ? el.value : "";
}

function syncProjectFromUI(){
 project.title = document.getElementById("title").value.trim();
 project.author = document.getElementById("author").value.trim();
 project.version = document.getElementById("version").value.trim() || "1.0";
 project.reviewDate = document.getElementById("reviewDate").value || todayDateIso();
 project.documentId = valueOf("documentId").trim();
 project.department = valueOf("department").trim();
 project.classification = valueOf("classification");
}

function syncUIToProject(){
 document.getElementById("title").value = project.title || "";
 document.getElementById("author").value = project.author || "";
 document.getElementById("version").value = project.version || "1.0";
 if(!project.reviewDate) project.reviewDate = todayDateIso();
 document.getElementById("reviewDate").value = project.reviewDate;
 const docId = document.getElementById("documentId"); if(docId) docId.value = project.documentId || "";
 const dept = document.getElementById("department"); if(dept) dept.value = project.department || "";
 const cls = document.getElementById("classification"); if(cls) cls.value = project.classification || "";
}

function updateCaptureButtonTitle(){
 const modeLabel = _captureTarget === "non-sap" ? "Non-SAP" : "SAP GUI";
 const label = `Capture (${displayHotkey(_captureHotkey)}) · ${modeLabel} · Session (${displayHotkey(_sessionHotkey)})`;
 ["captureBtn","captureBtnMain"].forEach((id)=>{
  const btn = document.getElementById(id);
  if(btn) btn.title = label;
 });
}

function displayHotkey(accelerator){
 return (accelerator || "").replace(/CommandOrControl/g, "Ctrl");
}

function isStepStatusBarOnly(step){
 if(step && step.snapshotMeta && step.snapshotMeta.statusBarOnly) return true;
 return isSapSuccessStatusBar(step && step.sapContext ? step.sapContext.statusBar : "");
}

function formatStepChangesMeta(step){
 if(isStepStatusBarOnly(step)) return "";
 if(!step || !Array.isArray(step.changes) || !step.changes.length) return "";
 if(descriptionListsChanges(step.description, step.changes)) return "";
 return step.changes.map(formatChangePhrase).join(" · ");
}

function formatStepMetaLine(step){
 if(isStepStatusBarOnly(step)) return "";
 if(step && step.snapshotMeta && step.snapshotMeta.foregroundIsSap === false) return "";
 const parts = [];
 const changeMeta = formatStepChangesMeta(step);
 const controlCount = step && step.snapshotMeta ? step.snapshotMeta.controlCount || 0 : 0;
 if(changeMeta) parts.push(changeMeta);
 const ctxMeta = formatSapContextMeta(step.sapContext);
 if(ctxMeta) parts.push(ctxMeta);
 if(controlCount) parts.push(`${controlCount} controls`);
 if(controlCount && step.snapshotMeta && step.snapshotMeta.captureSource) parts.push(step.snapshotMeta.captureSource);
 if(controlCount && step.sessionRecord) parts.push("session");
 return parts.join(" | ");
}

function formatSapContextMeta(ctx){
 if(!ctx) return "";
 if(!isSapForegroundContext(ctx)) return "";
 const parts = [];
 if(ctx.transaction) parts.push(ctx.transaction);
 if(ctx.systemName) parts.push(ctx.systemName);
 if(ctx.statusBar) parts.push("status");
 if(ctx.source && String(ctx.source).includes("ocr")) parts.push("OCR");
 if(ctx.source && (ctx.source.includes("scripting") || ctx.source.includes("sap-scripting"))){
  parts.push("scripting");
 } else if(ctx.source && (ctx.source.includes("uia") || ctx.source.includes("child"))){
  parts.push("status detected");
 } else if(ctx.source === "win32" || !ctx.statusBar){
  parts.push("window only");
 }
 return parts.length ? parts.join(" · ") : "";
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
  empty.textContent = "No steps yet — use Capture (Alt+C) or Session (Ctrl+Shift+S) while working in SAP.";
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
    <div class="desc">${escapeHtml(s.description||"")}${formatStepMetaLine(s) ? `<div class="step-meta">${escapeHtml(formatStepMetaLine(s))}</div>` : ""}</div>
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

function formatStepSavedStatus(step){
 if(!step) return "Step saved.";
 if(isStepStatusBarOnly(step)){
  return step.description ? `Step saved: "${step.description}".` : "Step saved (status bar).";
 }
 const parts = [];
 if(step.description) parts.push(`"${step.description}"`);
 const changeCount = (step.changes && step.changes.length) || 0;
 if(changeCount){
  const fields = formatStepChangesMeta(step);
  parts.push(`${changeCount} field(s): ${fields}`);
 } else if(step.snapshotMeta && step.snapshotMeta.error){
  const err = String(step.snapshotMeta.error);
  if(err.includes("no fields found") || err.includes("No SAP fields detected")){
   parts.push("no fields detected on screen");
  } else if(err.includes("Scripting unavailable") || err.includes("GetScriptingEngine")){
   parts.push("no fields — enable SAP GUI Scripting (Options → Accessibility & Scripting)");
  } else {
   parts.push("no fields — " + err);
  }
 } else if(step.sessionRecord && step.snapshotMeta && !step.snapshotMeta.controlCount && isSapForegroundContext(step.sapContext)){
  parts.push("no fields detected on screen");
 }
 const source = step.sapContext ? step.sapContext.source : "";
 if(source && source !== "unknown") parts.push(source);
 return "Step saved" + (parts.length ? ": " + parts.join(" · ") : "") + ".";
}

function onAnnotatorSaved(stepIndex){
 render();
 const step = project.steps[stepIndex];
 setStatus(formatStepSavedStatus(step));
}

function formatCaptureTimings(timings){
 if(!timings || !timings.totalMs) return "";
 const parts = [`${timings.totalMs}ms total`];
 if(timings.snapshotMs != null) parts.push(`SAP ${timings.snapshotMs}ms`);
 if(timings.vbsDiscoverMs != null) parts.push(`VBS discover ${timings.vbsDiscoverMs}ms`);
 if(timings.contextMs != null) parts.push(`context ${timings.contextMs}ms`);
 if(timings.screenshotMs != null) parts.push(`screenshot ${timings.screenshotMs}ms`);
 if(timings.ocrMs != null && timings.ocrMs > 50) parts.push(`OCR ${timings.ocrMs}ms`);
 return parts.join(" · ");
}

async function captureStep(){
 const inputDesc=document.getElementById("stepDesc").value.trim();
 setBusy(true);
 setStatus("Capturing SAP GUI…");
 try{
  const res = await ipcRenderer.invoke("capture-sapgui");
  applyCaptureResult(res, inputDesc, true);
 }catch(err){
  setStatus("Capture failed: "+err.message, "error");
 }finally{
  setBusy(false);
 }
}

function applySessionStepResult(payload, openAnnotatorAfter=true, inputDesc=""){
 const step = payload && payload.step ? payload.step : null;
 if(!step) return;
 const autoDesc = step.description || "";
 const changeMeta = formatStepChangesMeta(step);
 if(inputDesc){
  step.description = changeMeta ? `${inputDesc} (${changeMeta})` : inputDesc;
 } else if(!autoDesc && changeMeta){
  step.description = changeMeta;
 }
 syncProjectFromUI();
 project.steps.push(step);
 document.getElementById("stepDesc").value="";
 const stepIndex = project.steps.length - 1;
 render();

 const changeCount = (step.changes && step.changes.length) || 0;
 const controlCount = payload.snapshotControlCount || 0;
 const detail = step.description ? `"${step.description}"` : "SAP step";
 let changeNote = "";
 if(isStepStatusBarOnly(step)){
  changeNote = "";
 } else if(changeCount){
  changeNote = ` — ${changeCount} field(s) recorded`;
 } else if(payload.snapshotError){
  changeNote = ` — ${payload.snapshotError}`;
 } else if(!controlCount){
  changeNote = ` — no SAP fields detected`;
 }
 if(payload.snapshotDebugPath) changeNote += ` · Diagnostics: ${payload.snapshotDebugPath}`;
 else if(payload.logPath) changeNote += ` · log: ${payload.logPath}`;

 if(openAnnotatorAfter && step.image){
  _pendingAnnotateIndex = stepIndex;
  const timingNote = formatCaptureTimings(payload && payload.timings);
  setStatus(`Step ${step.stepNumber || stepIndex + 1}: ${detail}${changeNote}${timingNote ? " · "+timingNote : ""}. Opening annotator...`);
  startAnnotateFlow(step.image, stepIndex, ()=> onAnnotatorSaved(stepIndex));
 } else {
  const timingNote = formatCaptureTimings(payload && payload.timings);
  setStatus(`Step ${step.stepNumber || stepIndex + 1}: ${detail}${changeNote}${timingNote ? " · "+timingNote : ""}.`);
 }
}

function applyCaptureResult(res, inputDesc="", openAnnotatorAfter=false){
 if(res && res.step){
  applySessionStepResult(res, openAnnotatorAfter, inputDesc);
  return;
 }
 const file = (res && res.file) ? res.file : res;
 const context = res && res.context ? res.context : null;
 const autoDesc = context ? formatStepDescription(context) : (res && res.title) || "";
 const desc = inputDesc || autoDesc || "";
 syncProjectFromUI();
 const step = {description:desc, image:file};
 if(context) step.sapContext = context;
 project.steps.push(step);
 document.getElementById("stepDesc").value="";
 const stepIndex = project.steps.length - 1;
 render();

 const source = (context && context.source) ? context.source : "unknown";
 const detail = desc ? `"${desc}"` : "no title detected";
 const contextDetail = (res && res.contextDetail) || describeCaptureContext(context);

 if(openAnnotatorAfter && file){
  _pendingAnnotateIndex = stepIndex;
  const timingNote = formatCaptureTimings(res && res.timings);
  setStatus(`Step captured (${source}): ${detail}${contextDetail}${timingNote ? " · "+timingNote : ""}. Opening full-screen annotator...`);
  startAnnotateFlow(file, stepIndex, ()=> onAnnotatorSaved(stepIndex));
 } else {
  const timingNote = formatCaptureTimings(res && res.timings);
  setStatus(`Step captured (${source}): ${detail}${contextDetail}${timingNote ? " · "+timingNote : ""}.`);
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
 const keepDept = project.department || valueOf("department").trim();
 project = {
  title:"",
  author: keepAuthor,
  version:"1.0",
  reviewDate: todayDateIso(),
  documentId:"",
  department: keepDept,
  classification:"",
  steps:[]
 };
 _currentProjectPath = "";
 document.getElementById("title").value = "";
 document.getElementById("stepDesc").value = "";
 document.getElementById("version").value = "1.0";
 document.getElementById("reviewDate").value = todayDateIso();
 if(keepAuthor) document.getElementById("author").value = keepAuthor;
 syncUIToProject();
 try{ ipcRenderer.invoke("reset-session-recorder"); }catch(e){}
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

async function resetSessionBaseline(){
 try{ await ipcRenderer.invoke("reset-session-recorder"); }catch(e){}
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
  await resetSessionBaseline();
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
  let branding = null;
  try{
   const settings = await ipcRenderer.invoke("get-settings");
   branding = settings && settings.branding ? settings.branding : null;
  }catch(e){ /* export with default branding */ }
  const outPath = await exportWord(project, branding);
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

function readCaptureTargetFromUI(){
 const el = document.querySelector('input[name="captureTarget"]:checked');
 return el && el.value === "non-sap" ? "non-sap" : "sap";
}

function applyCaptureTargetToUI(target){
 _captureTarget = target === "non-sap" ? "non-sap" : "sap";
 const sapRadio = document.getElementById("captureTargetSap");
 const nonSapRadio = document.getElementById("captureTargetNonSap");
 if(sapRadio) sapRadio.checked = _captureTarget === "sap";
 if(nonSapRadio) nonSapRadio.checked = _captureTarget === "non-sap";
 updateCaptureButtonTitle();
}

async function persistCaptureTarget(){
 const captureTarget = readCaptureTargetFromUI();
 if(captureTarget === _captureTarget) return;
 _captureTarget = captureTarget;
 try{
  await ipcRenderer.invoke("save-settings", { captureTarget });
  updateCaptureButtonTitle();
 }catch(err){
  setStatus("Could not save capture mode: "+err.message, "error");
 }
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
  const maxFieldsInput = document.getElementById("maxFieldsInput");
  if(maxFieldsInput){
   const maxFields = Number(settings.maxCaptureFields);
   maxFieldsInput.value = Number.isFinite(maxFields) && maxFields >= 1 ? maxFields : 10;
  }
  applyBrandingToUI(settings.branding || {});
  document.getElementById("hotkeyStatus").textContent = "Current: " + displayHotkey(_captureHotkey);
  const paths = await ipcRenderer.invoke("get-data-paths");
  const pathEl = document.getElementById("dataFolderPath");
  if(pathEl && paths){
   const snap = paths.snapshotDebugPath || path.join(paths.dataDir || "", "snapshot-last.json");
   pathEl.textContent = snap;
   pathEl.title = snap;
  }
  document.getElementById("settingsModal").classList.add("open");
 }catch(err){
  setStatus("Could not load settings: "+err.message, "error");
 }
}

function applyBrandingToUI(branding){
 const b = branding || {};
 _brandingLogoPath = b.logoPath || "";
 const set = (id, val)=>{ const el = document.getElementById(id); if(el) el.value = val; };
 set("brandCompany", b.companyName || "");
 set("brandAccent", b.accentColor || "#F97316");
 set("brandFooter", b.footerText || "");
 set("brandTemplate", b.template || "standard");
 renderLogoPreview();
}

function renderLogoPreview(){
 const img = document.getElementById("brandLogoPreview");
 const name = document.getElementById("brandLogoName");
 if(img){
  if(_brandingLogoPath){
   try{ img.src = pathToFileURL(path.resolve(_brandingLogoPath)).href; }catch(e){ img.src = _brandingLogoPath; }
   img.style.display = "inline-block";
  } else {
   img.removeAttribute("src");
   img.style.display = "none";
  }
 }
 if(name) name.textContent = _brandingLogoPath ? path.basename(_brandingLogoPath) : "No logo selected";
}

async function pickLogo(){
 try{
  const result = await ipcRenderer.invoke("pick-logo-file");
  if(result && result.canceled) return;
  if(result && result.ok && result.path){
   _brandingLogoPath = result.path;
   renderLogoPreview();
  }
 }catch(err){
  setStatus("Could not choose logo: "+err.message, "error");
 }
}

function removeLogo(){
 _brandingLogoPath = "";
 renderLogoPreview();
}

function readBrandingFromUI(){
 return {
  companyName: (valueOf("brandCompany") || "").trim(),
  logoPath: _brandingLogoPath || "",
  accentColor: valueOf("brandAccent") || "#F97316",
  footerText: (valueOf("brandFooter") || "").trim(),
  template: valueOf("brandTemplate") || "standard",
 };
}

async function openDataFolder(){
 try{
  const result = await ipcRenderer.invoke("open-data-folder");
  if(result && result.path) setStatus("Opened: "+result.path);
 }catch(err){
  setStatus("Could not open data folder: "+err.message, "error");
 }
}

function closeSettings(){
 document.getElementById("settingsModal").classList.remove("open");
}

async function saveSettingsUi(){
 const input = document.getElementById("hotkeyInput");
 const hotkey = (input && input.value.trim()) || "Alt+C";
 const maxFieldsInput = document.getElementById("maxFieldsInput");
 let maxCaptureFields = maxFieldsInput ? parseInt(maxFieldsInput.value, 10) : 10;
 if(!Number.isFinite(maxCaptureFields) || maxCaptureFields < 1) maxCaptureFields = 1;
 if(maxCaptureFields > 50) maxCaptureFields = 50;
 const branding = readBrandingFromUI();
 try{
  const result = await ipcRenderer.invoke("save-settings", { captureHotkey: hotkey, maxCaptureFields, branding });
  if(result.hotkey && !result.hotkey.ok){
   document.getElementById("hotkeyStatus").textContent = result.hotkey.error;
   document.getElementById("hotkeyStatus").style.color = "#a00";
   return;
  }
  _captureHotkey = result.settings.captureHotkey;
  updateCaptureButtonTitle();
  closeSettings();
  setStatus("Settings saved. Capture hotkey: " + displayHotkey(_captureHotkey) + ` · max fields: ${maxCaptureFields} · template: ${branding.template}`);
 }catch(err){
  setStatus("Failed to save settings: "+err.message, "error");
 }
}

async function loadInitialSettings(){
 try{
  const settings = await ipcRenderer.invoke("get-settings");
  _captureHotkey = settings.captureHotkey || "Alt+C";
  _sessionHotkey = settings.sessionHotkey || "Ctrl+Shift+S";
  applyCaptureTargetToUI(settings.captureTarget);
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
  if(btn.dataset.action === "open-data-folder") openDataFolder();
  if(btn.dataset.action === "brand-pick-logo") pickLogo();
  if(btn.dataset.action === "brand-remove-logo") removeLogo();
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

 document.querySelector(".capture-target-options")?.addEventListener("change", (e)=>{
  if(e.target && e.target.name === "captureTarget") persistCaptureTarget();
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
  ipcRenderer.on("session-step-complete", (_event, payload)=> applyCaptureResult(payload, "", true));
  ipcRenderer.on("session-step-error", (_event, message)=> setStatus("Capture failed: "+message, "error"));
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
