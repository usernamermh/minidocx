const editor = document.getElementById("editor");
const appShell = document.getElementById("appShell");
const leftSidebar = document.getElementById("leftSidebar");
const topToolbar = document.getElementById("topToolbar");
const pageStage = document.getElementById("pageStage");
const outline = document.getElementById("outline");
const statusText = document.getElementById("statusText");
const serverAddress = document.getElementById("serverAddress");
const openBtn = document.getElementById("openBtn");
const openDocxInput = document.getElementById("openDocxInput");
const saveBtn = document.getElementById("saveBtn");
const exportBtn = document.getElementById("exportBtn");
const imageInput = document.getElementById("imageInput");
const fontFamily = document.getElementById("fontFamily");
const fontSize = document.getElementById("fontSize");
const highlightColor = document.getElementById("highlightColor");
const applyHighlightBtn = document.getElementById("applyHighlightBtn");
const pageSizeSelect = document.getElementById("pageSizeSelect");
const pageWidthInput = document.getElementById("pageWidthInput");
const pageHeightInput = document.getElementById("pageHeightInput");
const applyPageSizeBtn = document.getElementById("applyPageSizeBtn");
const outlineLevel1 = document.getElementById("outlineLevel1");
const outlineLevel2 = document.getElementById("outlineLevel2");
const outlineLevel3 = document.getElementById("outlineLevel3");
const paragraphStyleSelect = document.getElementById("paragraphStyleSelect");
const lineSpacingSelect = document.getElementById("lineSpacingSelect");
const spaceBeforeInput = document.getElementById("spaceBeforeInput");
const spaceAfterInput = document.getElementById("spaceAfterInput");
const saveStyleBtn = document.getElementById("saveStyleBtn");
const updateStyleBtn = document.getElementById("updateStyleBtn");
const formatPainterBtn = document.getElementById("formatPainterBtn");
const clearFormatBtn = document.getElementById("clearFormatBtn");
const resetParagraphSpacingBtn = document.getElementById("resetParagraphSpacingBtn");
const toggleNumberingBtn = document.getElementById("toggleNumberingBtn");
const numberLevelUpBtn = document.getElementById("numberLevelUpBtn");
const numberLevelDownBtn = document.getElementById("numberLevelDownBtn");
const numberFormatSelect = document.getElementById("numberFormatSelect");
const shortcutSettingsBtn = document.getElementById("shortcutSettingsBtn");
const shortcutList = document.getElementById("shortcutList");
const resetShortcutsBtn = document.getElementById("resetShortcutsBtn");
const shortcutModal = document.getElementById("shortcutModal");
const closeShortcutModalBtn = document.getElementById("closeShortcutModalBtn");
const historyPanelBtn = document.getElementById("historyPanelBtn");
const historyModal = document.getElementById("historyModal");
const closeHistoryModalBtn = document.getElementById("closeHistoryModalBtn");
const browseDirBtn = document.getElementById("browseDirBtn");
const addFavoriteDirBtn = document.getElementById("addFavoriteDirBtn");
const recentFileList = document.getElementById("recentFileList");
const favoriteDirList = document.getElementById("favoriteDirList");
const dirFileList = document.getElementById("dirFileList");
const fallbackFileBtn = document.querySelector(".fallback-file-btn");
const deleteTableBtn = document.getElementById("deleteTableBtn");
const sidebarToggleBtn = document.getElementById("sidebarToggleBtn");
const toolbarToggleBtn = document.getElementById("toolbarToggleBtn");

let currentStyles = { paragraph: [] };
let customShortcuts = {};
let currentFileHandle = null;
let currentFileName = "mini-docx.docx";
let formatPainterPayload = null;
let isDirty = false;
let isLoadingDocument = false;
let editorZoom = 1;
let savedSelectionRange = null;
let outlineFilter = { 1: true, 2: true, 3: false };
let pageSize = { widthTwips: 11906, heightTwips: 16838 };
let docxMeta = null;
let stylesDirty = false;
let numberingDirty = false;
let cleanEditorHtml = "";
let editorUndoStack = [];
let editorRedoStack = [];
let suppressEditorHistory = false;
let debugLogSequence = 0;

const SHORTCUT_STORAGE_KEY = "mini_docx_shortcuts";
const OUTLINE_FILTER_KEY = "mini_docx_outline_filter";
const LAYOUT_STORAGE_KEY = "mini_docx_layout";
const HANDLE_DB_NAME = "mini_docx_handles";
const HANDLE_DB_VERSION = 1;
const STORE_RECENT = "recent_files";
const STORE_FAVORITE = "favorite_dirs";
const MAX_RECENT_FILES = 12;
const NUMBER_FORMATS = new Set(["decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman"]);
const DEFAULT_NUMBER_FORMAT = "decimal";
const MAX_NUMBERING_LEVEL = 8;
const MAX_EDITOR_HISTORY = 200;
const NUMBERING_LEVEL_INDENT_PX = 24;
const MIN_NUMBERING_PREFIX_PX = 18;
const DEFAULT_FONT_FAMILY = '"Times New Roman", SimSun';
const DEFAULT_LINE_SPACING = 1.5;
const OPENED_DOCUMENT_WIDTH_MM = 350;
const DEFAULT_DOCUMENT_HEIGHT_MM = 297;
const DEFAULT_PARAGRAPH_SPACING = 1;
const ALLOWED_STYLE_ORDER = ["Normal", "Heading1", "Heading2", "Heading3"];
const DEFAULT_SHORTCUTS = {
  bold: "Ctrl+B",
  italic: "Ctrl+I",
  underline: "Ctrl+U",
  undo: "Ctrl+Z",
  redo: "Ctrl+Y",
  save: "Ctrl+S",
  heading1: "Ctrl+Alt+1",
  heading2: "Ctrl+Alt+2",
  heading3: "Ctrl+Alt+3",
  paragraph: "Ctrl+Alt+0",
};

function debugLog(event, data = {}) {
  debugLogSequence += 1;
  const payload = {
    event,
    data: {
      seq: debugLogSequence,
      ...data,
    },
  };
  try {
    fetch("/api/debug-log", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      keepalive: true,
    }).catch(() => {});
  } catch {
    // Best-effort logging only.
  }
}

function loadLayoutState() {
  try {
    const raw = window.localStorage.getItem(LAYOUT_STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveLayoutState(state) {
  try {
    window.localStorage.setItem(LAYOUT_STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Layout memory is nice to have, not required.
  }
}

function setSidebarCollapsed(collapsed) {
  appShell?.classList.toggle("is-sidebar-collapsed", collapsed);
  if (sidebarToggleBtn) {
    sidebarToggleBtn.textContent = collapsed ? "显示左侧" : "隐藏左侧";
    sidebarToggleBtn.setAttribute("aria-expanded", String(!collapsed));
  }
}

function setToolbarCollapsed(collapsed) {
  topToolbar?.classList.toggle("is-collapsed", collapsed);
  if (toolbarToggleBtn) {
    toolbarToggleBtn.textContent = collapsed ? "显示上方" : "隐藏上方";
    toolbarToggleBtn.setAttribute("aria-expanded", String(!collapsed));
  }
}

function initLayoutToggles() {
  const state = loadLayoutState();
  setSidebarCollapsed(Boolean(state.sidebarCollapsed));
  setToolbarCollapsed(Boolean(state.toolbarCollapsed));

  sidebarToggleBtn?.addEventListener("click", () => {
    const nextState = {
      ...loadLayoutState(),
      sidebarCollapsed: !appShell?.classList.contains("is-sidebar-collapsed"),
    };
    setSidebarCollapsed(nextState.sidebarCollapsed);
    saveLayoutState(nextState);
  });

  toolbarToggleBtn?.addEventListener("click", () => {
    const nextState = {
      ...loadLayoutState(),
      toolbarCollapsed: !topToolbar?.classList.contains("is-collapsed"),
    };
    setToolbarCollapsed(nextState.toolbarCollapsed);
    saveLayoutState(nextState);
  });
}

let numberingListSeed = Date.now();
const numberingMeasureCanvas = document.createElement("canvas");
const numberingMeasureContext = numberingMeasureCanvas.getContext("2d");

const SHORTCUT_ACTIONS = {
  bold: { label: "加粗", run: () => exec("bold") },
  italic: { label: "斜体", run: () => exec("italic") },
  underline: { label: "下划线", run: () => exec("underline") },
  undo: { label: "撤销", run: () => exec("undo") },
  redo: { label: "重做", run: () => exec("redo") },
  save: { label: "保存文件", run: () => saveDocx({ interactive: false, allowPicker: true }) },
  heading1: { label: "设为 H1", run: () => applyBlockPreset("H1") },
  heading2: { label: "设为 H2", run: () => applyBlockPreset("H2") },
  heading3: { label: "设为 H3", run: () => applyBlockPreset("H3") },
  paragraph: { label: "设为正文", run: () => applyBlockPreset("P") },
};

function setStatus(text) {
  statusText.textContent = text;
}

function cloneDocxMeta(meta) {
  if (!meta || typeof meta !== "object") return null;
  const next = {};
  if (typeof meta.source_docx_b64 === "string" && meta.source_docx_b64) {
    next.source_docx_b64 = meta.source_docx_b64;
  }
  if (typeof meta.styles_xml_b64 === "string" && meta.styles_xml_b64) {
    next.styles_xml_b64 = meta.styles_xml_b64;
  }
  if (typeof meta.numbering_xml_b64 === "string" && meta.numbering_xml_b64) {
    next.numbering_xml_b64 = meta.numbering_xml_b64;
  }
  return Object.keys(next).length ? next : null;
}

function nodeInEditor(node) {
  if (!node) return false;
  return node === editor || editor.contains(node);
}

function selectionInsideEditor(selection = window.getSelection()) {
  if (!selection || !selection.rangeCount) return false;
  const range = selection.getRangeAt(0);
  return nodeInEditor(range.commonAncestorContainer);
}

function captureEditorSelection() {
  const selection = window.getSelection();
  if (!selectionInsideEditor(selection)) return;
  const nextRange = selection.getRangeAt(0).cloneRange();
  if (
    savedSelectionRange
    && !savedSelectionRange.collapsed
    && nextRange.collapsed
    && document.activeElement
    && document.activeElement !== editor
    && !editor.contains(document.activeElement)
  ) {
    return;
  }
  savedSelectionRange = nextRange;
}

function restoreEditorSelection() {
  if (!savedSelectionRange || !nodeInEditor(savedSelectionRange.commonAncestorContainer)) {
    return false;
  }
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(savedSelectionRange.cloneRange());
  return true;
}

function activeEditorRange(selection = window.getSelection()) {
  if (selection && selectionInsideEditor(selection)) {
    return selection.getRangeAt(0);
  }
  if (savedSelectionRange && nodeInEditor(savedSelectionRange.commonAncestorContainer)) {
    return savedSelectionRange;
  }
  return null;
}

function nodePathFromEditor(node) {
  if (!node) return null;
  const path = [];
  let current = node;
  while (current && current !== editor) {
    const parent = current.parentNode;
    if (!parent) return null;
    const index = Array.prototype.indexOf.call(parent.childNodes, current);
    if (index < 0) return null;
    path.unshift(index);
    current = parent;
  }
  return current === editor ? path : null;
}

function nodeFromEditorPath(path) {
  if (!Array.isArray(path)) return null;
  let current = editor;
  for (const rawIndex of path) {
    const index = Number.parseInt(rawIndex, 10);
    if (!Number.isFinite(index) || index < 0 || index >= current.childNodes.length) {
      return null;
    }
    current = current.childNodes[index];
  }
  return current;
}

function clampRangeOffset(node, offset) {
  const raw = Number.parseInt(offset, 10);
  const normalized = Number.isFinite(raw) ? raw : 0;
  if (!node) return 0;
  if (node.nodeType === Node.TEXT_NODE) {
    const length = (node.textContent || "").length;
    return Math.max(Math.min(normalized, length), 0);
  }
  return Math.max(Math.min(normalized, node.childNodes.length), 0);
}

function serializeEditorSelection() {
  const selection = window.getSelection();
  if (!selectionInsideEditor(selection)) return null;
  const range = selection.getRangeAt(0);
  const startPath = nodePathFromEditor(range.startContainer);
  const endPath = nodePathFromEditor(range.endContainer);
  if (!startPath || !endPath) return null;
  return {
    startPath,
    startOffset: range.startOffset,
    endPath,
    endOffset: range.endOffset,
  };
}

function restoreSerializedEditorSelection(serialized) {
  if (!serialized) return false;
  const startNode = nodeFromEditorPath(serialized.startPath);
  const endNode = nodeFromEditorPath(serialized.endPath);
  if (!startNode || !endNode) return false;
  const range = document.createRange();
  range.setStart(startNode, clampRangeOffset(startNode, serialized.startOffset));
  range.setEnd(endNode, clampRangeOffset(endNode, serialized.endOffset));
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  savedSelectionRange = range.cloneRange();
  return true;
}

function placeCaretAtEditorEnd() {
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(editor);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
  savedSelectionRange = range.cloneRange();
}

function createEditorSnapshot() {
  return {
    html: editor.innerHTML,
    selection: serializeEditorSelection(),
  };
}

function trimEditorHistoryStacks() {
  if (editorUndoStack.length > MAX_EDITOR_HISTORY) {
    editorUndoStack = editorUndoStack.slice(editorUndoStack.length - MAX_EDITOR_HISTORY);
  }
  if (editorRedoStack.length > MAX_EDITOR_HISTORY) {
    editorRedoStack = editorRedoStack.slice(editorRedoStack.length - MAX_EDITOR_HISTORY);
  }
}

function recordUndoSnapshot() {
  if (isLoadingDocument || suppressEditorHistory) return;
  const snapshot = createEditorSnapshot();
  const last = editorUndoStack[editorUndoStack.length - 1];
  if (last && last.html === snapshot.html) return;
  editorUndoStack.push(snapshot);
  editorRedoStack = [];
  trimEditorHistoryStacks();
}

function resetEditorHistory() {
  editorUndoStack = [];
  editorRedoStack = [];
  recordUndoSnapshot();
}

function syncDirtyStateFromEditor() {
  if (isLoadingDocument) return;
  if (editor.innerHTML === cleanEditorHtml) {
    markClean();
    return;
  }
  markDirty();
}

function applyEditorSnapshot(snapshot) {
  if (!snapshot || typeof snapshot.html !== "string") return false;
  suppressEditorHistory = true;
  const previousLoading = isLoadingDocument;
  try {
    isLoadingDocument = true;
    editor.innerHTML = snapshot.html;
    refreshOutline();
  } finally {
    isLoadingDocument = previousLoading;
    suppressEditorHistory = false;
  }
  if (!restoreSerializedEditorSelection(snapshot.selection)) {
    placeCaretAtEditorEnd();
  }
  syncParagraphStyleSelect();
  syncDirtyStateFromEditor();
  return true;
}

function performEditorUndo() {
  if (!editorUndoStack.length) return false;
  const current = createEditorSnapshot();
  let previousIndex = -1;
  for (let index = editorUndoStack.length - 1; index >= 0; index -= 1) {
    if (editorUndoStack[index].html !== current.html) {
      previousIndex = index;
      break;
    }
  }
  if (previousIndex < 0) return false;
  const previous = editorUndoStack[previousIndex];
  editorUndoStack = editorUndoStack.slice(0, previousIndex);
  if (!previous) return false;
  editorRedoStack.push(current);
  trimEditorHistoryStacks();
  return applyEditorSnapshot(previous);
}

function performEditorRedo() {
  if (!editorRedoStack.length) return false;
  const current = createEditorSnapshot();
  const next = editorRedoStack.pop();
  const lastUndo = editorUndoStack[editorUndoStack.length - 1];
  if (!lastUndo || lastUndo.html !== current.html) {
    editorUndoStack.push(current);
  }
  trimEditorHistoryStacks();
  return applyEditorSnapshot(next);
}

function supportsFileSystemAccess() {
  return typeof window.showOpenFilePicker === "function" && typeof window.showSaveFilePicker === "function";
}

function updateFileUiState() {
  fallbackFileBtn.classList.toggle("is-visible", !supportsFileSystemAccess());
  const dirtyPrefix = isDirty ? "*" : "";
  saveBtn.textContent = currentFileHandle ? `${dirtyPrefix}保存文件 (${currentFileName})` : `${dirtyPrefix}保存文件`;
  document.title = `${isDirty ? "*" : ""}Mini DOCX Web Editor`;
}

function applyEditorZoom() {
  editor.style.zoom = String(editorZoom);
}

function setEditorZoom(nextZoom) {
  editorZoom = Math.min(Math.max(nextZoom, 0.7), 2);
  applyEditorZoom();
  setStatus(`正文缩放 ${Math.round(editorZoom * 100)}%`);
}

function adjustEditorZoom(delta) {
  setEditorZoom(Math.round((editorZoom + delta) * 10) / 10);
}

function resetEditorZoom() {
  setEditorZoom(1);
}

async function ensureHandlePermission(fileHandle, writable = false, interactive = true) {
  if (!fileHandle || typeof fileHandle.queryPermission !== "function") return true;
  const options = writable ? { mode: "readwrite" } : {};
  if (await fileHandle.queryPermission(options) === "granted") return true;
  if (!interactive) return false;
  return (await fileHandle.requestPermission(options)) === "granted";
}

function markDirty() {
  if (isLoadingDocument) return;
  isDirty = true;
  updateFileUiState();
}

function markClean() {
  isDirty = false;
  cleanEditorHtml = editor.innerHTML;
  updateFileUiState();
}

function confirmDiscardChanges() {
  if (!isDirty) return true;
  return window.confirm("当前有未保存内容，是否继续并放弃这些修改？");
}

function isUserCancel(error) {
  if (!error) return false;
  const message = String(error.message || error || "").toLowerCase();
  return error.name === "AbortError" || message.includes("aborted") || message.includes("cancel") || message.includes("denied");
}

function isStaleHandleError(error) {
  if (!error) return false;
  const message = String(error.message || error || "").toLowerCase();
  return message.includes("state cached") || message.includes("invalid") || message.includes("not found");
}

function handleAsyncError(error) {
  if (isUserCancel(error)) {
    setStatus("操作已取消");
    return;
  }
  setStatus(error.message || String(error));
  window.alert(error.message || String(error));
}

function openHandleDb() {
  return new Promise((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("当前浏览器不支持 IndexedDB。"));
      return;
    }
    const request = indexedDB.open(HANDLE_DB_NAME, HANDLE_DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_RECENT)) {
        db.createObjectStore(STORE_RECENT, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(STORE_FAVORITE)) {
        db.createObjectStore(STORE_FAVORITE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("无法打开本地存储。"));
  });
}

async function withHandleStore(storeName, mode, fn) {
  const db = await openHandleDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    const result = fn(store);
    tx.oncomplete = () => resolve(result);
    tx.onerror = () => reject(tx.error || new Error("存储操作失败。"));
  });
}

async function getStoreRecords(storeName) {
  return withHandleStore(storeName, "readonly", (store) => new Promise((resolve, reject) => {
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error || new Error("读取失败。"));
  }));
}

async function putStoreRecord(storeName, record) {
  return withHandleStore(storeName, "readwrite", (store) => new Promise((resolve, reject) => {
    const request = store.put(record);
    request.onsuccess = () => resolve(true);
    request.onerror = () => reject(request.error || new Error("保存失败。"));
  }));
}

async function deleteStoreRecord(storeName, id) {
  return withHandleStore(storeName, "readwrite", (store) => new Promise((resolve, reject) => {
    const request = store.delete(id);
    request.onsuccess = () => resolve(true);
    request.onerror = () => reject(request.error || new Error("删除失败。"));
  }));
}

async function isSameHandle(handleA, handleB) {
  if (!handleA || !handleB) return false;
  if (typeof handleA.isSameEntry !== "function") return false;
  try {
    return await handleA.isSameEntry(handleB);
  } catch {
    return false;
  }
}

async function recordRecentFile(handle) {
  if (!handle) return;
  const records = await getStoreRecords(STORE_RECENT);
  let existing = null;
  for (const item of records) {
    if (item.handle && (await isSameHandle(handle, item.handle))) {
      existing = item;
      break;
    }
  }
  const now = Date.now();
  const record = {
    id: existing?.id || (crypto.randomUUID ? crypto.randomUUID() : String(now)),
    name: handle.name || existing?.name || "unnamed.docx",
    handle,
    lastOpened: now,
  };
  await putStoreRecord(STORE_RECENT, record);
  const sorted = [...records.filter((item) => item.id !== record.id), record]
    .sort((a, b) => b.lastOpened - a.lastOpened);
  const extra = sorted.slice(MAX_RECENT_FILES);
  for (const item of extra) {
    await deleteStoreRecord(STORE_RECENT, item.id);
  }
}

async function addFavoriteDir(handle) {
  if (!handle) return;
  const records = await getStoreRecords(STORE_FAVORITE);
  for (const item of records) {
    if (item.handle && (await isSameHandle(handle, item.handle))) {
      return;
    }
  }
  const now = Date.now();
  const record = {
    id: crypto.randomUUID ? crypto.randomUUID() : String(now),
    name: handle.name || "目录",
    handle,
    createdAt: now,
  };
  await putStoreRecord(STORE_FAVORITE, record);
}

function openShortcutModal() {
  shortcutModal.classList.remove("hidden");
}

function closeShortcutModal() {
  shortcutModal.classList.add("hidden");
}

function defaultStyles() {
  return {
    paragraph: [
      { id: "Normal", name: "Normal", descriptor: [DEFAULT_FONT_FAMILY, 12, false, false, false], alignment: "left", outline_level: null, is_default: true, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "Heading1", name: "Heading 1", descriptor: [DEFAULT_FONT_FAMILY, 20, true, false, false], alignment: "left", outline_level: 0, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "Heading2", name: "Heading 2", descriptor: [DEFAULT_FONT_FAMILY, 16, true, false, false], alignment: "left", outline_level: 1, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "Heading3", name: "Heading 3", descriptor: [DEFAULT_FONT_FAMILY, 14, true, false, false], alignment: "left", outline_level: 2, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
    ],
  };
}

function cloneDescriptor(descriptor) {
  const fallback = [DEFAULT_FONT_FAMILY, 12, false, false, false, ""];
  const values = Array.isArray(descriptor) ? descriptor.slice(0, 6) : fallback.slice();
  while (values.length < 6) values.push(fallback[values.length]);
  values[0] = String(values[0] || DEFAULT_FONT_FAMILY);
  values[1] = Math.max(Number(values[1]) || 12, 1);
  values[2] = Boolean(values[2]);
  values[3] = Boolean(values[3]);
  values[4] = Boolean(values[4]);
  values[5] = String(values[5] || "");
  return values;
}

function fontFamilyForControl(family) {
  const raw = String(family || DEFAULT_FONT_FAMILY);
  if (raw.includes("Times New Roman") && raw.includes("SimSun")) {
    return "Times New Roman";
  }
  return raw.split(",")[0].replace(/["']/g, "").trim() || "Times New Roman";
}

function fontFamilyForDocument(family) {
  return fontFamilyForControl(family) === "Times New Roman" ? DEFAULT_FONT_FAMILY : family;
}

function normalizeStyleNumbering(numbering) {
  if (!numbering || typeof numbering !== "object") return null;
  const numId = String(numbering.num_id || numbering.numId || "").trim();
  if (!numId) return null;
  return {
    num_id: numId,
    ilvl: Math.max(Math.min(Number.parseInt(numbering.ilvl, 10) || 0, MAX_NUMBERING_LEVEL), 0),
  };
}

function normalizeStyles(styles) {
  const incoming = new Map();
  if (styles && Array.isArray(styles.paragraph)) {
    styles.paragraph.forEach((style) => {
      const id = String(style?.id || "").trim();
      if (ALLOWED_STYLE_ORDER.includes(id)) incoming.set(id, style);
    });
  }
  const merged = defaultStyles().paragraph.map((baseStyle) => {
    const style = incoming.get(baseStyle.id) || {};
    return {
      ...baseStyle,
      descriptor: cloneDescriptor(style.descriptor || baseStyle.descriptor),
      alignment: ["left", "center", "right", "justify"].includes(style.alignment) ? style.alignment : baseStyle.alignment,
      line_spacing: Math.max(Number(style.line_spacing ?? baseStyle.line_spacing) || DEFAULT_LINE_SPACING, 1),
      space_before: Math.max(Number(style.space_before ?? baseStyle.space_before) || DEFAULT_PARAGRAPH_SPACING, 0),
      space_after: Math.max(Number(style.space_after ?? baseStyle.space_after) || DEFAULT_PARAGRAPH_SPACING, 0),
      numbering: normalizeStyleNumbering(style.numbering),
    };
  });
  return { paragraph: merged };
}

function styleMap() {
  return new Map(currentStyles.paragraph.map((style) => [style.id, style]));
}

function getStyleById(id) {
  return styleMap().get(id) || styleMap().get("Normal") || null;
}

function styleIdFromTag(tagName) {
  if (tagName === "H1") return "Heading1";
  if (tagName === "H2") return "Heading2";
  if (tagName === "H3") return "Heading3";
  return "Normal";
}

function styleIdFromBlockStyleKey(key) {
  if (key === "heading1") return "Heading1";
  if (key === "heading2") return "Heading2";
  if (key === "heading3") return "Heading3";
  return "Normal";
}

function tagNameFromStyle(style) {
  if (!style) return "P";
  if (style.outline_level === 0) return "H1";
  if (style.outline_level === 1) return "H2";
  if (style.outline_level === 2) return "H3";
  return "P";
}

function blockStyleKey(style) {
  if (!style) return "normal";
  if (style.outline_level === 0) return "heading1";
  if (style.outline_level === 1) return "heading2";
  if (style.outline_level === 2) return "heading3";
  return "normal";
}

function styleLabel(style) {
  if (style.outline_level === 0) return `${style.name} (H1)`;
  if (style.outline_level === 1) return `${style.name} (H2)`;
  if (style.outline_level === 2) return `${style.name} (H3)`;
  return style.name;
}

function populateStyleSelect(selectedId = paragraphStyleSelect.value || "Normal") {
  paragraphStyleSelect.innerHTML = "";
  currentStyles.paragraph.forEach((style) => {
    const option = document.createElement("option");
    option.value = style.id;
    option.textContent = styleLabel(style);
    paragraphStyleSelect.appendChild(option);
  });
  paragraphStyleSelect.value = getStyleById(selectedId)?.id || "Normal";
}

function setServerAddress() {
  serverAddress.textContent = window.location.origin;
}

function ensureStarterContent() {
  if (editor.innerHTML.trim()) {
    return;
  }
  editor.innerHTML = '<h1 data-style-id="Heading1">Mini DOCX Web Editor</h1><p data-style-id="Normal">在这里开始编辑。</p>';
  refreshOutline();
}

function exec(command, value = null) {
  if (!["undo", "redo"].includes(command)) {
    recordUndoSnapshot();
  }
  restoreEditorSelection();
  const block = currentBlockElement();
  debugLog("exec", {
    command,
    value,
    blockTag: block?.tagName || null,
    blockStyleId: block?.dataset?.styleId || null,
    blockText: (block?.textContent || "").slice(0, 80),
    fontFamilyControl: fontFamily?.value || null,
    fontSizeControl: fontSize?.value || null,
  });
  editor.focus();
  document.execCommand(command, false, value);
  captureEditorSelection();
  refreshOutline();
  if (!["undo", "redo"].includes(command)) {
    markDirty();
  }
}

function applyTextBackground(color) {
  if (!color) return;
  recordUndoSnapshot();
  restoreEditorSelection();
  editor.focus();
  document.execCommand("styleWithCSS", false, true);
  const ok = document.execCommand("hiliteColor", false, color);
  if (!ok) {
    document.execCommand("backColor", false, color);
  }
  captureEditorSelection();
  refreshOutline();
  markDirty();
}

function applyBlockPreset(tagName) {
  exec("formatBlock", tagName === "P" ? "<p>" : `<${tagName.toLowerCase()}>`);
  const block = currentBlockElement();
  if (!block) return;
  const styleId = styleIdFromTag(block.tagName);
  applyStyleVisuals(block, getStyleById(styleId));
  paragraphStyleSelect.value = styleId;
  refreshOutline();
}

function captureFormatPainter() {
  restoreEditorSelection();
  const block = currentBlockElement();
  if (!block) {
    setStatus("请先把光标放在段落中。");
    return;
  }
  formatPainterPayload = {
    styleId: block.dataset.styleId || styleIdFromTag(block.tagName),
    descriptor: descriptorFromStyle(window.getComputedStyle(block)),
    alignment: block.style.textAlign || window.getComputedStyle(block).textAlign || "left",
    metrics: paragraphMetricsFromElement(block),
  };
  formatPainterBtn.classList.add("is-active");
  setStatus("格式刷已开启，请点击目标段落。");
}

function clearFormatPainter() {
  formatPainterPayload = null;
  formatPainterBtn.classList.remove("is-active");
}

function applyFormatPainterToBlock(block) {
  if (!formatPainterPayload || !block) return false;
  recordUndoSnapshot();
  const style = getStyleById(formatPainterPayload.styleId) || getStyleById("Normal");
  const target = applyStyleVisuals(block, style);
  target.style.fontFamily = formatPainterPayload.descriptor[0];
  target.style.fontSize = `${formatPainterPayload.descriptor[1] / 0.75}px`;
  target.style.fontWeight = formatPainterPayload.descriptor[2] ? "700" : "400";
  target.style.fontStyle = formatPainterPayload.descriptor[3] ? "italic" : "normal";
  target.style.textDecoration = formatPainterPayload.descriptor[4] ? "underline" : "none";
  target.style.textAlign = formatPainterPayload.alignment;
  applyParagraphMetrics(target, formatPainterPayload.metrics);
  refreshOutline();
  markDirty();
  setStatus("已应用格式刷。");
  return true;
}

function paragraphMetricsFromElement(element) {
  const style = window.getComputedStyle(element);
  return {
    lineSpacing: Math.max(parseFloat(style.lineHeight) / parseFloat(style.fontSize || "16"), 1) || DEFAULT_LINE_SPACING,
    spaceBefore: Math.max(parseFloat(element.dataset.spaceBefore || "0"), 0),
    spaceAfter: Math.max(parseFloat(element.dataset.spaceAfter || String(parseFloat(style.marginBottom || "0"))), 0),
  };
}

function ensureLineSpacingOption(value) {
  const normalized = String(Math.round(Math.max(Number(value) || 1, 1) * 10) / 10);
  const exists = Array.from(lineSpacingSelect.options).some((option) => option.value === normalized);
  if (!exists) {
    const option = document.createElement("option");
    option.value = normalized;
    option.textContent = `${normalized} 倍行距`;
    lineSpacingSelect.appendChild(option);
  }
  return normalized;
}

function syncParagraphMetricsControls() {
  const block = currentBlockElement();
  if (!block) return;
  const metrics = paragraphMetricsFromElement(block);
  lineSpacingSelect.value = ensureLineSpacingOption(metrics.lineSpacing);
  spaceBeforeInput.value = String(Math.round(metrics.spaceBefore));
  spaceAfterInput.value = String(Math.round(metrics.spaceAfter));
}

function applyParagraphMetrics(element, metrics) {
  if (!element) return;
  const lineSpacing = Math.max(Number(metrics.lineSpacing) || 1, 1);
  const spaceBefore = Math.max(Number(metrics.spaceBefore) || 0, 0);
  const spaceAfter = Math.max(Number(metrics.spaceAfter) || 0, 0);
  element.style.lineHeight = String(lineSpacing);
  element.style.marginTop = `${spaceBefore}px`;
  element.style.marginBottom = `${spaceAfter}px`;
  element.dataset.spaceBefore = String(spaceBefore);
  element.dataset.spaceAfter = String(spaceAfter);
}

function applyCurrentParagraphMetrics() {
  restoreEditorSelection();
  const blocks = selectedBlockElements();
  if (!blocks.length) {
    setStatus("请先选中段落。");
    return;
  }
  recordUndoSnapshot();
  const metrics = {
    lineSpacing: lineSpacingSelect.value,
    spaceBefore: spaceBeforeInput.value,
    spaceAfter: spaceAfterInput.value,
  };
  blocks.forEach((block) => applyParagraphMetrics(block, metrics));
  markDirty();
  setStatus(`已更新段落间距（${blocks.length} 段）。`);
}

function applyParagraphAlignment(command) {
  restoreEditorSelection();
  const alignment = {
    justifyLeft: "left",
    justifyCenter: "center",
    justifyRight: "right",
    justifyFull: "justify",
  }[command];
  if (!alignment) return;
  const blocks = selectedBlockElements();
  if (!blocks.length) {
    setStatus("请先选中段落。");
    return;
  }
  recordUndoSnapshot();
  blocks.forEach((block) => {
    block.style.textAlign = alignment;
  });
  markDirty();
  refreshOutline();
  syncParagraphStyleSelect();
  const labels = { left: "左对齐", center: "居中", right: "右对齐", justify: "两端对齐" };
  setStatus(`已应用${labels[alignment]}（${blocks.length} 段）`);
}

function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function insertImage(dataUrl, name) {
  const html = `<p><img src="${dataUrl}" alt="${name}" data-name="${name}"></p>`;
  exec("insertHTML", html);
}

function replaceTag(element, tagName) {
  if (element.tagName === tagName) return element;
  const replacement = document.createElement(tagName.toLowerCase());
  Array.from(element.attributes).forEach((attr) => replacement.setAttribute(attr.name, attr.value));
  replacement.innerHTML = element.innerHTML;
  replacement.style.cssText = element.style.cssText;
  element.replaceWith(replacement);
  return replacement;
}

function ensureCellParagraphs(cell) {
  const elements = Array.from(cell.childNodes).filter((node) => node.nodeType === Node.ELEMENT_NODE);
  if (!elements.length && cell.textContent.trim()) {
    const p = document.createElement("p");
    p.textContent = cell.textContent;
    p.dataset.styleId = "Normal";
    cell.innerHTML = "";
    cell.appendChild(p);
    return;
  }
  Array.from(cell.childNodes).forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
      const p = document.createElement("p");
      p.textContent = node.textContent;
      p.dataset.styleId = "Normal";
      cell.replaceChild(p, node);
    }
  });
  if (!cell.querySelector("p, h1, h2, h3")) {
    const p = document.createElement("p");
    p.innerHTML = "<br>";
    p.dataset.styleId = "Normal";
    cell.appendChild(p);
  }
}

function normalizeTableStructure(table) {
  Array.from(table.querySelectorAll("td, th")).forEach((cell) => ensureCellParagraphs(cell));
}

function currentTableCell() {
  const selection = window.getSelection();
  if (!selection.rangeCount) return null;
  let node = selection.anchorNode;
  if (!node) return null;
  if (node.nodeType === Node.TEXT_NODE) node = node.parentNode;
  return node && node.closest ? node.closest("td, th") : null;
}

function currentTable() {
  return currentTableCell()?.closest("table") || null;
}

function makeEditableCell() {
  const td = document.createElement("td");
  const p = document.createElement("p");
  p.dataset.styleId = "Normal";
  p.innerHTML = "<br>";
  applyStyleVisuals(p, getStyleById("Normal"));
  td.appendChild(p);
  return td;
}

function placeCaretInTable(table) {
  if (!table) return;
  const paragraph = table.querySelector("td p, th p");
  if (!paragraph) return;
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(paragraph);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
  savedSelectionRange = range.cloneRange();
  editor.focus();
}

function insertTable() {
  const hadSelection = restoreEditorSelection();
  let anchorCell = currentTableCell();
  let anchorBlock = anchorCell ? null : currentBlockElement();
  if (anchorBlock === editor) {
    anchorBlock = null;
  }

  const rowInput = window.prompt("插入表格行数", "2");
  if (rowInput === null) {
    setStatus("已取消插入表格。");
    return;
  }
  const colInput = window.prompt("插入表格列数", "2");
  if (colInput === null) {
    setStatus("已取消插入表格。");
    return;
  }
  const rows = Math.max(Number(rowInput) || 2, 1);
  const cols = Math.max(Number(colInput) || 2, 1);
  if (hadSelection) {
    restoreEditorSelection();
  }
  recordUndoSnapshot();
  const table = document.createElement("table");
  table.className = "editor-table";
  const tbody = document.createElement("tbody");
  for (let r = 0; r < rows; r += 1) {
    const tr = document.createElement("tr");
    for (let c = 0; c < cols; c += 1) {
      tr.appendChild(makeEditableCell());
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  let inserted = false;
  if (anchorCell && anchorCell.isConnected) {
    const hostTable = anchorCell.closest("table");
    if (hostTable && hostTable.parentNode) {
      hostTable.parentNode.insertBefore(table, hostTable.nextSibling);
      inserted = true;
    }
  }

  if (!inserted && anchorBlock && anchorBlock.isConnected && anchorBlock.parentNode) {
    anchorBlock.parentNode.insertBefore(table, anchorBlock.nextSibling);
    inserted = true;
  }

  if (!inserted) {
    const fallbackCell = currentTableCell();
    if (fallbackCell) {
      const hostTable = fallbackCell.closest("table");
      if (hostTable && hostTable.parentNode) {
        hostTable.parentNode.insertBefore(table, hostTable.nextSibling);
        inserted = true;
      }
    }
  }

  if (!inserted) {
    const fallbackBlock = currentBlockElement();
    if (fallbackBlock && fallbackBlock !== editor && fallbackBlock.parentNode) {
      fallbackBlock.parentNode.insertBefore(table, fallbackBlock.nextSibling);
      inserted = true;
    }
  }

  if (!inserted) {
    editor.appendChild(table);
  }

  placeCaretInTable(table);
  markDirty();
  refreshOutline();
  setStatus(`已插入 ${rows} x ${cols} 表格。`);
}

function addTableRow() {
  const cell = currentTableCell();
  const table = currentTable();
  if (!cell || !table) {
    setStatus("请先把光标放在表格单元格中。");
    return;
  }
  recordUndoSnapshot();
  const referenceRow = cell.parentElement;
  const tr = document.createElement("tr");
  Array.from(referenceRow.cells).forEach(() => tr.appendChild(makeEditableCell()));
  referenceRow.after(tr);
  markDirty();
  setStatus("已添加表格行。");
}

function removeTableRow() {
  const cell = currentTableCell();
  const table = currentTable();
  if (!cell || !table) {
    setStatus("请先把光标放在表格单元格中。");
    return;
  }
  recordUndoSnapshot();
  const tbody = table.tBodies[0];
  if (tbody.rows.length <= 1) {
    table.remove();
    markDirty();
    setStatus("已删除表格");
    return;
  }
  cell.parentElement.remove();
  markDirty();
  setStatus("已删除表格行。");
}

function addTableColumn() {
  const cell = currentTableCell();
  const table = currentTable();
  if (!cell || !table) {
    setStatus("请先把光标放在表格单元格中。");
    return;
  }
  recordUndoSnapshot();
  const index = cell.cellIndex;
  Array.from(table.rows).forEach((row) => {
    const newCell = makeEditableCell();
    row.children[index].after(newCell);
  });
  markDirty();
  setStatus("已添加表格列。");
}

function removeTableColumn() {
  const cell = currentTableCell();
  const table = currentTable();
  if (!cell || !table) {
    setStatus("请先把光标放在表格单元格中。");
    return;
  }
  recordUndoSnapshot();
  const index = cell.cellIndex;
  const columnCount = table.rows[0]?.cells.length || 0;
  if (columnCount <= 1) {
    table.remove();
    markDirty();
    setStatus("已删除表格");
    return;
  }
  Array.from(table.rows).forEach((row) => row.children[index].remove());
  markDirty();
  setStatus("已删除表格列。");
}

function deleteCurrentTable() {
  const table = currentTable();
  if (!table) {
    setStatus("请先把光标放在表格中。");
    return;
  }
  recordUndoSnapshot();
  table.remove();
  markDirty();
  setStatus("已删除表格");
}

function blockAlignment(element) {
  const align = window.getComputedStyle(element).textAlign;
  if (align === "center") return "align_center";
  if (align === "right") return "align_right";
  if (align === "justify") return "align_justify";
  return "align_left";
}

function normalizeNumberFormat(value) {
  const raw = String(value || DEFAULT_NUMBER_FORMAT);
  return NUMBER_FORMATS.has(raw) ? raw : DEFAULT_NUMBER_FORMAT;
}

function normalizeNumberLevel(value) {
  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed)) return 0;
  return Math.max(Math.min(parsed, MAX_NUMBERING_LEVEL), 0);
}

function defaultLevelText(level) {
  const normalized = normalizeNumberLevel(level);
  return Array.from({ length: normalized + 1 }, (_, index) => `%${index + 1}`).join(".") + ".";
}

function nextNumberingListId() {
  numberingListSeed += 1;
  return `list-${numberingListSeed}`;
}

function normalizeNumberingData(numbering) {
  if (!numbering || typeof numbering !== "object" || numbering.enabled === false) {
    return null;
  }
  const level = normalizeNumberLevel(numbering.ilvl ?? numbering.level ?? 0);
  const listId = String(numbering.list_id || numbering.listId || numbering.num_id || "").trim() || nextNumberingListId();
  const start = Math.max(Number.parseInt(numbering.start, 10) || 1, 1);
  return {
    enabled: true,
    list_id: listId,
    ilvl: level,
    num_fmt: normalizeNumberFormat(numbering.num_fmt || numbering.numFormat || numbering.format),
    lvl_text: String(numbering.lvl_text || numbering.level_text || "").trim() || defaultLevelText(level),
    start,
  };
}

function clearNumberingData(element) {
  if (!element) return;
  delete element.dataset.numberingEnabled;
  delete element.dataset.numberingListId;
  delete element.dataset.numberingLevel;
  delete element.dataset.numberingFormat;
  delete element.dataset.numberingText;
  delete element.dataset.numberingStart;
  delete element.dataset.numberingLabel;
  element.classList.remove("numbered-paragraph");
  element.style.removeProperty("--numbering-level");
  element.style.removeProperty("--numbering-prefix-width");
  element.style.removeProperty("--numbering-indent-step");
}

function setNumberingData(element, rawNumbering) {
  if (!element) return null;
  const numbering = normalizeNumberingData(rawNumbering);
  if (!numbering) {
    clearNumberingData(element);
    return null;
  }
  element.dataset.numberingEnabled = "1";
  element.dataset.numberingListId = numbering.list_id;
  element.dataset.numberingLevel = String(numbering.ilvl);
  element.dataset.numberingFormat = numbering.num_fmt;
  element.dataset.numberingText = numbering.lvl_text;
  element.dataset.numberingStart = String(numbering.start);
  return numbering;
}

function numberingFromElement(element) {
  if (!element || element.dataset.numberingEnabled !== "1") return null;
  const numbering = normalizeNumberingData({
    enabled: true,
    list_id: element.dataset.numberingListId,
    ilvl: element.dataset.numberingLevel,
    num_fmt: element.dataset.numberingFormat,
    lvl_text: element.dataset.numberingText,
    start: element.dataset.numberingStart,
  });
  if (!numbering) return null;
  if (
    element.dataset.numberingListId !== numbering.list_id
    || element.dataset.numberingLevel !== String(numbering.ilvl)
    || element.dataset.numberingFormat !== numbering.num_fmt
    || element.dataset.numberingText !== numbering.lvl_text
    || element.dataset.numberingStart !== String(numbering.start)
  ) {
    setNumberingData(element, numbering);
  }
  return numbering;
}

function paragraphElementsInEditor() {
  return Array.from(editor.querySelectorAll("p, h1, h2, h3, div"));
}

function alphaIndex(value, upper = false) {
  let num = Math.max(Number.parseInt(value, 10) || 1, 1);
  let text = "";
  while (num > 0) {
    const mod = (num - 1) % 26;
    text = String.fromCharCode((upper ? 65 : 97) + mod) + text;
    num = Math.floor((num - 1) / 26);
  }
  return text;
}

function romanIndex(value, upper = true) {
  let num = Math.max(Number.parseInt(value, 10) || 1, 1);
  const map = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let result = "";
  map.forEach(([amount, symbol]) => {
    while (num >= amount) {
      result += symbol;
      num -= amount;
    }
  });
  return upper ? result : result.toLowerCase();
}

function formatIndex(value, format) {
  const normalized = normalizeNumberFormat(format);
  if (normalized === "upperLetter") return alphaIndex(value, true);
  if (normalized === "lowerLetter") return alphaIndex(value, false);
  if (normalized === "upperRoman") return romanIndex(value, true);
  if (normalized === "lowerRoman") return romanIndex(value, false);
  return String(Math.max(Number.parseInt(value, 10) || 1, 1));
}

function numberingLabelFromTemplate(template, counters, formats, level) {
  const raw = String(template || defaultLevelText(level));
  return raw.replace(/%(\d+)/g, (full, token) => {
    const index = Number.parseInt(token, 10) - 1;
    if (index < 0 || index > level) return "";
    const current = counters[index];
    if (!current) return "";
    return formatIndex(current, formats[index] || DEFAULT_NUMBER_FORMAT);
  });
}

function isHeadingStyleBlock(block) {
  if (!block) return false;
  if (["H1", "H2", "H3"].includes(block.tagName)) return true;
  const styleId = block.dataset.styleId || styleIdFromTag(block.tagName);
  const style = getStyleById(styleId);
  return [0, 1, 2].includes(style?.outline_level);
}

function numberingVisualLevel(block, numbering) {
  if (!numbering) return 0;
  return isHeadingStyleBlock(block) ? 0 : numbering.ilvl;
}

function fontSpecFromComputedStyle(style) {
  const fontStyle = style.fontStyle || "normal";
  const fontVariant = style.fontVariant || "normal";
  const fontWeight = style.fontWeight || "400";
  const fontSize = style.fontSize || "16px";
  const fontFamily = style.fontFamily || `${DEFAULT_FONT_FAMILY}, serif`;
  return `${fontStyle} ${fontVariant} ${fontWeight} ${fontSize} ${fontFamily}`;
}

function estimateNumberingLabelWidthPx(block, label) {
  const text = String(label || "");
  if (numberingMeasureContext) {
    numberingMeasureContext.font = fontSpecFromComputedStyle(window.getComputedStyle(block || editor));
  }
  const measured = numberingMeasureContext ? numberingMeasureContext.measureText(text).width : (text.length * 8);
  return Math.max(Math.ceil(measured) + 1, MIN_NUMBERING_PREFIX_PX);
}

function refreshNumberingVisuals() {
  const states = new Map();
  const visualItems = [];
  paragraphElementsInEditor().forEach((block) => {
    const numbering = numberingFromElement(block);
    if (!numbering) {
      if (block.dataset.numberingLabel) delete block.dataset.numberingLabel;
      if (block.classList.contains("numbered-paragraph")) block.classList.remove("numbered-paragraph");
      if (block.style.getPropertyValue("--numbering-level")) block.style.removeProperty("--numbering-level");
      if (block.style.getPropertyValue("--numbering-prefix-width")) block.style.removeProperty("--numbering-prefix-width");
      if (block.style.getPropertyValue("--numbering-indent-step")) block.style.removeProperty("--numbering-indent-step");
      return;
    }

    if (!states.has(numbering.list_id)) {
      states.set(numbering.list_id, {
        counters: Array(MAX_NUMBERING_LEVEL + 1).fill(0),
        starts: Array(MAX_NUMBERING_LEVEL + 1).fill(1),
        formats: Array(MAX_NUMBERING_LEVEL + 1).fill(DEFAULT_NUMBER_FORMAT),
      });
    }
    const state = states.get(numbering.list_id);
    state.starts[numbering.ilvl] = numbering.start;
    state.formats[numbering.ilvl] = numbering.num_fmt;

    for (let index = 0; index < numbering.ilvl; index += 1) {
      if (state.counters[index] <= 0) {
        state.counters[index] = Math.max(state.starts[index] || 1, 1);
      }
    }
    if (state.counters[numbering.ilvl] <= 0) {
      state.counters[numbering.ilvl] = Math.max(numbering.start, 1);
    } else {
      state.counters[numbering.ilvl] += 1;
    }
    for (let index = numbering.ilvl + 1; index <= MAX_NUMBERING_LEVEL; index += 1) {
      state.counters[index] = 0;
    }

    const label = numberingLabelFromTemplate(numbering.lvl_text, state.counters, state.formats, numbering.ilvl)
      || `${formatIndex(state.counters[numbering.ilvl], numbering.num_fmt)}.`;
    visualItems.push({
      block,
      numbering,
      label,
      visualLevel: numberingVisualLevel(block, numbering),
      isHeading: isHeadingStyleBlock(block),
      prefixWidthPx: estimateNumberingLabelWidthPx(block, label),
    });
  });

  let maxHeadingPrefixWidth = 0;
  visualItems.forEach((item) => {
    if (!item.isHeading) return;
    maxHeadingPrefixWidth = Math.max(maxHeadingPrefixWidth, item.prefixWidthPx);
  });

  visualItems.forEach((item) => {
    const levelText = String(item.visualLevel);
    const prefixWidthPx = item.isHeading
      ? (maxHeadingPrefixWidth || item.prefixWidthPx)
      : item.prefixWidthPx;
    const prefixWidth = `${prefixWidthPx}px`;
    if (item.block.dataset.numberingLabel !== item.label) item.block.dataset.numberingLabel = item.label;
    if (!item.block.classList.contains("numbered-paragraph")) item.block.classList.add("numbered-paragraph");
    if (item.block.style.getPropertyValue("--numbering-level") !== levelText) {
      item.block.style.setProperty("--numbering-level", levelText);
    }
    if (item.block.style.getPropertyValue("--numbering-prefix-width") !== prefixWidth) {
      item.block.style.setProperty("--numbering-prefix-width", prefixWidth);
    }
    if (item.block.style.getPropertyValue("--numbering-indent-step") !== `${NUMBERING_LEVEL_INDENT_PX}px`) {
      item.block.style.setProperty("--numbering-indent-step", `${NUMBERING_LEVEL_INDENT_PX}px`);
    }
  });
}

function inferListIdFromSelection(blocks) {
  const existing = blocks.map((block) => numberingFromElement(block)).find(Boolean);
  if (existing && existing.list_id) return existing.list_id;
  const ordered = paragraphElementsInEditor();
  const firstBlock = blocks[0];
  const firstIndex = ordered.indexOf(firstBlock);
  if (firstIndex > 0) {
    for (let index = firstIndex - 1; index >= 0; index -= 1) {
      const candidate = numberingFromElement(ordered[index]);
      if (candidate && candidate.list_id) {
        return candidate.list_id;
      }
      if ((ordered[index].textContent || "").trim()) {
        break;
      }
    }
  }
  return nextNumberingListId();
}

function syncNumberingControls() {
  const block = currentBlockElement();
  const numbering = block ? numberingFromElement(block) : null;
  if (numberFormatSelect) {
    numberFormatSelect.value = numbering ? normalizeNumberFormat(numbering.num_fmt) : DEFAULT_NUMBER_FORMAT;
  }
  if (toggleNumberingBtn) {
    toggleNumberingBtn.classList.toggle("is-active", Boolean(numbering));
  }
}

function toggleNumbering() {
  restoreEditorSelection();
  const blocks = selectedBlockElements();
  if (!blocks.length) {
    setStatus("请先选中段落，再设置编号。");
    return;
  }
  const allNumbered = blocks.every((block) => Boolean(numberingFromElement(block)));
  recordUndoSnapshot();
  if (allNumbered) {
    blocks.forEach((block) => clearNumberingData(block));
    numberingDirty = true;
    refreshOutline();
    markDirty();
    setStatus(`已取消编号（${blocks.length} 段）。`);
    return;
  }
  const targetFormat = normalizeNumberFormat(numberFormatSelect?.value);
  const listId = inferListIdFromSelection(blocks);
  blocks.forEach((block) => {
    const existing = numberingFromElement(block);
    const level = normalizeNumberLevel(existing?.ilvl ?? 0);
    setNumberingData(block, {
      enabled: true,
      list_id: existing?.list_id || listId,
      ilvl: level,
      num_fmt: existing?.num_fmt || targetFormat,
      lvl_text: existing?.lvl_text || defaultLevelText(level),
      start: existing?.start || 1,
    });
  });
  numberingDirty = true;
  refreshOutline();
  markDirty();
  setStatus(`已应用编号（${blocks.length} 段）。`);
}

function updateNumberingLevel(delta) {
  restoreEditorSelection();
  const blocks = selectedBlockElements();
  const targets = blocks.filter((block) => Boolean(numberingFromElement(block)));
  if (!targets.length) {
    setStatus("请先选中已编号段落。");
    return;
  }
  recordUndoSnapshot();
  targets.forEach((block) => {
    const current = numberingFromElement(block);
    const level = normalizeNumberLevel((current?.ilvl ?? 0) + delta);
    setNumberingData(block, {
      ...current,
      ilvl: level,
      lvl_text: defaultLevelText(level),
    });
  });
  numberingDirty = true;
  refreshOutline();
  markDirty();
  setStatus(`已更新编号级别（${targets.length} 段）。`);
}

function applyNumberFormatToSelection() {
  const format = normalizeNumberFormat(numberFormatSelect?.value);
  const blocks = selectedBlockElements();
  const targets = blocks.filter((block) => Boolean(numberingFromElement(block)));
  if (!targets.length) {
    syncNumberingControls();
    return;
  }
  recordUndoSnapshot();
  targets.forEach((block) => {
    const current = numberingFromElement(block);
    if (!current) return;
    setNumberingData(block, {
      ...current,
      num_fmt: format,
    });
  });
  numberingDirty = true;
  refreshOutline();
  markDirty();
  setStatus(`已更新编号格式（${targets.length} 段）。`);
}

function descriptorFromStyle(style) {
  const rawFamily = style.fontFamily || DEFAULT_FONT_FAMILY;
  const family = rawFamily.includes("Times New Roman") && rawFamily.includes("SimSun")
    ? DEFAULT_FONT_FAMILY
    : rawFamily.split(",")[0].replace(/["']/g, "").trim() || DEFAULT_FONT_FAMILY;
  const size = Math.max(Math.round(parseFloat(style.fontSize || "16") * 0.75), 1);
  const weight = parseInt(style.fontWeight || "400", 10);
  return [
    family,
    size,
    weight >= 600 || style.fontWeight === "bold",
    style.fontStyle === "italic",
    style.textDecorationLine.includes("underline"),
    normalizeBackgroundColor(style.backgroundColor || ""),
  ];
}

function inlineFontElementForNode(node, block) {
  let current = node;
  while (current && current !== block) {
    if (
      current.nodeType === Node.ELEMENT_NODE
      && current.matches
      && (
        current.matches("font")
        || current.matches("span[style*='font'], span[style*='text-decoration'], span[style*='background']")
      )
    ) {
      return current;
    }
    current = current.parentNode;
  }
  if (block && block.querySelector) {
    return block.querySelector("font, span[style*='font'], span[style*='text-decoration'], span[style*='background']");
  }
  return null;
}

function selectionDescriptorOrBlockDescriptor(block) {
  const range = activeEditorRange();
  if (range) {
    let node = range.startContainer;
    if (node && node.nodeType === Node.TEXT_NODE) node = node.parentNode;
    if (node && block && block.contains(node)) {
      const inlineNode = inlineFontElementForNode(node, block);
      return descriptorFromStyle(window.getComputedStyle(inlineNode || node));
    }
  }
  return descriptorFromStyle(window.getComputedStyle(block));
}

function summarizeCurrentSelectionForDebug() {
  const block = currentBlockElement();
  const range = activeEditorRange();
  const descriptor = block ? selectionDescriptorOrBlockDescriptor(block) : null;
  return {
    hasRange: Boolean(range),
    blockTag: block?.tagName || null,
    blockStyleId: block?.dataset?.styleId || null,
    blockText: (block?.textContent || "").slice(0, 80),
    descriptor,
    fontFamilyControl: fontFamily?.value || null,
    fontSizeControl: fontSize?.value || null,
  };
}

function descriptorCssTokens(descriptor) {
  const [family, size, bold, italic, underline, background] = cloneDescriptor(descriptor);
  return {
    family: String(family || DEFAULT_FONT_FAMILY).split(",")[0].replace(/["']/g, "").trim().toLowerCase(),
    sizePx: Math.max(Number(size) / 0.75, 1),
    weight: bold ? 700 : 400,
    style: italic ? "italic" : "normal",
    decoration: underline ? "underline" : "none",
    background: normalizeBackgroundColor(background || ""),
  };
}

function normalizeFontWeight(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "normal") return 400;
  if (raw === "bold") return 700;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) ? parsed : 400;
}

function normalizeTextDecoration(value) {
  const raw = String(value || "").toLowerCase();
  return raw.includes("underline") ? "underline" : "none";
}

function runLooksLikeDescriptor(runElement, descriptor) {
  const expected = descriptorCssTokens(descriptor);
  const style = window.getComputedStyle(runElement);
  const family = String(style.fontFamily || "").split(",")[0].replace(/["']/g, "").trim().toLowerCase();
  const sizePx = parseFloat(style.fontSize || "0");
  const weight = normalizeFontWeight(style.fontWeight);
  const fontStyle = String(style.fontStyle || "normal").toLowerCase();
  const decoration = normalizeTextDecoration(style.textDecorationLine || style.textDecoration);
  const background = normalizeBackgroundColor(style.backgroundColor || "");
  if (family !== expected.family) return false;
  if (!Number.isFinite(sizePx) || Math.abs(sizePx - expected.sizePx) > 0.8) return false;
  if (weight !== expected.weight) return false;
  if (fontStyle !== expected.style) return false;
  if (decoration !== expected.decoration) return false;
  if (background !== expected.background) return false;
  return true;
}

function applyDescriptorToRun(runElement, descriptor) {
  const [family, size, bold, italic, underline, background] = cloneDescriptor(descriptor);
  runElement.style.fontFamily = family;
  runElement.style.fontSize = `${Math.max(Number(size) / 0.75, 1)}px`;
  runElement.style.fontWeight = bold ? "700" : "400";
  runElement.style.fontStyle = italic ? "italic" : "normal";
  runElement.style.textDecoration = underline ? "underline" : "none";
  if (background) {
    runElement.style.backgroundColor = background;
  } else {
    runElement.style.removeProperty("background-color");
  }
}

function inlineStyledRuns(block) {
  if (!block || !block.querySelectorAll) return [];
  return Array.from(block.querySelectorAll("span, font")).filter((node) => {
    if (node.tagName === "FONT") return true;
    const styleAttr = node.getAttribute("style") || "";
    return /font|text-decoration|background/i.test(styleAttr);
  });
}

function applyFontFamilyToBlock(block, family) {
  if (!block || !family) return false;
  block.style.fontFamily = family;
  const runs = inlineStyledRuns(block);
  if (!runs.length) {
    const text = block.textContent || "";
    block.innerHTML = "";
    const span = document.createElement("span");
    span.textContent = text;
    span.style.fontFamily = family;
    block.appendChild(span);
    return true;
  }
  runs.forEach((run) => {
    run.style.fontFamily = family;
  });
  return true;
}

function applyFontSizeToBlock(block, pointSize) {
  if (!block || !Number.isFinite(pointSize) || pointSize <= 0) return false;
  const px = `${pointSize / 0.75}px`;
  block.style.fontSize = px;
  const runs = inlineStyledRuns(block);
  if (!runs.length) {
    const text = block.textContent || "";
    block.innerHTML = "";
    const span = document.createElement("span");
    span.textContent = text;
    span.style.fontSize = px;
    block.appendChild(span);
    return true;
  }
  runs.forEach((run) => {
    run.style.fontSize = px;
  });
  return true;
}

function syncStyledRunsForStyleUpdate(block, previousDescriptor, nextDescriptor) {
  if (!block) return;
  const runs = inlineStyledRuns(block);
  if (!runs.length) {
    Array.from(block.childNodes).forEach((node) => {
      if (node.nodeType !== Node.TEXT_NODE || !node.textContent) return;
      const span = document.createElement("span");
      span.textContent = node.textContent;
      applyDescriptorToRun(span, nextDescriptor);
      node.replaceWith(span);
    });
    return;
  }
  runs.forEach((run) => {
    applyDescriptorToRun(run, nextDescriptor);
  });
}

function normalizeBackgroundColor(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw || raw === "transparent") return "";
  if (raw.startsWith("#")) {
    if (raw.length === 4) {
      return `#${raw[1]}${raw[1]}${raw[2]}${raw[2]}${raw[3]}${raw[3]}`;
    }
    return raw.length === 7 ? raw : "";
  }
  const match = raw.match(/^rgba?\(([^)]+)\)$/);
  if (match) {
    const parts = match[1].split(",").map((item) => item.trim());
    const r = Number(parts[0]);
    const g = Number(parts[1]);
    const b = Number(parts[2]);
    const a = parts.length > 3 ? Number(parts[3]) : 1;
    if ([r, g, b].some((n) => Number.isNaN(n))) return "";
    if (a === 0) return "";
    const hex = [r, g, b].map((n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0")).join("");
    return `#${hex}`;
  }
  return "";
}

function mmToTwips(mm) {
  const value = Number(mm);
  if (!Number.isFinite(value)) return 0;
  return Math.round((value * 1440) / 25.4);
}

function twipsToMm(twips) {
  const value = Number(twips);
  if (!Number.isFinite(value)) return 0;
  return Math.round((value * 25.4) / 1440);
}

function mmToPx(mm) {
  const value = Number(mm);
  if (!Number.isFinite(value)) return 0;
  return (value * 96) / 25.4;
}

function applyPageSizeToEditor() {
  const widthMm = twipsToMm(pageSize.widthTwips);
  const heightMm = twipsToMm(pageSize.heightTwips);
  const widthPx = Math.max(mmToPx(widthMm), 420);
  const heightPx = Math.max(mmToPx(heightMm), 420);
  editor.style.width = `${Math.round(widthPx)}px`;
  editor.style.minHeight = `${Math.round(heightPx)}px`;
}

function syncPageSizeControls() {
  if (!pageSizeSelect || !pageWidthInput || !pageHeightInput) return;
  const widthMm = twipsToMm(pageSize.widthTwips);
  const heightMm = twipsToMm(pageSize.heightTwips);
  pageWidthInput.value = String(widthMm || 210);
  pageHeightInput.value = String(heightMm || 297);
  const isA4 = widthMm === 210 && heightMm === 297;
  const isLetter = widthMm === 216 && heightMm === 279;
  pageSizeSelect.value = isA4 ? "A4" : isLetter ? "Letter" : "Custom";
}

function setPageSizeFromMm(widthMm, heightMm) {
  const width = Math.max(Number(widthMm) || 0, 50);
  const height = Math.max(Number(heightMm) || 0, 50);
  pageSize = { widthTwips: mmToTwips(width), heightTwips: mmToTwips(height) };
  applyPageSizeToEditor();
  syncPageSizeControls();
  markDirty();
}

function applyStyleVisuals(element, style) {
  if (!style) return element;
  let target = replaceTag(element, tagNameFromStyle(style));
  target.dataset.styleId = style.id;
  target.style.fontFamily = style.descriptor[0];
  target.style.fontSize = `${style.descriptor[1] / 0.75}px`;
  target.style.fontWeight = style.descriptor[2] ? "700" : "400";
  target.style.fontStyle = style.descriptor[3] ? "italic" : "normal";
  target.style.textDecoration = style.descriptor[4] ? "underline" : "none";
  target.style.textAlign = style.alignment || "left";
  applyParagraphMetrics(target, {
    lineSpacing: style.line_spacing || DEFAULT_LINE_SPACING,
    spaceBefore: style.space_before ?? DEFAULT_PARAGRAPH_SPACING,
    spaceAfter: style.space_after ?? 0,
  });
  return target;
}

function currentBlockElement() {
  const range = activeEditorRange();
  if (!range) return null;
  let node = range.startContainer;
  if (!node) return null;
  if (node.nodeType === Node.TEXT_NODE) node = node.parentNode;
  return node && node.closest ? node.closest("p, h1, h2, h3, div") : null;
}

function blockElementFromNode(node) {
  if (!node) return null;
  let target = node;
  if (target.nodeType === Node.TEXT_NODE) target = target.parentNode;
  return target && target.closest ? target.closest("p, h1, h2, h3, div") : null;
}

function collectBlocksBetween(startBlock, endBlock) {
  if (!startBlock || !endBlock || !nodeInEditor(startBlock) || !nodeInEditor(endBlock)) return [];
  let start = startBlock;
  let end = endBlock;
  const position = start.compareDocumentPosition(end);
  if (position & Node.DOCUMENT_POSITION_PRECEDING) {
    start = endBlock;
    end = startBlock;
  }
  const blocks = [];
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_ELEMENT, {
    acceptNode(node) {
      return ["P", "H1", "H2", "H3", "DIV"].includes(node.tagName)
        ? NodeFilter.FILTER_ACCEPT
        : NodeFilter.FILTER_SKIP;
    },
  });
  let current = walker.nextNode();
  let collecting = false;
  while (current) {
    if (current === start) collecting = true;
    if (collecting) blocks.push(current);
    if (current === end) break;
    current = walker.nextNode();
  }
  return blocks;
}

function selectedBlockElements() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return [];
  if (selection.isCollapsed) {
    const block = currentBlockElement();
    return block ? [block] : [];
  }
  const range = selection.getRangeAt(0);
  const startBlock = blockElementFromNode(range.startContainer);
  const endBlock = blockElementFromNode(range.endContainer);
  if (startBlock && endBlock && startBlock !== endBlock) {
    const between = collectBlocksBetween(startBlock, endBlock);
    if (between.length) return between;
  }
  const selected = Array.from(editor.querySelectorAll("p, h1, h2, h3, div")).filter((block) => {
    try {
      return range.intersectsNode(block);
    } catch {
      return false;
    }
  });
  if (selected.length) return selected;
  const block = currentBlockElement();
  return block ? [block] : [];
}

function syncParagraphStyleSelect() {
  const block = currentBlockElement();
  if (!block) {
    syncNumberingControls();
    return;
  }
  const styleId = block.dataset.styleId || styleIdFromTag(block.tagName);
  paragraphStyleSelect.value = getStyleById(styleId)?.id || "Normal";
  syncParagraphMetricsControls();
  syncNumberingControls();
}

function ensureSelectHasOption(select, value) {
  if (!select || !value) return;
  const normalized = String(value);
  const exists = Array.from(select.options).some((option) => option.value === normalized);
  if (!exists) {
    const option = document.createElement("option");
    option.value = normalized;
    option.textContent = normalized;
    select.appendChild(option);
  }
}

function syncTextFormatControls() {
  const block = currentBlockElement();
  if (!block) {
    debugLog("syncTextFormatControls:skip", summarizeCurrentSelectionForDebug());
    return;
  }
  const descriptor = selectionDescriptorOrBlockDescriptor(block);
  const family = fontFamilyForControl(descriptor[0]);
  const size = Math.max(Number(descriptor[1]) || 12, 1);
  ensureSelectHasOption(fontFamily, family);
  if (fontFamily) {
    fontFamily.value = family;
  }
  if (fontSize) {
    fontSize.value = String(size);
  }
  debugLog("syncTextFormatControls", {
    ...summarizeCurrentSelectionForDebug(),
    appliedFamily: family,
    appliedSize: size,
  });
}

function syncSelectionUi() {
  captureEditorSelection();
  debugLog("syncSelectionUi", summarizeCurrentSelectionForDebug());
  syncTextFormatControls();
  syncParagraphStyleSelect();
}

function normalizeEditorStructure() {
  const children = Array.from(editor.childNodes);
  children.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
      const p = document.createElement("p");
      p.textContent = node.textContent;
      p.dataset.styleId = "Normal";
      editor.replaceChild(p, node);
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }
    if (node.tagName === "TABLE") {
      normalizeTableStructure(node);
      return;
    }
    if (!["P", "H1", "H2", "H3", "DIV"].includes(node.tagName)) {
      const p = document.createElement("p");
      p.innerHTML = node.outerHTML;
      p.dataset.styleId = "Normal";
      editor.replaceChild(p, node);
      return;
    }
    let target = node;
    if (node.tagName === "DIV") {
      const p = document.createElement("p");
      p.innerHTML = node.innerHTML;
      p.style.cssText = node.style.cssText;
      if (node.dataset.styleId) p.dataset.styleId = node.dataset.styleId;
      editor.replaceChild(p, node);
      target = p;
    }
    if (!target.dataset.styleId) {
      target.dataset.styleId = styleIdFromTag(target.tagName);
    }
  });
}

function slugify(text, fallback) {
  const slug = text
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
  return slug || fallback;
}

function refreshOutline() {
  normalizeEditorStructure();
  refreshNumberingVisuals();
  outline.innerHTML = "";
  const headings = Array.from(editor.querySelectorAll("h1, h2, h3"))
    .filter((heading) => outlineFilter[Number(heading.tagName.slice(1))]);
  if (!headings.length) {
    const empty = document.createElement("div");
    empty.textContent = "暂无标题导航";
    empty.className = "outline-item";
    outline.appendChild(empty);
    syncParagraphStyleSelect();
    return;
  }
  headings.forEach((heading, index) => {
    const level = Number(heading.tagName.slice(1));
    const text = heading.textContent.trim() || `标题 ${index + 1}`;
    const headingId = slugify(text, `heading-${index + 1}`);
    heading.dataset.headingId = headingId;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `outline-item level-${level}`;
    button.textContent = text;
    button.addEventListener("click", () => {
      const headingLineHeight = parseFloat(window.getComputedStyle(heading).lineHeight || "24") || 24;
      const editorLineHeight = parseFloat(window.getComputedStyle(editor).lineHeight || "24") || 24;
      const baseLineHeight = headingLineHeight || editorLineHeight || 24;
      const extraOffset = baseLineHeight * 30;
      const stageRect = pageStage.getBoundingClientRect();
      const headingRect = heading.getBoundingClientRect();
      const baseTop = pageStage.scrollTop + (headingRect.top - stageRect.top);
      const top = baseTop - Math.max((pageStage.clientHeight - heading.offsetHeight) / 2, 24) - extraOffset;
      pageStage.scrollTo({ top: Math.max(top, 0), behavior: "smooth" });
      const range = document.createRange();
      range.selectNodeContents(heading);
      range.collapse(false);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      editor.focus();
      syncParagraphStyleSelect();
    });
    outline.appendChild(button);
  });
  syncParagraphStyleSelect();
}

function supportsDirectoryAccess() {
  return typeof window.showDirectoryPicker === "function";
}

function formatDateTime(value) {
  if (!value) return "";
  try {
    return new Date(value).toLocaleString("zh-CN");
  } catch {
    return "";
  }
}

function createHistoryItem(title, meta, actions) {
  const row = document.createElement("div");
  row.className = "history-item";
  const info = document.createElement("div");
  const titleEl = document.createElement("div");
  titleEl.className = "history-item-title";
  titleEl.textContent = title || "未命名";
  info.appendChild(titleEl);
  if (meta) {
    const metaEl = document.createElement("div");
    metaEl.className = "history-item-meta";
    metaEl.textContent = meta;
    info.appendChild(metaEl);
  }
  const actionsWrap = document.createElement("div");
  actionsWrap.className = "history-item-actions";
  (actions || []).forEach((action) => actionsWrap.appendChild(action));
  row.append(info, actionsWrap);
  return row;
}

async function renderRecentList() {
  recentFileList.innerHTML = "";
  let records = [];
  try {
    records = await getStoreRecords(STORE_RECENT);
  } catch (error) {
    recentFileList.textContent = error.message || String(error);
    return;
  }
  records.sort((a, b) => (b.lastOpened || 0) - (a.lastOpened || 0));
  if (!records.length) {
    recentFileList.textContent = "暂无记录";
    return;
  }
  records.forEach((record) => {
    const openBtn = document.createElement("button");
    openBtn.type = "button";
    openBtn.textContent = "打开";
    openBtn.addEventListener("click", async () => {
      try {
        if (!record.handle) {
          setStatus("无法打开此记录，请手动选择文件。");
          return;
        }
        await openDocxFromHandle(record.handle);
        setStatus(`已打开 ${record.name}`);
      } catch (error) {
        handleAsyncError(error);
      }
    });
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "移除";
    removeBtn.addEventListener("click", async () => {
      await deleteStoreRecord(STORE_RECENT, record.id);
      renderRecentList();
    });
    recentFileList.appendChild(createHistoryItem(record.name, formatDateTime(record.lastOpened), [openBtn, removeBtn]));
  });
}

async function renderFavoriteList() {
  favoriteDirList.innerHTML = "";
  let records = [];
  try {
    records = await getStoreRecords(STORE_FAVORITE);
  } catch (error) {
    favoriteDirList.textContent = error.message || String(error);
    return;
  }
  if (!records.length) {
    favoriteDirList.textContent = "暂无常用目录";
    return;
  }
  records.forEach((record) => {
    const browseBtn = document.createElement("button");
    browseBtn.type = "button";
    browseBtn.textContent = "浏览";
    browseBtn.addEventListener("click", async () => {
      if (!record.handle) {
        setStatus("目录句柄已失效，请重新添加。");
        return;
      }
      const files = await listDocxInDirectory(record.handle);
      renderDirFiles(files, record.name);
    });
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "移除";
    removeBtn.addEventListener("click", async () => {
      await deleteStoreRecord(STORE_FAVORITE, record.id);
      renderFavoriteList();
    });
    favoriteDirList.appendChild(createHistoryItem(record.name, "", [browseBtn, removeBtn]));
  });
}

function renderDirFiles(files, label) {
  dirFileList.innerHTML = "";
  if (!files.length) {
    dirFileList.textContent = label ? `未在 ${label} 中找到 docx 文件` : "未找到 docx 文件";
    return;
  }
  files.forEach((file) => {
    const openBtn = document.createElement("button");
    openBtn.type = "button";
    openBtn.textContent = "打开";
    openBtn.addEventListener("click", async () => {
      try {
        await openDocxFromHandle(file.handle);
        setStatus(`已打开 ${file.name}`);
      } catch (error) {
        handleAsyncError(error);
      }
    });
    dirFileList.appendChild(createHistoryItem(file.path || file.name, "", [openBtn]));
  });
}

async function listDocxInDirectory(dirHandle, prefix = "") {
  const results = [];
  for await (const [name, handle] of dirHandle.entries()) {
    if (handle.kind === "file") {
      if (name.toLowerCase().endsWith(".docx")) {
        results.push({ name, path: `${prefix}${name}`, handle });
      }
      continue;
    }
    if (handle.kind === "directory") {
      const nested = await listDocxInDirectory(handle, `${prefix}${name}/`);
      results.push(...nested);
    }
  }
  return results;
}

function openHistoryModal() {
  historyModal.classList.remove("hidden");
  renderRecentList();
  renderFavoriteList();
}

function closeHistoryModal() {
  historyModal.classList.add("hidden");
}

async function openDocxFromHandle(fileHandle) {
  await ensureHandlePermission(fileHandle, false);
  const file = await fileHandle.getFile();
  currentFileHandle = fileHandle;
  currentFileName = file.name || "mini-docx.docx";
  await openDocx(file);
  try {
    await recordRecentFile(fileHandle);
  } catch {
    // Ignore storage errors for recent list.
  }
  updateFileUiState();
}

async function openDocxPicker() {
  if (!confirmDiscardChanges()) return;
  if (!supportsFileSystemAccess()) {
    openDocxInput.click();
    return;
  }
  const [fileHandle] = await window.showOpenFilePicker({
    multiple: false,
    types: [{ description: "Word 文档", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }],
  });
  if (!fileHandle) return;
  await openDocxFromHandle(fileHandle);
}

function mergeRuns(runs) {
  const merged = [];
  runs.forEach((run) => {
    if (!run.text) return;
    const prev = merged[merged.length - 1];
    if (prev && JSON.stringify(prev.descriptor) === JSON.stringify(run.descriptor)) {
      prev.text += run.text;
    } else {
      merged.push(run);
    }
  });
  return merged;
}

function collectRuns(node, inheritedStyle, bucket) {
  if (node.nodeType === Node.TEXT_NODE) {
    if (node.textContent) {
      bucket.push({ text: node.textContent, descriptor: inheritedStyle });
    }
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return;
  }
  if (node.tagName === "IMG") {
    return;
  }
  const nextStyle = descriptorFromStyle(window.getComputedStyle(node));
  Array.from(node.childNodes).forEach((child) => collectRuns(child, nextStyle, bucket));
}

function paragraphsFromContainer(container) {
  const blocks = [];
  Array.from(container.children).forEach((element) => {
    if (!["P", "H1", "H2", "H3", "DIV"].includes(element.tagName)) return;
    const styleId = element.dataset.styleId || styleIdFromTag(element.tagName);
    const style = getStyleById(styleId);
    const numbering = numberingFromElement(element);
    const runs = [];
    const baseStyle = descriptorFromStyle(window.getComputedStyle(element));
    Array.from(element.childNodes).forEach((node) => collectRuns(node, baseStyle, runs));
    const block = {
      type: "paragraph",
      style: blockStyleKey(style),
      style_id: styleId,
      style_name: style ? style.name : styleId,
      alignment: blockAlignment(element),
      line_spacing: Math.max(parseFloat(element.style.lineHeight || window.getComputedStyle(element).lineHeight) / parseFloat(window.getComputedStyle(element).fontSize || "16"), 1) || DEFAULT_LINE_SPACING,
      space_before: Math.max(Number(element.dataset.spaceBefore || 0), 0),
      space_after: Math.max(Number(element.dataset.spaceAfter || parseFloat(window.getComputedStyle(element).marginBottom || "0")), 0),
      runs: mergeRuns(runs),
    };
    if (numbering) {
      block.numbering = numbering;
    }
    blocks.push(block);
  });
  if (!blocks.length) {
    blocks.push({ type: "paragraph", style: "normal", style_id: "Normal", style_name: "Normal", alignment: "align_left", runs: [] });
  }
  return blocks;
}

function editorToDocument() {
  normalizeEditorStructure();
  const blocks = [];
  Array.from(editor.children).forEach((element) => {
    if (element.tagName === "TABLE") {
      normalizeTableStructure(element);
      const rows = Array.from(element.rows).map((row) => Array.from(row.cells).map((cell) => ({
        width: Math.max(Math.round(cell.getBoundingClientRect().width * 15), 1200),
        paragraphs: paragraphsFromContainer(cell),
      })));
      blocks.push({ type: "table", rows });
      return;
    }
    const onlyImage = element.children.length === 1 && element.querySelector("img") && element.textContent.trim() === "";
    if (onlyImage) {
      const img = element.querySelector("img");
      blocks.push({
        type: "image",
        name: img.dataset.name || img.alt || "image.png",
        mime: (img.src.match(/^data:([^;]+);base64,/) || [])[1] || "image/png",
        data_url: img.src,
        width_px: Math.round(img.getBoundingClientRect().width) || img.naturalWidth || 320,
        height_px: Math.round(img.getBoundingClientRect().height) || img.naturalHeight || 180,
      });
      return;
    }

    blocks.push(...paragraphsFromContainer({ children: [element] }));
  });
  const payload = {
    blocks,
    styles: currentStyles,
    page: {
      width_twips: pageSize.widthTwips,
      height_twips: pageSize.heightTwips,
    },
  };
  const meta = cloneDocxMeta(docxMeta);
  if (meta) {
    meta.styles_dirty = Boolean(stylesDirty);
    meta.numbering_dirty = Boolean(numberingDirty);
    meta.content_dirty = Boolean(isDirty);
    meta.page_dirty = false;
    payload._docx_meta = meta;
  }
  return payload;
}

function spansFromRuns(runs) {
  return (runs || []).map((run) => {
    const [family, size, bold, italic, underline, background] = cloneDescriptor(run.descriptor);
    const span = document.createElement("span");
    span.textContent = run.text;
    span.style.fontFamily = family;
    span.style.fontSize = `${Math.max(Number(size) / 0.75, 1)}px`;
    span.style.fontWeight = bold ? "700" : "400";
    span.style.fontStyle = italic ? "italic" : "normal";
    span.style.textDecoration = underline ? "underline" : "none";
    if (background) {
      span.style.backgroundColor = background;
    }
    return span.outerHTML.replace(/\n/g, "<br>");
  }).join("");
}

function renderParagraphBlock(block, parent) {
  const style = getStyleById(block.style_id || styleIdFromBlockStyleKey(block.style));
  const el = document.createElement(tagNameFromStyle(style).toLowerCase());
  el.dataset.styleId = (style && style.id) || "Normal";
  el.style.textAlign = { align_left: "left", align_center: "center", align_right: "right", align_justify: "justify" }[block.alignment] || (style ? style.alignment : "left");
  el.innerHTML = spansFromRuns(block.runs);
  applyStyleVisuals(el, getStyleById(el.dataset.styleId));
  el.style.textAlign = { align_left: "left", align_center: "center", align_right: "right", align_justify: "justify" }[block.alignment] || el.style.textAlign;
  applyParagraphMetrics(el, {
    lineSpacing: block.line_spacing || DEFAULT_LINE_SPACING,
    spaceBefore: block.space_before ?? DEFAULT_PARAGRAPH_SPACING,
    spaceAfter: block.space_after ?? 0,
  });
  setNumberingData(el, block.numbering);
  if (!el.textContent.trim()) {
    el.innerHTML = "<br>";
  }
  parent.appendChild(el);
}

function renderTableBlock(block) {
  const table = document.createElement("table");
  table.className = "editor-table";
  const tbody = document.createElement("tbody");
  (block.rows || []).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      (cell.paragraphs || []).forEach((paragraph) => renderParagraphBlock(paragraph, td));
      ensureCellParagraphs(td);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  return table;
}

function loadDocument(documentData) {
  isLoadingDocument = true;
  clearFormatPainter();
  docxMeta = cloneDocxMeta(documentData._docx_meta);
  stylesDirty = false;
  numberingDirty = false;
  currentStyles = normalizeStyles(documentData.styles);
  populateStyleSelect();
  pageSize = {
    widthTwips: mmToTwips(OPENED_DOCUMENT_WIDTH_MM),
    heightTwips: Number(documentData.page?.height_twips) || mmToTwips(DEFAULT_DOCUMENT_HEIGHT_MM),
  };
  applyPageSizeToEditor();
  syncPageSizeControls();
  editor.innerHTML = "";
  (documentData.blocks || []).forEach((block) => {
    if (block.type === "table") {
      editor.appendChild(renderTableBlock(block));
      return;
    }
    if (block.type === "image") {
      const p = document.createElement("p");
      const img = document.createElement("img");
      p.dataset.styleId = "Normal";
      img.src = block.data_url;
      img.alt = block.name || "image";
      img.dataset.name = block.name || "image.png";
      if (block.width_px) img.style.width = `${block.width_px}px`;
      if (block.height_px) img.style.height = "auto";
      p.appendChild(img);
      editor.appendChild(p);
      return;
    }
    renderParagraphBlock(block, editor);
  });
  ensureStarterContent();
  refreshOutline();
  isLoadingDocument = false;
  resetEditorHistory();
  markClean();
}

function applyParagraphStyle(styleId) {
  restoreEditorSelection();
  const style = getStyleById(styleId);
  if (!style) return;
  const blocks = selectedBlockElements();
  if (!blocks.length) {
    setStatus("请先把光标放在要应用样式的段落中。");
    paragraphStyleSelect.value = styleId;
    return;
  }
  recordUndoSnapshot();
  const updatedBlocks = blocks.map((block) => applyStyleVisuals(block, style));
  updatedBlocks[0]?.focus();
  paragraphStyleSelect.value = style.id;
  syncParagraphMetricsControls();
  refreshOutline();
  markDirty();
  setStatus(`已应用样式：${style.name}（${updatedBlocks.length} 段）`);
}

function uniqueStyleId(name) {
  const used = new Set(currentStyles.paragraph.map((style) => style.id));
  const base = String(name || "Custom Style").replace(/[^A-Za-z0-9\u4e00-\u9fa5]+/g, " ").trim() || "Custom Style";
  let raw = base
    .split(/\s+/)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join("")
    .replace(/[^A-Za-z0-9]/g, "");
  if (!raw) raw = "CustomStyle";
  let candidate = raw;
  let index = 2;
  while (used.has(candidate)) {
    candidate = `${raw}${index}`;
    index += 1;
  }
  return candidate;
}

function normalizeShortcutValue(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/control/gi, "Ctrl")
    .replace(/command/gi, "Meta")
    .replace(/cmd/gi, "Meta")
    .replace(/option/gi, "Alt")
    .replace(/shift/gi, "Shift");
}

function loadShortcuts() {
  try {
    const raw = window.localStorage.getItem(SHORTCUT_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    customShortcuts = { ...DEFAULT_SHORTCUTS, ...parsed };
  } catch {
    customShortcuts = { ...DEFAULT_SHORTCUTS };
  }
}

function loadOutlineFilter() {
  try {
    const raw = window.localStorage.getItem(OUTLINE_FILTER_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    if (parsed && typeof parsed === "object") {
      outlineFilter = {
        1: Boolean(parsed[1]),
        2: Boolean(parsed[2]),
        3: Boolean(parsed[3]),
      };
    }
  } catch {
    // Ignore storage errors.
  }
}

function persistOutlineFilter() {
  try {
    window.localStorage.setItem(OUTLINE_FILTER_KEY, JSON.stringify(outlineFilter));
  } catch {
    // Ignore storage errors.
  }
}

function persistShortcuts() {
  window.localStorage.setItem(SHORTCUT_STORAGE_KEY, JSON.stringify(customShortcuts));
}

function renderShortcutInputs() {
  shortcutList.innerHTML = "";
  const shortcutOwners = new Map();
  Object.entries(customShortcuts).forEach(([key, value]) => {
    const normalized = normalizeShortcutValue(value);
    if (!normalized) return;
    if (!shortcutOwners.has(normalized)) shortcutOwners.set(normalized, []);
    shortcutOwners.get(normalized).push(key);
  });
  Object.entries(SHORTCUT_ACTIONS).forEach(([key, action]) => {
    const row = document.createElement("div");
    row.className = "shortcut-item";
    const meta = document.createElement("div");
    meta.className = "shortcut-meta";
    const label = document.createElement("label");
    label.textContent = action.label;
    label.htmlFor = `shortcut-${key}`;
    meta.appendChild(label);
    const normalized = normalizeShortcutValue(customShortcuts[key]);
    const owners = normalized ? shortcutOwners.get(normalized) || [] : [];
    if (owners.length > 1) {
      row.classList.add("has-conflict");
      const conflict = document.createElement("div");
      conflict.className = "shortcut-conflict";
      conflict.textContent = `与 ${owners.filter((item) => item !== key).map((item) => SHORTCUT_ACTIONS[item].label).join("、")} 冲突`;
      meta.appendChild(conflict);
    }
    const input = document.createElement("input");
    input.id = `shortcut-${key}`;
    input.type = "text";
    input.value = customShortcuts[key] || "";
    input.placeholder = DEFAULT_SHORTCUTS[key] || "";
    input.setAttribute("autocomplete", "off");
    input.addEventListener("keydown", (event) => {
      event.preventDefault();
      if (event.key === "Backspace" || event.key === "Delete") {
        customShortcuts[key] = "";
      } else {
        const combo = shortcutFromKeyboardEvent(event);
        if (!combo) return;
        customShortcuts[key] = combo;
      }
      persistShortcuts();
      renderShortcutInputs();
      setStatus(`已更新快捷键：${action.label}`);
    });
    row.append(meta, input);
    shortcutList.appendChild(row);
  });
  const hint = document.createElement("p");
  hint.className = "shortcut-hint";
  hint.textContent = "在输入框按下组合键即可保存，退格键/删除键可清空。";
  shortcutList.appendChild(hint);
}

function shortcutFromKeyboardEvent(event) {
  const parts = [];
  if (event.ctrlKey) parts.push("Ctrl");
  if (event.altKey) parts.push("Alt");
  if (event.shiftKey) parts.push("Shift");
  if (event.metaKey) parts.push("Meta");
  const key = event.key.length === 1 ? event.key.toUpperCase() : event.key;
  if (["Control", "Alt", "Shift", "Meta"].includes(key)) {
    return parts.length ? parts.join("+") : "";
  }
  parts.push(key === " " ? "Space" : key);
  return parts.join("+");
}

function matchesShortcut(event, shortcut) {
  const expected = normalizeShortcutValue(shortcut);
  if (!expected) return false;
  return normalizeShortcutValue(shortcutFromKeyboardEvent(event)) === expected;
}

function nativeUndoRedoAction(event) {
  if (!(event.ctrlKey || event.metaKey)) return "";
  if (event.altKey) return "";
  const key = String(event.key || "").toLowerCase();
  if (key === "z" && !event.shiftKey) return "undo";
  if (key === "y" && !event.shiftKey) return "redo";
  if (key === "z" && event.shiftKey) return "redo";
  return "";
}

function isTextInputOutsideEditor(target) {
  if (!target || !target.closest) return false;
  if (editor.contains(target) || target === editor) return false;
  if (target.closest(".shortcut-item")) return true;
  if (target.closest("textarea, select")) return true;
  const input = target.closest("input");
  if (input) {
    const type = String(input.type || "").toLowerCase();
    return !["button", "checkbox", "radio", "range", "color", "file"].includes(type);
  }
  const contentEditable = target.closest("[contenteditable='true']");
  return Boolean(contentEditable && contentEditable !== editor);
}

function saveCurrentStyle() {
  updateStyleFromSelection();
}

function updateStyleFromSelection() {
  restoreEditorSelection();
  const block = currentBlockElement();
  if (!block) {
    setStatus("请先把光标放在段落中，再更新样式。");
    return;
  }
  const styleId = paragraphStyleSelect.value;
  const style = getStyleById(styleId);
  if (!style) {
    setStatus("未找到该样式");
    return;
  }
  const previousDescriptor = cloneDescriptor(style.descriptor);
  const metrics = paragraphMetricsFromElement(block);
  const nextDescriptor = selectionDescriptorOrBlockDescriptor(block);
  if (fontFamily?.value) {
    nextDescriptor[0] = fontFamilyForDocument(fontFamily.value);
  }
  if (fontSize?.value) {
    nextDescriptor[1] = Math.max(Number(fontSize.value) || nextDescriptor[1], 1);
  }
  style.descriptor = nextDescriptor;
  style.alignment = { align_left: "left", align_center: "center", align_right: "right", align_justify: "justify" }[blockAlignment(block)] || "left";
  style.line_spacing = Number(metrics.lineSpacing) || style.line_spacing || DEFAULT_LINE_SPACING;
  style.space_before = Number(metrics.spaceBefore) || DEFAULT_PARAGRAPH_SPACING;
  style.space_after = Number(metrics.spaceAfter) || DEFAULT_PARAGRAPH_SPACING;

  const targets = Array.from(editor.querySelectorAll(`[data-style-id="${style.id}"]`));
  targets.forEach((target) => {
    applyStyleVisuals(target, style);
    syncStyledRunsForStyleUpdate(target, previousDescriptor, style.descriptor);
  });
  stylesDirty = true;
  syncParagraphMetricsControls();
  refreshOutline();
  markDirty();
  setStatus(`已用当前段落更新样式：${style.name}（${targets.length} 段）`);
}

async function openDocx(file) {
  setStatus(`正在打开 ${file.name}...`);
  const dataUrl = await toBase64(file);
  const response = await fetch("/api/import-docx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: file.name, data: dataUrl.split(",", 2)[1] }),
  });
  const result = await response.json();
  if (!response.ok) {
    throw new Error(result.error || "打开文件失败。");
  }
  loadDocument(result.document);
  currentFileName = file.name || currentFileName;
  updateFileUiState();
  setStatus(`已打开 ${file.name}`);
}

async function requestDocxBlob(filename) {
  const response = await fetch("/api/export-docx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filename, document: editorToDocument() }),
  });
  if (!response.ok) {
    const result = await response.json();
    throw new Error(result.error || "保存文件失败。");
  }
  return response.blob();
}

async function writeBlobToHandle(fileHandle, blob) {
  const writable = await fileHandle.createWritable();
  await writable.write(blob);
  await writable.close();
}

async function saveDocx(options = {}) {
  const interactive = options.interactive !== false;
  const allowPicker = options.allowPicker === true;
  setStatus("正在保存文件...");
  const targetName = currentFileName || "mini-docx.docx";
  try {
    if (!currentFileHandle) {
      if (supportsFileSystemAccess()) {
        if (!interactive && !allowPicker) {
          setStatus("请先点击“保存文件”选择保存位置。");
          return;
        }
        currentFileHandle = await window.showSaveFilePicker({
          suggestedName: targetName,
          types: [{ description: "Word 文档", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }],
        });
        currentFileName = currentFileHandle.name || targetName;
      } else {
        if (!interactive && !allowPicker) {
          setStatus("当前环境无法静默保存，请点击“保存文件”。");
          return;
        }
        return exportDocx();
      }
    }
    const hasPermission = await ensureHandlePermission(currentFileHandle, true, interactive || allowPicker);
    if (!hasPermission) {
      if (!interactive && allowPicker && supportsFileSystemAccess()) {
        currentFileHandle = null;
        return saveDocx({ interactive: true, allowPicker: true });
      }
      setStatus("保存需要权限，请点击“保存文件”。");
      return;
    }
    const blob = await requestDocxBlob(currentFileName);
    await writeBlobToHandle(currentFileHandle, blob);
    markClean();
    setStatus(`已保存：${currentFileName}`);
  } catch (error) {
    if (isStaleHandleError(error)) {
      currentFileHandle = null;
      if (interactive) {
        return saveDocx({ interactive });
      }
      setStatus("保存失败：文件句柄已失效，请重新点击“保存文件”选择位置。");
      return;
    }
    if (!interactive) {
      setStatus(error.message || String(error));
      return;
    }
    throw error;
  }
}

async function exportDocx() {
  const targetName = currentFileName || "mini-docx.docx";
  setStatus("正在导出 DOCX...");
  const blob = await requestDocxBlob(targetName);
  if (supportsFileSystemAccess()) {
    const exportHandle = await window.showSaveFilePicker({
      suggestedName: targetName,
      types: [{ description: "Word 文档", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }],
    });
    await writeBlobToHandle(exportHandle, blob);
    setStatus(`已导出：${exportHandle.name || targetName}`);
    return;
  }
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = targetName.endsWith(".docx") ? targetName : `${targetName}.docx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setStatus(`已导出：${link.download}`);
}

document.getElementById("newBtn").addEventListener("click", () => {
  if (!confirmDiscardChanges()) return;
  docxMeta = null;
  stylesDirty = false;
  numberingDirty = false;
  currentStyles = normalizeStyles(defaultStyles());
  currentFileHandle = null;
  currentFileName = "mini-docx.docx";
  populateStyleSelect("Normal");
  editor.innerHTML = "";
  ensureStarterContent();
  resetEditorHistory();
  markClean();
  setStatus("已新建文档");
});

openBtn.addEventListener("click", async () => {
  try {
    await openDocxPicker();
  } catch (error) {
    handleAsyncError(error);
  }
});

saveBtn.addEventListener("click", async () => {
  try {
    await saveDocx();
  } catch (error) {
    handleAsyncError(error);
  }
});

exportBtn?.addEventListener("click", async () => {
  try {
    await exportDocx();
  } catch (error) {
    handleAsyncError(error);
  }
});

document.getElementById("insertTableBtn").addEventListener("click", insertTable);
document.getElementById("addRowBtn").addEventListener("click", addTableRow);
document.getElementById("removeRowBtn").addEventListener("click", removeTableRow);
document.getElementById("addColBtn").addEventListener("click", addTableColumn);
document.getElementById("removeColBtn").addEventListener("click", removeTableColumn);
deleteTableBtn.addEventListener("click", deleteCurrentTable);

document.getElementById("refreshOutlineBtn")?.addEventListener("click", refreshOutline);
if (outlineLevel1 && outlineLevel2 && outlineLevel3) {
  outlineLevel1.checked = outlineFilter[1];
  outlineLevel2.checked = outlineFilter[2];
  outlineLevel3.checked = outlineFilter[3];
  const syncOutlineFilter = () => {
    outlineFilter = {
      1: outlineLevel1.checked,
      2: outlineLevel2.checked,
      3: outlineLevel3.checked,
    };
    persistOutlineFilter();
    refreshOutline();
  };
  [outlineLevel1, outlineLevel2, outlineLevel3].forEach((input) => {
    input.addEventListener("change", syncOutlineFilter);
  });
}

Array.from(document.querySelectorAll("[data-cmd]")).forEach((button) => {
  button.addEventListener("click", () => exec(button.dataset.cmd));
});

Array.from(document.querySelectorAll("[data-align]")).forEach((button) => {
  button.addEventListener("click", () => applyParagraphAlignment(button.dataset.align));
});

Array.from(document.querySelectorAll(".toolbar-wrap button")).forEach((button) => {
  button.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });
});

[fontFamily, fontSize, paragraphStyleSelect, lineSpacingSelect, numberFormatSelect].forEach((control) => {
  if (!control) return;
  control.addEventListener("pointerdown", captureEditorSelection);
  control.addEventListener("mousedown", captureEditorSelection);
});

fontFamily.addEventListener("change", () => {
  debugLog("fontFamily:change", summarizeCurrentSelectionForDebug());
  const selectedFamily = fontFamilyForDocument(fontFamily.value);
  const selection = window.getSelection();
  const collapsed = !selection || selection.rangeCount === 0 || selection.getRangeAt(0).collapsed;
  if (collapsed) {
    restoreEditorSelection();
    const blocks = selectedBlockElements();
    if (!blocks.length) {
      setStatus("请先把光标放在正文中。");
      return;
    }
    recordUndoSnapshot();
    blocks.forEach((block) => applyFontFamilyToBlock(block, selectedFamily));
    captureEditorSelection();
    syncSelectionUi();
    markDirty();
    refreshOutline();
    debugLog("fontFamily:block-apply", {
      ...summarizeCurrentSelectionForDebug(),
      appliedToBlocks: blocks.length,
      family: selectedFamily,
    });
    return;
  }
  exec("fontName", selectedFamily);
});

fontSize.addEventListener("change", () => {
  debugLog("fontSize:change", summarizeCurrentSelectionForDebug());
  const targetPointSize = Number(fontSize.value);
  if (!Number.isFinite(targetPointSize) || targetPointSize <= 0) return;
  const selection = window.getSelection();
  const collapsed = !selection || selection.rangeCount === 0 || selection.getRangeAt(0).collapsed;
  if (collapsed) {
    restoreEditorSelection();
    const blocks = selectedBlockElements();
    if (!blocks.length) {
      setStatus("请先选中文字或把光标放在正文中。");
      return;
    }
    recordUndoSnapshot();
    blocks.forEach((block) => applyFontSizeToBlock(block, targetPointSize));
    captureEditorSelection();
    syncSelectionUi();
    markDirty();
    refreshOutline();
    setStatus(`已应用字号 ${targetPointSize}pt`);
    return;
  }
  if (!selectionInsideEditor() && !restoreEditorSelection()) {
    setStatus("请先选中文字或把光标放在正文中。");
    return;
  }

  const targetPx = `${targetPointSize / 0.75}px`;
  recordUndoSnapshot();
  editor.focus();

  const snapshot = serializeEditorSelection();
  const existingLegacySize7 = new Set(editor.querySelectorAll("font[size='7']"));
  const existingFontSizeSpans = new Set(editor.querySelectorAll("span[style*='font-size']"));

  document.execCommand("styleWithCSS", false, false);
  if (snapshot) {
    restoreSerializedEditorSelection(snapshot);
  }
  document.execCommand("fontSize", false, "7");

  let touched = false;
  Array.from(editor.querySelectorAll("font[size='7']")).forEach((fontTag) => {
    if (existingLegacySize7.has(fontTag)) return;
    fontTag.removeAttribute("size");
    fontTag.style.fontSize = targetPx;
    touched = true;
  });

  if (!touched) {
    if (snapshot) {
      restoreSerializedEditorSelection(snapshot);
    }
    document.execCommand("styleWithCSS", false, true);
    if (snapshot) {
      restoreSerializedEditorSelection(snapshot);
    }
    document.execCommand("fontSize", false, "7");
    Array.from(editor.querySelectorAll("span[style*='font-size']")).forEach((span) => {
      if (existingFontSizeSpans.has(span)) return;
      span.style.fontSize = targetPx;
      touched = true;
    });
  }

  markDirty();
  captureEditorSelection();
  refreshOutline();
  if (!touched) {
    setStatus("未检测到可应用字号的选区，请重新选择文字后再试。");
    return;
  }
  setStatus(`已应用字号 ${targetPointSize}pt`);
});

applyHighlightBtn.addEventListener("click", () => applyTextBackground(highlightColor.value));

if (applyPageSizeBtn && pageWidthInput && pageHeightInput && pageSizeSelect) {
  pageSizeSelect.addEventListener("change", () => {
    if (pageSizeSelect.value === "A4") {
      setPageSizeFromMm(210, 297);
      return;
    }
    if (pageSizeSelect.value === "Letter") {
      setPageSizeFromMm(216, 279);
      return;
    }
  });
  applyPageSizeBtn.addEventListener("click", () => {
    setPageSizeFromMm(pageWidthInput.value, pageHeightInput.value);
  });
}

paragraphStyleSelect.addEventListener("change", () => applyParagraphStyle(paragraphStyleSelect.value));
lineSpacingSelect.addEventListener("change", applyCurrentParagraphMetrics);
spaceBeforeInput.addEventListener("change", applyCurrentParagraphMetrics);
spaceAfterInput.addEventListener("change", applyCurrentParagraphMetrics);
toggleNumberingBtn?.addEventListener("click", toggleNumbering);
numberLevelUpBtn?.addEventListener("click", () => updateNumberingLevel(1));
numberLevelDownBtn?.addEventListener("click", () => updateNumberingLevel(-1));
numberFormatSelect?.addEventListener("change", applyNumberFormatToSelection);
resetParagraphSpacingBtn.addEventListener("click", () => {
  lineSpacingSelect.value = String(DEFAULT_LINE_SPACING);
  spaceBeforeInput.value = "0";
  spaceAfterInput.value = "0";
  applyCurrentParagraphMetrics();
});
clearFormatBtn.addEventListener("click", () => {
  exec("removeFormat");
  setStatus("已清除选区文字格式。");
});
saveStyleBtn.addEventListener("click", saveCurrentStyle);
updateStyleBtn.addEventListener("click", updateStyleFromSelection);
formatPainterBtn.addEventListener("click", () => {
  if (formatPainterPayload) {
    clearFormatPainter();
    setStatus("已关闭格式刷。");
    return;
  }
  captureFormatPainter();
});
shortcutSettingsBtn.addEventListener("click", openShortcutModal);
closeShortcutModalBtn.addEventListener("click", closeShortcutModal);
shortcutModal.addEventListener("click", (event) => {
  if (event.target.dataset.closeModal === "true") {
    closeShortcutModal();
  }
});
historyPanelBtn.addEventListener("click", openHistoryModal);
closeHistoryModalBtn.addEventListener("click", closeHistoryModal);
historyModal.addEventListener("click", (event) => {
  if (event.target.dataset.closeModal === "true") {
    closeHistoryModal();
  }
});
browseDirBtn.addEventListener("click", async () => {
  if (!supportsDirectoryAccess()) {
    setStatus("当前浏览器不支持目录选择。");
    return;
  }
  try {
    const handle = await window.showDirectoryPicker();
    const files = await listDocxInDirectory(handle);
    renderDirFiles(files, handle.name || "目录");
  } catch (error) {
    handleAsyncError(error);
  }
});
addFavoriteDirBtn.addEventListener("click", async () => {
  if (!supportsDirectoryAccess()) {
    setStatus("当前浏览器不支持目录选择。");
    return;
  }
  try {
    const handle = await window.showDirectoryPicker();
    await addFavoriteDir(handle);
    renderFavoriteList();
  } catch (error) {
    handleAsyncError(error);
  }
});
resetShortcutsBtn.addEventListener("click", () => {
  customShortcuts = { ...DEFAULT_SHORTCUTS };
  persistShortcuts();
  renderShortcutInputs();
  setStatus("已重置快捷键。");
});

openDocxInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  try {
    if (!confirmDiscardChanges()) return;
    currentFileHandle = null;
    await openDocx(file);
    updateFileUiState();
  } catch (error) {
    handleAsyncError(error);
  } finally {
    event.target.value = "";
  }
});

imageInput.addEventListener("change", async (event) => {
  const files = Array.from(event.target.files || []);
  for (const file of files) {
    const dataUrl = await toBase64(file);
    insertImage(dataUrl, file.name);
  }
  if (files.length) markDirty();
  setStatus(`已插入 ${files.length} 张图片`);
  event.target.value = "";
});

editor.addEventListener("input", refreshOutline);
editor.addEventListener("input", markDirty);
editor.addEventListener("beforeinput", (event) => {
  if (suppressEditorHistory || isLoadingDocument) return;
  if (event.inputType === "historyUndo" || event.inputType === "historyRedo") return;
  recordUndoSnapshot();
});
editor.addEventListener("keyup", refreshOutline);
editor.addEventListener("click", syncParagraphStyleSelect);
editor.addEventListener("click", (event) => {
  const block = event.target.closest && event.target.closest("p, h1, h2, h3, div");
  if (formatPainterPayload && block) {
    applyFormatPainterToBlock(block);
  }
});
editor.addEventListener("paste", async (event) => {
  const clipboard = event.clipboardData;
  if (!clipboard) {
    window.setTimeout(refreshOutline, 20);
    return;
  }
  const items = Array.from(clipboard.items || []);
  const imageItems = items.filter((item) => item.type && item.type.startsWith("image/"));
  if (!imageItems.length) {
    window.setTimeout(refreshOutline, 20);
    return;
  }
  event.preventDefault();
  for (const item of imageItems) {
    const file = item.getAsFile();
    if (!file) continue;
    const dataUrl = await toBase64(file);
    insertImage(dataUrl, file.name || "pasted-image.png");
  }
  markDirty();
  refreshOutline();
});
editor.addEventListener("keydown", (event) => {
  if (event.key === "Tab") {
    const targets = selectedBlockElements().filter((block) => Boolean(numberingFromElement(block)));
    if (targets.length) {
      event.preventDefault();
      updateNumberingLevel(event.shiftKey ? -1 : 1);
    }
    return;
  }
  if (event.key !== "Enter" || event.shiftKey) return;
  const current = currentBlockElement();
  const currentNumbering = numberingFromElement(current);
  const hasText = (current?.textContent || "").trim().length > 0;
  if (!current || !currentNumbering || !hasText) return;
  window.setTimeout(() => {
    const nextBlock = currentBlockElement();
    if (!nextBlock || nextBlock === current || numberingFromElement(nextBlock)) {
      return;
    }
    setNumberingData(nextBlock, currentNumbering);
    refreshOutline();
    markDirty();
  }, 0);
});
document.addEventListener("selectionchange", () => {
  if (document.activeElement === editor || editor.contains(document.activeElement) || editor.contains(window.getSelection()?.anchorNode)) {
    debugLog("selectionchange", summarizeCurrentSelectionForDebug());
    window.setTimeout(syncSelectionUi, 0);
  }
});

editor.addEventListener("mouseup", () => {
  debugLog("editor:mouseup", summarizeCurrentSelectionForDebug());
  window.setTimeout(syncSelectionUi, 0);
});

editor.addEventListener("keyup", () => {
  debugLog("editor:keyup", summarizeCurrentSelectionForDebug());
  window.setTimeout(syncSelectionUi, 0);
});

editor.addEventListener("focus", () => {
  debugLog("editor:focus", summarizeCurrentSelectionForDebug());
  window.setTimeout(syncSelectionUi, 0);
});

document.addEventListener("keydown", (event) => {
  const target = event.target;
  const editingShortcut = target && target.closest && target.closest(".shortcut-item");
  if (editingShortcut) return;

  if (event.ctrlKey || event.metaKey) {
    if (["=", "+"].includes(event.key)) {
      event.preventDefault();
      adjustEditorZoom(0.1);
      return;
    }
    if (["-", "_"].includes(event.key)) {
      event.preventDefault();
      adjustEditorZoom(-0.1);
      return;
    }
    if (event.key === "0") {
      event.preventDefault();
      resetEditorZoom();
      return;
    }
  }

  const activeInsideEditor = document.activeElement === editor || editor.contains(document.activeElement);
  const editorContext = activeInsideEditor || selectionInsideEditor();
  const nativeUndoRedo = nativeUndoRedoAction(event);
  const textInputOutsideEditor = isTextInputOutsideEditor(target);
  const allowEditorUndoRedo = editorContext || (!textInputOutsideEditor && (editorUndoStack.length > 0 || editorRedoStack.length > 0));
  if (nativeUndoRedo && allowEditorUndoRedo) {
    event.preventDefault();
    if (nativeUndoRedo === "undo") {
      if (!performEditorUndo()) {
        exec("undo");
      }
    } else if (!performEditorRedo()) {
      exec("redo");
    }
    return;
  }

  for (const [key, action] of Object.entries(SHORTCUT_ACTIONS)) {
    if (editorContext && (key === "undo" || key === "redo")) {
      continue;
    }
    if (matchesShortcut(event, customShortcuts[key])) {
      event.preventDefault();
      Promise.resolve(action.run()).catch((error) => {
        handleAsyncError(error);
      });
      return;
    }
  }
});

pageStage.addEventListener("wheel", (event) => {
  if (!(event.ctrlKey || event.metaKey)) return;
  event.preventDefault();
  adjustEditorZoom(event.deltaY < 0 ? 0.1 : -0.1);
}, { passive: false });

window.addEventListener("beforeunload", (event) => {
  if (!isDirty) return;
  event.preventDefault();
  event.returnValue = "";
});

currentStyles = normalizeStyles(defaultStyles());
loadOutlineFilter();
loadShortcuts();
initLayoutToggles();
if (saveStyleBtn) {
  saveStyleBtn.title = "保存到当前 H1、H2、H3 或 Normal 样式";
}
populateStyleSelect("Normal");
renderShortcutInputs();
setServerAddress();
updateFileUiState();
applyEditorZoom();
ensureStarterContent();
applyPageSizeToEditor();
syncPageSizeControls();
resetEditorHistory();
markClean();

