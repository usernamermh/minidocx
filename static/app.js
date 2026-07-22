
// 新增一个样式：
// 1. 在 static/app.js 的 ALLOWED_STYLE_ORDER 里加 ID
// 2. 在 defaultStyles() 里加一条样式定义
// 3. 在 docx_io.py 的 ALLOWED_STYLE_ORDER 里加同样的 ID
// 4. 在 _builtin_styles() 里加对应定义
// 5. 如果需要兼容导入名称，再补 STYLE_ALIAS_MAP

const editor = document.getElementById("editor");
const appShell = document.getElementById("appShell");
const leftSidebar = document.getElementById("leftSidebar");
const topToolbar = document.getElementById("topToolbar");
const pageStage = document.getElementById("pageStage");
const chapterFoldOverlay = document.getElementById("chapterFoldOverlay");
const outline = document.getElementById("outline");
const secondaryOutline = document.getElementById("secondaryOutline");
const secondaryOutlineLevelText = document.getElementById("secondaryOutlineLevelText");
const primaryOutlineMinSlider = document.getElementById("primaryOutlineMinSlider");
const primaryOutlineMaxSlider = document.getElementById("primaryOutlineMaxSlider");
const secondaryOutlineMaxSlider = document.getElementById("secondaryOutlineMaxSlider");
const primaryOutlineMinText = document.getElementById("primaryOutlineMinText");
const primaryOutlineMaxText = document.getElementById("primaryOutlineMaxText");
const secondaryOutlineMaxText = document.getElementById("secondaryOutlineMaxText");
const primaryOutlineLevelToggles = document.getElementById("primaryOutlineLevelToggles");
const statusText = document.getElementById("statusText");
const operationStatusText = document.getElementById("operationStatusText");
const serverAddress = document.getElementById("serverAddress");
const openBtn = document.getElementById("openBtn");
const openDocxInput = document.getElementById("openDocxInput");
const exportBtn = document.getElementById("exportBtn");
const fontFamily = document.getElementById("fontFamily");
const fontSize = document.getElementById("fontSize");
const highlightColor = document.getElementById("highlightColor");
const applyHighlightBtn = document.getElementById("applyHighlightBtn");
const pageWidthInput = document.getElementById("pageWidthInput");
const pageHeightInput = document.getElementById("pageHeightInput");
const applyPageSizeBtn = document.getElementById("applyPageSizeBtn");
const tableAdvancedToggle = document.getElementById("tableAdvancedToggle");
const tableAdvancedContent = document.getElementById("tableAdvancedContent");
const pageAdvancedToggle = document.getElementById("pageAdvancedToggle");
const pageAdvancedContent = document.getElementById("pageAdvancedContent");
const highlightAdvancedToggle = document.getElementById("highlightAdvancedToggle");
const highlightAdvancedContent = document.getElementById("highlightAdvancedContent");
const paragraphStyleSelect = document.getElementById("paragraphStyleSelect");
const lineSpacingSelect = document.getElementById("lineSpacingSelect");
const spaceBeforeInput = document.getElementById("spaceBeforeInput");
const spaceAfterInput = document.getElementById("spaceAfterInput");
const saveStyleBtn = document.getElementById("saveStyleBtn");
const updateStyleBtn = document.getElementById("updateStyleBtn");
const formatPainterBtn = document.getElementById("formatPainterBtn");
const clearFormatBtn = document.getElementById("clearFormatBtn");
const toggleNumberingBtn = document.getElementById("toggleNumberingBtn");
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
const resourceCpuText = document.getElementById("resourceCpuText");
const resourceMemoryText = document.getElementById("resourceMemoryText");
const resourceMemoryDetail = document.getElementById("resourceMemoryDetail");
const resourceStatusText = document.getElementById("resourceStatusText");
const cleanResourcesBtn = document.getElementById("cleanResourcesBtn");
const collapseAllChaptersBtn = document.getElementById("collapseAllChaptersBtn");
const expandAllChaptersBtn = document.getElementById("expandAllChaptersBtn");
const findReplaceModal = document.getElementById("findReplaceModal");
const closeFindReplaceBtn = document.getElementById("closeFindReplaceBtn");
const findTextInput = document.getElementById("findTextInput");
const replaceTextInput = document.getElementById("replaceTextInput");
const findCaseSensitive = document.getElementById("findCaseSensitive");
const findMatchStatus = document.getElementById("findMatchStatus");
const findPreviousBtn = document.getElementById("findPreviousBtn");
const findNextBtn = document.getElementById("findNextBtn");
const replaceCurrentBtn = document.getElementById("replaceCurrentBtn");
const replaceAllBtn = document.getElementById("replaceAllBtn");

let currentStyles = { paragraph: [] };
let customShortcuts = {};
let currentFileHandle = null;
let currentFileName = "mini-docx.docx";
let currentFilePath = null;
let formatPainterPayload = null;
let isDirty = false;
let isLoadingDocument = false;
let editorZoom = 1;
let savedSelectionRange = null;
let lastClickedParagraph = null;
let outlineFilter = { 0: true, 1: true, 2: true, 3: true };
let showOtherOutlineBranches = true;
let outlineConfig = { primaryMin: 0, primaryMax: 3, secondaryMax: 5 };
let pageSize = { widthTwips: 11906, heightTwips: 16838 };
let docxMeta = null;
let stylesDirty = false;
let numberingDirty = false;
let pageDirty = false;
let cleanEditorHtml = "";
let editorUndoStack = [];
let editorRedoStack = [];
let suppressEditorHistory = false;
let debugLogSequence = 0;
let inputHistorySnapshotRecorded = false;
let inputHistoryResetTimer = null;
let editorRefreshTimer = null;
let activePrimaryOutlineElement = null;
let activePrimaryOutlineBlockIndex = null;
let primaryOutlineWasManuallySelected = false;
let activeSecondaryOutlineElement = null;
let activeSecondaryOutlineBlockIndex = null;
let secondaryOutlineWasManuallySelected = false;
let outlineDragItem = null;
let chapterFoldOverlayFrame = null;
let pendingFoldOverlayHeadings = null;

const SHORTCUT_STORAGE_KEY = "mini_docx_shortcuts";
const OUTLINE_FILTER_KEY = "mini_docx_outline_filter";
const OUTLINE_CONFIG_KEY = "mini_docx_outline_config_v3";
const LEGACY_OUTLINE_CONFIG_KEY = "mini_docx_outline_config_v2";
const LAYOUT_STORAGE_KEY = "mini_docx_layout";
const LAYOUT_STRUCTURE_VERSION = 9;
const HANDLE_DB_NAME = "mini_docx_handles";
const HANDLE_DB_VERSION = 1;
const STORE_RECENT = "recent_files";
const STORE_FAVORITE = "favorite_dirs";
const MAX_RECENT_FILES = 12;
const NUMBER_FORMATS = new Set(["decimal", "upperLetter", "lowerLetter", "upperRoman", "lowerRoman"]);
const DEFAULT_NUMBER_FORMAT = "decimal";
const MAX_NUMBERING_LEVEL = 8;
const MAX_EDITOR_HISTORY = 80;
const MAX_EDITOR_HISTORY_BYTES = 8 * 1024 * 1024;
const INPUT_HISTORY_DEBOUNCE_MS = 750;
const EDITOR_REFRESH_DEBOUNCE_MS = 180;
const NUMBERING_LEVEL_INDENT_PX = 24;
const MIN_NUMBERING_PREFIX_PX = 18;
const DEFAULT_FONT_FAMILY = '"Times New Roman", SimSun';
const DEFAULT_LINE_SPACING = 1.5;
const DEFAULT_DOCUMENT_WIDTH_MM = 210;
const DEFAULT_DOCUMENT_HEIGHT_MM = 297;
const DEFAULT_PARAGRAPH_SPACING = 1;
const BLOCK_INDENT_STEP_EM = 2;
const MAX_BLOCK_INDENT_LEVEL = 20;
const MIN_SIDEBAR_WIDTH = 220;
const MAX_SIDEBAR_WIDTH = 520;
const SIDEBAR_WIDTH_STEP = 16;
const ALLOWED_STYLE_ORDER = ["Normal", "Heading1", "Heading2", "Heading3", "NormalL1", "NormalL2", "NormalL3"];
const RESOURCE_REFRESH_MS = 5000;
const DEBUG_LOG_ENABLED = false;
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
  if (!DEBUG_LOG_ENABLED) return;
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
  topToolbar?.parentElement?.classList.toggle("is-toolbar-collapsed", collapsed);
  if (toolbarToggleBtn) {
    toolbarToggleBtn.textContent = collapsed ? "显示上方" : "隐藏上方";
    toolbarToggleBtn.setAttribute("aria-expanded", String(!collapsed));
  }
}

function clampLayoutSize(value, minimum, maximum) {
  return Math.min(Math.max(Math.round(Number(value) || minimum), minimum), maximum);
}

function setSidebarWidth(width) {
  const nextWidth = clampLayoutSize(width, MIN_SIDEBAR_WIDTH, MAX_SIDEBAR_WIDTH);
  appShell?.style.setProperty("--sidebar-width", `${nextWidth}px`);
  requestAnimationFrame(constrainExpandedToolPanels);
  return nextWidth;
}

function resizeDirectionFromWheel(event) {
  if (event.deltaY === 0) return 0;
  return event.deltaY < 0 ? 1 : -1;
}

function initLayoutToggles() {
  const state = loadLayoutState();
  const initialSidebarWidth = setSidebarWidth(state.sidebarWidth || leftSidebar?.getBoundingClientRect().width || 300);
  topToolbar?.style.removeProperty("height");
  setSidebarCollapsed(Boolean(state.sidebarCollapsed));
  setToolbarCollapsed(Boolean(state.toolbarCollapsed));

  const initialState = {
    ...state,
    sidebarWidth: initialSidebarWidth,
    layoutVersion: LAYOUT_STRUCTURE_VERSION,
  };
  saveLayoutState(initialState);

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

  leftSidebar?.querySelector(".sidebar-toggle-strip")?.addEventListener("wheel", (event) => {
    if (appShell?.classList.contains("is-sidebar-collapsed")) return;
    const direction = resizeDirectionFromWheel(event);
    if (!direction) return;
    event.preventDefault();
    const currentWidth = leftSidebar.getBoundingClientRect().width;
    const sidebarWidth = setSidebarWidth(currentWidth + direction * SIDEBAR_WIDTH_STEP);
    saveLayoutState({ ...loadLayoutState(), sidebarWidth });
  }, { passive: false });

}

function constrainExpandedToolPanel(content) {
  if (!content || content.classList.contains("hidden-tool-content") || !topToolbar) return;
  const toolbarRect = topToolbar.getBoundingClientRect();
  const inset = 10;
  const availableWidth = Math.max(0, toolbarRect.width - inset * 2);
  content.style.maxWidth = `${availableWidth}px`;
  content.style.transform = "translateX(-50%)";

  const panelRect = content.getBoundingClientRect();
  const leftLimit = toolbarRect.left + inset;
  const rightLimit = toolbarRect.right - inset;
  let shiftX = 0;
  if (panelRect.left < leftLimit) shiftX += leftLimit - panelRect.left;
  if (panelRect.right + shiftX > rightLimit) shiftX -= panelRect.right + shiftX - rightLimit;
  content.style.transform = `translateX(calc(-50% + ${Math.round(shiftX)}px))`;
}

function constrainExpandedToolPanels() {
  [highlightAdvancedContent, tableAdvancedContent, pageAdvancedContent].forEach(constrainExpandedToolPanel);
}

function syncToolbarHeightToExpandedPanel() {
  if (!topToolbar || topToolbar.classList.contains("is-collapsed")) return;
  topToolbar.style.removeProperty("height");
  const expandedContents = [highlightAdvancedContent, tableAdvancedContent, pageAdvancedContent]
    .filter((content) => content && !content.classList.contains("hidden-tool-content"));
  if (!expandedContents.length) {
    topToolbar.classList.remove("has-expanded-tool");
    return;
  }
  topToolbar.classList.add("has-expanded-tool");
  requestAnimationFrame(() => {
    const toolbarRect = topToolbar.getBoundingClientRect();
    const requiredHeight = expandedContents.reduce((height, content) => {
      const contentRect = content.getBoundingClientRect();
      return Math.max(height, contentRect.bottom - toolbarRect.top + 12);
    }, topToolbar.scrollHeight);
    topToolbar.style.height = `${Math.ceil(requiredHeight)}px`;
  });
}

function setAdvancedToolGroupExpanded(toggle, content, expanded) {
  if (!toggle || !content) return;
  content.classList.toggle("hidden-tool-content", !expanded);
  toggle.setAttribute("aria-expanded", String(expanded));
  toggle.textContent = toggle.dataset.label || toggle.textContent;
  toggle.title = expanded ? `收起${toggle.textContent}` : `展开${toggle.textContent}`;
  requestAnimationFrame(() => {
    if (expanded) constrainExpandedToolPanel(content);
    syncToolbarHeightToExpandedPanel();
  });
}

function initAdvancedToolGroups() {
  const groups = [
    [highlightAdvancedToggle, highlightAdvancedContent],
    [tableAdvancedToggle, tableAdvancedContent],
    [pageAdvancedToggle, pageAdvancedContent],
  ];
  groups.forEach(([toggle, content]) => {
    if (!toggle || !content) return;
    setAdvancedToolGroupExpanded(toggle, content, false);
    toggle.addEventListener("click", () => {
      const expanded = toggle.getAttribute("aria-expanded") !== "true";
      if (expanded) {
        groups.forEach(([otherToggle, otherContent]) => {
          if (otherToggle !== toggle) setAdvancedToolGroupExpanded(otherToggle, otherContent, false);
        });
      }
      setAdvancedToolGroupExpanded(toggle, content, expanded);
    });
  });
  window.addEventListener("resize", () => requestAnimationFrame(constrainExpandedToolPanels));
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
  if (operationStatusText) {
    operationStatusText.textContent = text;
    operationStatusText.title = text;
  }
}

function windowsPath(path) {
  return String(path || "").replaceAll("/", "\\");
}

function updateCurrentFileStatus() {
  if (!statusText) return;
  const displayPath = currentFilePath ? windowsPath(currentFilePath) : "尚未保存";
  statusText.textContent = `当前文件：${displayPath}`;
  statusText.title = currentFilePath ? displayPath : "当前文件尚未选择保存位置";
}

function formatPercent(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "--%";
  return `${number.toFixed(number % 1 === 0 ? 0 : 1)}%`;
}

function formatMemoryMb(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "--";
  if (number >= 1024) return `${(number / 1024).toFixed(1)} GB`;
  return `${Math.round(number)} MB`;
}

function renderResourceStats(data) {
  if (!data || !data.ok) return;
  if (resourceCpuText) resourceCpuText.textContent = formatPercent(data.cpu_percent);
  if (resourceMemoryText) resourceMemoryText.textContent = formatPercent(data.memory?.percent);
  if (resourceMemoryDetail && data.memory) {
    resourceMemoryDetail.textContent = `${formatMemoryMb(data.memory.used_mb)} / ${formatMemoryMb(data.memory.total_mb)}`;
  }
}

async function refreshResourceStats() {
  if (!resourceCpuText && !resourceMemoryText) return;
  try {
    const response = await fetch("/api/resource-stats", { cache: "no-store" });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || "资源状态获取失败");
    }
    renderResourceStats(data);
    if (resourceStatusText) resourceStatusText.textContent = "";
  } catch (error) {
    if (resourceStatusText) resourceStatusText.textContent = error.message || "资源状态获取失败";
  }
}

async function cleanResources() {
  if (!cleanResourcesBtn) return;
  cleanResourcesBtn.disabled = true;
  if (resourceStatusText) resourceStatusText.textContent = "正在清理资源...";
  try {
    const response = await fetch("/api/clean-resources", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}",
    });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || "资源清理失败");
    }
    renderResourceStats(data.stats);
    const adminHint = data.is_admin ? "" : "，非管理员权限效果有限";
    if (resourceStatusText) {
      resourceStatusText.textContent = `已清理 ${data.cleaned_count} 个进程${adminHint}`;
    }
    setStatus(`已清理 ${data.cleaned_count} 个进程的工作集`);
  } catch (error) {
    if (resourceStatusText) resourceStatusText.textContent = error.message || "资源清理失败";
    setStatus(error.message || "资源清理失败");
  } finally {
    cleanResourcesBtn.disabled = false;
  }
}

function cloneDocxMeta(meta) {
  if (!meta || typeof meta !== "object") return null;
  const next = {};
  if (typeof meta.source_token === "string" && meta.source_token) {
    next.source_token = meta.source_token;
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
  const snapshot = {
    html: editor.innerHTML,
    selection: serializeEditorSelection(),
  };
  snapshot.bytes = new Blob([snapshot.html]).size;
  return snapshot;
}

function trimEditorHistoryStacks() {
  const trimStack = (stack) => {
    let totalBytes = stack.reduce((sum, snapshot) => sum + Number(snapshot.bytes || snapshot.html?.length || 0), 0);
    while (stack.length > MAX_EDITOR_HISTORY || (totalBytes > MAX_EDITOR_HISTORY_BYTES && stack.length > 1)) {
      const removed = stack.shift();
      totalBytes -= Number(removed?.bytes || removed?.html?.length || 0);
    }
    return stack;
  };
  editorUndoStack = trimStack(editorUndoStack);
  editorRedoStack = trimStack(editorRedoStack);
}

function resetInputHistoryCoalescing() {
  inputHistorySnapshotRecorded = false;
  if (inputHistoryResetTimer) {
    window.clearTimeout(inputHistoryResetTimer);
    inputHistoryResetTimer = null;
  }
}

function recordUndoSnapshot({ coalesceInput = false } = {}) {
  if (isLoadingDocument || suppressEditorHistory) return;
  if (coalesceInput && inputHistorySnapshotRecorded) return;
  const snapshot = createEditorSnapshot();
  const last = editorUndoStack[editorUndoStack.length - 1];
  if (last && last.html === snapshot.html) return;
  editorUndoStack.push(snapshot);
  editorRedoStack = [];
  trimEditorHistoryStacks();
  if (coalesceInput) {
    inputHistorySnapshotRecorded = true;
    if (inputHistoryResetTimer) window.clearTimeout(inputHistoryResetTimer);
    inputHistoryResetTimer = window.setTimeout(resetInputHistoryCoalescing, INPUT_HISTORY_DEBOUNCE_MS);
  } else {
    resetInputHistoryCoalescing();
  }
}

function resetEditorHistory() {
  resetInputHistoryCoalescing();
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
  document.title = `${isDirty ? "*" : ""}Mini DOCX Web Editor`;
}

function applyEditorZoom() {
  editor.style.zoom = String(editorZoom);
  requestAnimationFrame(() => {
    positionChapterFoldOverlay(visibleOutlineHeadings());
  });
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

const findReplaceState = {
  matches: [],
  currentIndex: -1,
  query: "",
  caseSensitive: false,
};
let findInputRefreshTimer = null;
let activeFindHighlightElement = null;
let pendingDeleteStructure = null;

function clearPersistentFindHighlight() {
  activeFindHighlightElement?.classList.remove("find-match-active");
  activeFindHighlightElement = null;
}

function setPersistentFindHighlight(element) {
  if (activeFindHighlightElement === element) return;
  clearPersistentFindHighlight();
  activeFindHighlightElement = element || null;
  activeFindHighlightElement?.classList.add("find-match-active");
}

function decodeFindReplaceText(value) {
  return String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\\(r\\n|n|r|t|\\)/g, (match, token) => {
      if (token === "r\\n" || token === "n" || token === "r") return "\n";
      if (token === "t") return "\t";
      return "\\";
    });
}

function searchDocumentSnapshot() {
  const blocks = allBlockElements();
  const segments = [];
  let text = "";
  blocks.forEach((block, index) => {
    if (index > 0) text += "\n";
    const blockText = blockPlainTextWithBreaks(block);
    const start = text.length;
    text += blockText;
    segments.push({ block, blockText, start, end: text.length });
  });
  return { blocks, segments, text };
}

function documentMatches(snapshot, query, caseSensitive) {
  if (!query) return [];
  const source = caseSensitive ? snapshot.text : snapshot.text.toLocaleLowerCase();
  const needle = caseSensitive ? query : query.toLocaleLowerCase();
  const matches = [];
  let offset = 0;
  while (offset <= source.length - needle.length) {
    const index = source.indexOf(needle, offset);
    if (index < 0) break;
    matches.push({ start: index, end: index + needle.length });
    offset = index + Math.max(needle.length, 1);
  }
  return matches;
}

function blockPointAtTextOffset(block, requestedOffset) {
  let remaining = Math.max(Number(requestedOffset) || 0, 0);
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
    acceptNode(node) {
      if (node.nodeType === Node.TEXT_NODE) return NodeFilter.FILTER_ACCEPT;
      if (node.nodeType === Node.ELEMENT_NODE && node.tagName === "BR" && node.dataset.editorPlaceholder !== "true") {
        return NodeFilter.FILTER_ACCEPT;
      }
      return NodeFilter.FILTER_SKIP;
    },
  });
  let node = walker.nextNode();
  while (node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = node.textContent?.length || 0;
      if (remaining <= length) return { node, offset: remaining };
      remaining -= length;
    } else {
      const parent = node.parentNode;
      const index = Array.prototype.indexOf.call(parent.childNodes, node);
      if (remaining === 0) return { node: parent, offset: index };
      remaining -= 1;
      if (remaining === 0) return { node: parent, offset: index + 1 };
    }
    node = walker.nextNode();
  }
  return { node: block, offset: block.childNodes.length };
}

function snapshotLocation(snapshot, offset) {
  let low = 0;
  let high = snapshot.segments.length - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const segment = snapshot.segments[middle];
    if (offset < segment.start) {
      high = middle - 1;
    } else if (offset > segment.end) {
      low = middle + 1;
    } else {
      return { ...segment, segmentIndex: middle, localOffset: offset - segment.start };
    }
  }
  if (high >= 0) {
    const previous = snapshot.segments[high];
    return {
      ...previous,
      segmentIndex: high,
      localOffset: previous.blockText.length,
    };
  }
  const last = snapshot.segments.at(-1);
  return last ? {
    ...last,
    segmentIndex: snapshot.segments.length - 1,
    localOffset: last.blockText.length,
  } : null;
}

function rangeForDocumentMatch(snapshot, match) {
  const startLocation = snapshotLocation(snapshot, match.start);
  const endLocation = snapshotLocation(snapshot, match.end);
  if (!startLocation || !endLocation) return null;
  const startPoint = blockPointAtTextOffset(startLocation.block, startLocation.localOffset);
  const endPoint = blockPointAtTextOffset(endLocation.block, endLocation.localOffset);
  const range = document.createRange();
  range.setStart(startPoint.node, startPoint.offset);
  range.setEnd(endPoint.node, endPoint.offset);
  return { range, startLocation, endLocation, startPoint, endPoint };
}

function updateFindMatchStatus(message = "") {
  if (!findMatchStatus) return;
  if (message) {
    findMatchStatus.textContent = message;
    return;
  }
  if (!decodeFindReplaceText(findTextInput?.value)) {
    findMatchStatus.textContent = "请输入查找内容";
  } else if (!findReplaceState.matches.length) {
    findMatchStatus.textContent = "未找到匹配内容";
  } else {
    findMatchStatus.textContent = `${findReplaceState.currentIndex + 1} / ${findReplaceState.matches.length}`;
  }
}

function highlightFindMatch() {
  const match = findReplaceState.matches[findReplaceState.currentIndex];
  if (!match) {
    updateFindMatchStatus();
    return false;
  }
  const snapshot = searchDocumentSnapshot();
  const matchRange = rangeForDocumentMatch(snapshot, match);
  if (!matchRange) return false;
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(matchRange.range);
  savedSelectionRange = matchRange.range.cloneRange();
  expandChapterContaining(matchRange.startLocation.block);
  matchRange.startLocation.block.scrollIntoView({ block: "center", behavior: "smooth" });
  setPersistentFindHighlight(matchRange.startLocation.block);
  updateFindMatchStatus();
  return true;
}

function refreshFindMatches(preferredIndex = 0) {
  const query = decodeFindReplaceText(findTextInput?.value);
  const caseSensitive = Boolean(findCaseSensitive?.checked);
  const snapshot = searchDocumentSnapshot();
  findReplaceState.query = query;
  findReplaceState.caseSensitive = caseSensitive;
  findReplaceState.matches = documentMatches(snapshot, query, caseSensitive);
  if (!findReplaceState.matches.length) {
    findReplaceState.currentIndex = -1;
  } else {
    findReplaceState.currentIndex = Math.max(Math.min(preferredIndex, findReplaceState.matches.length - 1), 0);
  }
  updateFindMatchStatus();
}

function moveFindMatch(delta) {
  const query = decodeFindReplaceText(findTextInput?.value);
  if (!query) {
    refreshFindMatches();
    findTextInput?.focus();
    return;
  }
  if (
    query !== findReplaceState.query
    || Boolean(findCaseSensitive?.checked) !== findReplaceState.caseSensitive
    || !findReplaceState.matches.length
  ) refreshFindMatches();
  if (!findReplaceState.matches.length) return;
  findReplaceState.currentIndex = (
    findReplaceState.currentIndex + delta + findReplaceState.matches.length
  ) % findReplaceState.matches.length;
  highlightFindMatch();
}

function openFindReplaceModal() {
  captureEditorSelection();
  findReplaceModal?.classList.remove("hidden");
  refreshFindMatches(Math.max(findReplaceState.currentIndex, 0));
  window.setTimeout(() => {
    findTextInput?.focus();
    findTextInput?.select();
  }, 0);
}

function closeFindReplaceModal() {
  findReplaceModal?.classList.add("hidden");
  clearPersistentFindHighlight();
  restoreEditorSelection();
  editor.focus();
}

function cloneParagraphShell(block) {
  const clone = block.cloneNode(false);
  clone.removeAttribute("id");
  return clone;
}

function removePlaceholderBreaks(fragment) {
  fragment.querySelectorAll?.('br[data-editor-placeholder="true"]').forEach((node) => node.remove());
  return fragment;
}

function ensureParagraphContent(block) {
  if (!block.hasChildNodes()) {
    const placeholder = document.createElement("br");
    placeholder.dataset.editorPlaceholder = "true";
    block.appendChild(placeholder);
  }
}

function appendReplacementText(block, text, descriptor) {
  if (!text) return;
  const span = document.createElement("span");
  span.textContent = text;
  applyDescriptorToRun(span, descriptor);
  block.appendChild(span);
}

function replaceDocumentMatch(match, replacementText, recordHistory = true, finalize = true) {
  const snapshot = searchDocumentSnapshot();
  const located = rangeForDocumentMatch(snapshot, match);
  if (!located) return false;
  const { startLocation, endLocation, startPoint } = located;
  const startBlock = startLocation.block;
  const endBlock = endLocation.block;
  if (!startBlock.parentNode || startBlock.parentNode !== endBlock.parentNode) {
    setStatus("暂不支持跨表格单元格替换。");
    return false;
  }

  const parent = startBlock.parentNode;
  const siblings = Array.from(parent.children);
  const startSiblingIndex = siblings.indexOf(startBlock);
  const endSiblingIndex = siblings.indexOf(endBlock);
  if (startSiblingIndex < 0 || endSiblingIndex < startSiblingIndex) return false;

  const beforeRange = document.createRange();
  beforeRange.setStart(startBlock, 0);
  beforeRange.setEnd(startPoint.node, startPoint.offset);
  const before = removePlaceholderBreaks(beforeRange.cloneContents());

  const endPoint = blockPointAtTextOffset(endBlock, endLocation.localOffset);
  const afterRange = document.createRange();
  afterRange.setStart(endPoint.node, endPoint.offset);
  afterRange.setEnd(endBlock, endBlock.childNodes.length);
  const after = removePlaceholderBreaks(afterRange.cloneContents());

  const contextNode = startPoint.node.nodeType === Node.TEXT_NODE
    ? startPoint.node.parentElement
    : startBlock;
  const descriptor = descriptorFromStyle(window.getComputedStyle(contextNode || startBlock));
  const replacementLines = String(replacementText).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const newBlocks = [];

  if (replacementLines.length === 1) {
    const block = cloneParagraphShell(startBlock);
    block.appendChild(before);
    appendReplacementText(block, replacementLines[0], descriptor);
    block.appendChild(after);
    ensureParagraphContent(block);
    newBlocks.push(block);
  } else {
    const first = cloneParagraphShell(startBlock);
    first.appendChild(before);
    appendReplacementText(first, replacementLines[0], descriptor);
    ensureParagraphContent(first);
    newBlocks.push(first);

    for (let index = 1; index < replacementLines.length - 1; index += 1) {
      const middle = cloneParagraphShell(startBlock);
      appendReplacementText(middle, replacementLines[index], descriptor);
      ensureParagraphContent(middle);
      newBlocks.push(middle);
    }

    const last = cloneParagraphShell(endBlock);
    appendReplacementText(last, replacementLines.at(-1), descriptor);
    last.appendChild(after);
    ensureParagraphContent(last);
    newBlocks.push(last);
  }

  if (recordHistory) recordUndoSnapshot();
  const insertion = document.createDocumentFragment();
  newBlocks.forEach((block) => insertion.appendChild(block));
  parent.insertBefore(insertion, startBlock);
  siblings.slice(startSiblingIndex, endSiblingIndex + 1).forEach((block) => block.remove());
  if (finalize) {
    markDirty();
    refreshOutline();
  }
  return true;
}

function replaceInlineDocumentMatch(snapshot, match, replacementText) {
  const located = rangeForDocumentMatch(snapshot, match);
  if (!located || located.startLocation.block !== located.endLocation.block) return false;
  const contextNode = located.range.startContainer.nodeType === Node.TEXT_NODE
    ? located.range.startContainer.parentElement
    : located.startLocation.block;
  const descriptor = descriptorFromStyle(window.getComputedStyle(contextNode || located.startLocation.block));
  located.range.deleteContents();
  if (replacementText) {
    const span = document.createElement("span");
    span.textContent = replacementText;
    applyDescriptorToRun(span, descriptor);
    located.range.insertNode(span);
  }
  return true;
}

function replaceCurrentFindMatch() {
  const match = findReplaceState.matches[findReplaceState.currentIndex];
  if (!match) {
    refreshFindMatches();
    return;
  }
  const replacement = decodeFindReplaceText(replaceTextInput?.value);
  const nextOffset = match.start + replacement.length;
  if (!replaceDocumentMatch(match, replacement)) return;
  refreshFindMatches(0);
  const nextIndex = findReplaceState.matches.findIndex((item) => item.start >= nextOffset);
  if (nextIndex >= 0) findReplaceState.currentIndex = nextIndex;
  highlightFindMatch();
  setStatus("已替换当前匹配项。");
}

function replaceAllFindMatches() {
  const query = decodeFindReplaceText(findTextInput?.value);
  if (!query) {
    findTextInput?.focus();
    return;
  }
  refreshFindMatches();
  const matches = [...findReplaceState.matches];
  if (!matches.length) return;
  const replacement = decodeFindReplaceText(replaceTextInput?.value);
  recordUndoSnapshot();
  let replaced = 0;
  const snapshot = searchDocumentSnapshot();
  const inlineBatch = !replacement.includes("\n") && matches.every((match) => {
    const start = snapshotLocation(snapshot, match.start);
    const end = snapshotLocation(snapshot, match.end);
    return start && end && start.block === end.block;
  });
  if (inlineBatch) {
    for (let index = matches.length - 1; index >= 0; index -= 1) {
      if (replaceInlineDocumentMatch(snapshot, matches[index], replacement)) replaced += 1;
    }
  } else {
    for (let index = matches.length - 1; index >= 0; index -= 1) {
      if (replaceDocumentMatch(matches[index], replacement, false, false)) replaced += 1;
    }
  }
  markDirty();
  refreshOutline();
  refreshFindMatches();
  updateFindMatchStatus(`已替换 ${replaced} 处`);
  setStatus(`已替换 ${replaced} 处。`);
}

function handleFindShortcut(event) {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) return false;
  if (String(event.key || "").toLowerCase() !== "f" && String(event.code || "") !== "KeyF") return false;
  event.preventDefault();
  event.stopPropagation();
  openFindReplaceModal();
  return true;
}

function defaultStyles() {
  // 字体族，字号，加粗，斜体，下划线
  return {
    paragraph: [
      { id: "Normal", name: "Normal", descriptor: [DEFAULT_FONT_FAMILY, 12, false, false, false], alignment: "left", outline_level: null, is_default: true, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "Heading1", name: "L0", descriptor: [DEFAULT_FONT_FAMILY, 20, true, false, false], alignment: "left", outline_level: 0, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "Heading2", name: "L1", descriptor: [DEFAULT_FONT_FAMILY, 16, true, false, false], alignment: "left", outline_level: 1, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "Heading3", name: "L2", descriptor: [DEFAULT_FONT_FAMILY, 14, true, false, false], alignment: "left", outline_level: 2, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "NormalL1", name: "L3", descriptor: [DEFAULT_FONT_FAMILY, 10, true, false, false], alignment: "left", outline_level: 3, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "NormalL2", name: "L4", descriptor: [DEFAULT_FONT_FAMILY, 10, true, true, false], alignment: "left", outline_level: 4, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
      { id: "NormalL3", name: "L5", descriptor: [DEFAULT_FONT_FAMILY, 10, true, true, true], alignment: "left", outline_level: 5, is_default: false, line_spacing: DEFAULT_LINE_SPACING, space_before: DEFAULT_PARAGRAPH_SPACING, space_after: DEFAULT_PARAGRAPH_SPACING },
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
  if (key === "code") return "Normal";
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
  formatPainterBtn?.classList.add("is-active");
  setStatus("格式刷已开启，请点击目标段落。");
}

function clearFormatPainter() {
  formatPainterPayload = null;
  formatPainterBtn?.classList.remove("is-active");
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

function normalizeBlockIndentLevel(value) {
  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed)) return 0;
  return Math.max(Math.min(parsed, MAX_BLOCK_INDENT_LEVEL), 0);
}

function applyBlockIndent(element, indentLevel) {
  if (!element) return;
  const level = normalizeBlockIndentLevel(indentLevel);
  const style = getStyleById(element.dataset.styleId || styleIdFromTag(element.tagName));
  const styleIndent = Math.max(Number(style?.outline_level) || 0, 0);
  element.dataset.indentLevel = String(level);
  // Heading indentation belongs to the L0–L5 style, not to whitespace that
  // users type into the title. Manual paragraph indentation is additive.
  const totalIndent = styleIndent + (level * BLOCK_INDENT_STEP_EM);
  element.style.marginLeft = totalIndent > 0 ? `${totalIndent}em` : "0";
}

function indentLevelFromElement(element) {
  if (!element) return 0;
  return normalizeBlockIndentLevel(element.dataset.indentLevel || 0);
}

function updateBlockIndent(delta) {
  restoreEditorSelection();
  const blocks = selectedBlockElements();
  if (!blocks.length) {
    setStatus("请先选中段落。");
    return;
  }
  recordUndoSnapshot();
  blocks.forEach((block) => applyBlockIndent(block, indentLevelFromElement(block) + delta));
  markDirty();
  refreshOutline();
  setStatus(`已更新段落缩进（${blocks.length} 段）。`);
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

function replaceTag(element, tagName) {
  if (element.tagName === tagName) return element;
  const replacement = document.createElement(tagName.toLowerCase());
  Array.from(element.attributes).forEach((attr) => replacement.setAttribute(attr.name, attr.value));
  replacement.innerHTML = element.innerHTML;
  replacement.style.cssText = element.style.cssText;
  element.replaceWith(replacement);
  if (activePrimaryOutlineElement === element) activePrimaryOutlineElement = replacement;
  if (activeSecondaryOutlineElement === element) activeSecondaryOutlineElement = replacement;
  if (activeFindHighlightElement === element) activeFindHighlightElement = replacement;
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
    setEmptyBlockPlaceholder(p);
    p.dataset.styleId = "Normal";
    cell.appendChild(p);
  }
}

function setEmptyBlockPlaceholder(block) {
  if (!block) return;
  const placeholder = document.createElement("br");
  placeholder.dataset.editorPlaceholder = "true";
  block.replaceChildren(placeholder);
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
  setEmptyBlockPlaceholder(p);
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
  return Array.from(editor.querySelectorAll("p, h1, h2, h3"));
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

function estimateNumberingLabelWidthPx(block, label, fontSpecCache = null) {
  const text = String(label || "");
  if (numberingMeasureContext) {
    const cacheKey = block
      ? `${block.dataset.styleId || block.tagName}|${block.style.fontFamily}|${block.style.fontSize}|${block.style.fontWeight}|${block.style.fontStyle}`
      : "editor";
    let fontSpec = fontSpecCache?.get(cacheKey);
    if (!fontSpec) {
      fontSpec = fontSpecFromComputedStyle(window.getComputedStyle(block || editor));
      fontSpecCache?.set(cacheKey, fontSpec);
    }
    numberingMeasureContext.font = fontSpec;
  }
  const measured = numberingMeasureContext ? numberingMeasureContext.measureText(text).width : (text.length * 8);
  return Math.max(Math.ceil(measured) + 1, MIN_NUMBERING_PREFIX_PX);
}

function refreshNumberingVisuals() {
  const states = new Map();
  const visualItems = [];
  const fontSpecCache = new Map();
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
      prefixWidthPx: estimateNumberingLabelWidthPx(block, label, fontSpecCache),
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
  return Array.from(block.querySelectorAll("span, font, b, strong, i, em, u, s, code")).filter((node) => {
    if (node.tagName === "FONT") return true;
    if (["B", "STRONG", "I", "EM", "U", "S", "CODE"].includes(node.tagName)) return true;
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
  const descriptorWithoutBackground = cloneDescriptor(nextDescriptor);
  descriptorWithoutBackground[5] = "";
  const runs = inlineStyledRuns(block);
  if (!runs.length) {
    Array.from(block.childNodes).forEach((node) => {
      if (node.nodeType !== Node.TEXT_NODE || !node.textContent) return;
      const span = document.createElement("span");
      span.textContent = node.textContent;
      applyDescriptorToRun(span, descriptorWithoutBackground);
      node.replaceWith(span);
    });
    return;
  }
  runs.forEach((run) => {
    const preservedBackground = normalizeBackgroundColor(
      run.style.backgroundColor || window.getComputedStyle(run).backgroundColor,
    );
    applyDescriptorToRun(run, descriptorWithoutBackground);
    if (preservedBackground) run.style.backgroundColor = preservedBackground;
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
  if (!pageWidthInput || !pageHeightInput) return;
  const widthMm = twipsToMm(pageSize.widthTwips);
  const heightMm = twipsToMm(pageSize.heightTwips);
  pageWidthInput.value = String(widthMm || 210);
  pageHeightInput.value = String(heightMm || 297);
}

function setPageSizeFromMm(widthMm, heightMm) {
  const width = Math.max(Number(widthMm) || 0, 50);
  const height = Math.max(Number(heightMm) || 0, 50);
  pageSize = { widthTwips: mmToTwips(width), heightTwips: mmToTwips(height) };
  applyPageSizeToEditor();
  syncPageSizeControls();
  pageDirty = true;
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
  const outlineLevel = Number(style.outline_level);
  if (Number.isFinite(outlineLevel) && outlineLevel >= 0) {
    target.dataset.outlineIndent = String(outlineLevel);
    target.style.setProperty("--outline-indent", `${outlineLevel}em`);
  } else {
    delete target.dataset.outlineIndent;
    target.style.removeProperty("--outline-indent");
  }
  applyParagraphMetrics(target, {
    lineSpacing: style.line_spacing || DEFAULT_LINE_SPACING,
    spaceBefore: style.space_before ?? DEFAULT_PARAGRAPH_SPACING,
    spaceAfter: style.space_after ?? 0,
  });
  applyBlockIndent(target, indentLevelFromElement(target));
  return target;
}

function currentBlockElement() {
  const range = activeEditorRange();
  if (!range) return null;
  let node = range.startContainer;
  if (!node) return null;
  if (node.nodeType === Node.TEXT_NODE) node = node.parentNode;
  const block = node && node.closest ? node.closest("p, h1, h2, h3") : null;
  return block && nodeInEditor(block) ? block : null;
}

function blockElementFromNode(node) {
  if (!node) return null;
  let target = node;
  if (target.nodeType === Node.TEXT_NODE) target = target.parentNode;
  const block = target && target.closest ? target.closest("p, h1, h2, h3") : null;
  return block && nodeInEditor(block) ? block : null;
}

function collectBlocksBetween(startBlock, endBlock) {
  if (!startBlock || !endBlock || !nodeInEditor(startBlock) || !nodeInEditor(endBlock)) return [];
  const blocks = allBlockElements();
  const startIndex = blocks.indexOf(startBlock);
  const endIndex = blocks.indexOf(endBlock);
  if (startIndex < 0 || endIndex < 0) return [];
  return blocks.slice(Math.min(startIndex, endIndex), Math.max(startIndex, endIndex) + 1);
}

function selectedBlockElements() {
  normalizeEditorStructure();
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) {
    debugLog("prefix:selectedBlocks:none", { reason: "no-selection-or-range" });
  return [];
}

  if (selection.isCollapsed) {
    const block = currentBlockElement();
    const result = block ? [block] : [];
    debugLog("prefix:selectedBlocks:collapsed", {
      anchorText: selection.anchorNode?.textContent?.slice(0, 80) || "",
      resultCount: result.length,
      resultTags: result.map((item) => item.tagName),
      resultTexts: result.map((item) => (item.textContent || "").slice(0, 120)),
    });
    return result;
  }
  const range = selection.getRangeAt(0);
  const anchorBlock = blockElementFromNode(selection.anchorNode);
  const focusBlock = blockElementFromNode(selection.focusNode);
  const startBlock = blockElementFromNode(range.startContainer);
  const endBlock = blockElementFromNode(range.endContainer);
  if (anchorBlock && focusBlock && anchorBlock !== focusBlock) {
    const between = collectBlocksBetween(anchorBlock, focusBlock);
    if (between.length) {
      debugLog("prefix:selectedBlocks:anchor-focus", {
        anchorTag: anchorBlock.tagName,
        focusTag: focusBlock.tagName,
        resultCount: between.length,
        resultTags: between.map((item) => item.tagName),
        resultTexts: between.map((item) => (item.textContent || "").slice(0, 120)),
      });
      return between;
    }
  }
  if (startBlock && endBlock && startBlock !== endBlock) {
    const between = collectBlocksBetween(startBlock, endBlock);
    if (between.length) {
      debugLog("prefix:selectedBlocks:range", {
        startTag: startBlock.tagName,
        endTag: endBlock.tagName,
        resultCount: between.length,
        resultTags: between.map((item) => item.tagName),
        resultTexts: between.map((item) => (item.textContent || "").slice(0, 120)),
      });
      return between;
    }
  }
  const selected = Array.from(editor.querySelectorAll("p, h1, h2, h3")).filter((block) => {
    try {
      return range.intersectsNode(block);
    } catch {
      return false;
    }
  });
  if (selected.length) {
    const deduped = Array.from(new Set(selected));
    debugLog("prefix:selectedBlocks:intersects", {
      resultCount: deduped.length,
      resultTags: deduped.map((item) => item.tagName),
      resultTexts: deduped.map((item) => (item.textContent || "").slice(0, 120)),
      rangeText: range.toString().slice(0, 200),
    });
    return deduped;
  }
  const block = currentBlockElement();
  const result = block ? [block] : [];
  debugLog("prefix:selectedBlocks:fallback", {
    resultCount: result.length,
    resultTags: result.map((item) => item.tagName),
    resultTexts: result.map((item) => (item.textContent || "").slice(0, 120)),
  });
  return result;
}

function blockPlainTextWithBreaks(block) {
  if (!block) return "";
  if (block.childNodes.length === 1 && block.firstChild?.nodeName === "BR") return "";
  const lines = [];
  let currentLine = "";

  const flushLine = () => {
    lines.push(currentLine);
    currentLine = "";
  };

  const walk = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      currentLine += node.textContent || "";
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }
    if (node.tagName === "BR") {
      if (node.dataset.editorPlaceholder === "true") return;
      flushLine();
      return;
    }
    if (node.tagName === "IMG") {
      return;
    }
    Array.from(node.childNodes).forEach(walk);
  };

  Array.from(block.childNodes).forEach(walk);
  lines.push(currentLine);
  return lines.join("\n");
}

function setBlockPlainTextWithBreaks(block, text) {
  if (!block) return;
  const normalized = String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  block.innerHTML = "";
  if (!lines.length) {
    block.appendChild(document.createElement("br"));
    return;
  }
  lines.forEach((line, index) => {
    if (line) {
      block.appendChild(document.createTextNode(line));
    }
    if (index < lines.length - 1) {
      block.appendChild(document.createElement("br"));
    }
  });
  if (!normalized || normalized.endsWith("\n")) {
    const placeholder = document.createElement("br");
    placeholder.dataset.editorPlaceholder = "true";
    block.appendChild(placeholder);
  }
}

function adjustSelectedBlockLinePrefixes(mode) {
  const blocks = selectedBlockElements().filter((block) => ["P", "H1", "H2", "H3"].includes(block.tagName));
  debugLog("prefix:adjust:start", {
    mode,
    blockCount: blocks.length,
    blocks: blocks.map((block, index) => ({
      index,
      tag: block.tagName,
      styleId: block.dataset.styleId || styleIdFromTag(block.tagName),
      text: blockPlainTextWithBreaks(block).slice(0, 300),
    })),
  });
  if (!blocks.length) {
    setStatus("请先选中要处理的段落。");
    return;
  }
  recordUndoSnapshot();
  const changedBlocks = [];
  blocks.forEach((block) => {
    const source = blockPlainTextWithBreaks(block);
    const sourceLines = source
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\n");
    const nextLines = sourceLines.map((line) => {
      if (mode === "indent") return `  ${line}`;
      if (/^[ \u00A0]{2}/.test(line)) return line.slice(2);
      if (/^[ \u00A0]/.test(line)) return line.slice(1);
      return line;
    });
    const next = nextLines.join("\n");
    debugLog("prefix:adjust:block", {
      mode,
      tag: block.tagName,
      styleId: block.dataset.styleId || styleIdFromTag(block.tagName),
      source,
      sourceLines,
      nextLines,
      next,
      changed: next !== source,
    });
    if (next === source) return;
    setBlockPlainTextWithBreaks(block, next);
    debugLog("prefix:adjust:block:after-write", {
      mode,
      tag: block.tagName,
      styleId: block.dataset.styleId || styleIdFromTag(block.tagName),
      afterText: blockPlainTextWithBreaks(block),
      afterHtml: block.innerHTML,
    });
    changedBlocks.push(block);
  });
  if (!changedBlocks.length) {
    debugLog("prefix:adjust:no-change", { mode });
    setStatus(mode === "indent" ? "选中段落已全部带有前导空格。" : "选中段落没有可删除的前导空格。");
    return;
  }
  selectBlockRange(changedBlocks[0], changedBlocks[changedBlocks.length - 1]);
  refreshOutline();
  markDirty();
  debugLog("prefix:adjust:done", {
    mode,
    changedCount: changedBlocks.length,
    changedTexts: changedBlocks.map((block) => blockPlainTextWithBreaks(block).slice(0, 300)),
  });
  setStatus(mode === "indent" ? `已为 ${changedBlocks.length} 段逐行添加 2 个前导空格。` : `已为 ${changedBlocks.length} 段逐行删除前导空格。`);
}

  function isRightPrefixShortcut(event) {
    const key = String(event.key || "");
    const code = String(event.code || "");
    const keyCode = Number(event.keyCode || event.which || 0);
    return [">", "》", ".", "。", "=", "+"].includes(key) || code === "Period" || code === "Equal" || keyCode === 190 || keyCode === 187;
  }

  function isLeftPrefixShortcut(event) {
    const key = String(event.key || "");
    const code = String(event.code || "");
    const keyCode = Number(event.keyCode || event.which || 0);
    return ["<", "《", ",", "，", "-", "_"].includes(key) || code === "Comma" || code === "Minus" || keyCode === 188 || keyCode === 189;
  }

  function handlePrefixShortcut(event, source = "document") {
    if (event.__prefixShortcutHandled) return true;
    const target = event.target;
    const editingShortcut = target && target.closest && target.closest(".shortcut-item");
    if (editingShortcut) return false;

    debugLog("prefix:keydown", {
      source,
      key: event.key,
      code: event.code,
      keyCode: Number(event.keyCode || event.which || 0),
      ctrlKey: event.ctrlKey,
      metaKey: event.metaKey,
      shiftKey: event.shiftKey,
      altKey: event.altKey,
      selectionText: window.getSelection()?.toString()?.slice(0, 200) || "",
      targetTag: target?.tagName || null,
      targetText: target?.textContent?.slice(0, 120) || "",
    });

    if (!(event.ctrlKey || event.metaKey)) return false;
    const editorActive = selectionInsideEditor() || document.activeElement === editor || editor.contains(document.activeElement);
    if (!editorActive) return false;

    if (isRightPrefixShortcut(event)) {
      event.__prefixShortcutHandled = true;
      debugLog("prefix:shortcut-match", { source, side: "right", key: event.key, code: event.code, keyCode: Number(event.keyCode || event.which || 0), shiftKey: event.shiftKey });
      event.preventDefault();
      event.stopPropagation();
      adjustSelectedBlockLinePrefixes("indent");
      return true;
    }
    if (isLeftPrefixShortcut(event)) {
      event.__prefixShortcutHandled = true;
      debugLog("prefix:shortcut-match", { source, side: "left", key: event.key, code: event.code, keyCode: Number(event.keyCode || event.which || 0), shiftKey: event.shiftKey });
      event.preventDefault();
      event.stopPropagation();
      adjustSelectedBlockLinePrefixes("outdent");
      return true;
    }
    return false;
  }

  function handleGlobalSaveShortcut(event, source = "document") {
    const target = event.target;
    const editingShortcut = target && target.closest && target.closest(".shortcut-item");
    if (editingShortcut) return false;
    if (!(event.ctrlKey || event.metaKey) || event.altKey) return false;
    if (String(event.key || "").toLowerCase() !== "s" && String(event.code || "") !== "KeyS") return false;
    debugLog("save:shortcut-match", { source, key: event.key, code: event.code });
    event.preventDefault();
    event.stopPropagation();
    Promise.resolve(saveDocx({ interactive: false, allowPicker: true })).catch(handleAsyncError);
    return true;
  }

function moveCaretToBlockStart(block) {
  if (!block) return;
  const range = document.createRange();
  range.selectNodeContents(block);
  range.collapse(true);
  const selection = window.getSelection();
  if (!selection) return;
  selection.removeAllRanges();
  selection.addRange(range);
  captureEditorSelection();
}

function selectBlockRange(startBlock, endBlock) {
  if (!startBlock || !endBlock) return;
  const range = document.createRange();
  range.setStartBefore(startBlock);
  range.setEndAfter(endBlock);
  const selection = window.getSelection();
  if (!selection) return;
  selection.removeAllRanges();
  selection.addRange(range);
  captureEditorSelection();
}

function allBlockElements() {
  normalizeEditorStructure();
  return Array.from(editor.querySelectorAll("p, h1, h2, h3"));
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

function captureDeleteStructure() {
  return Array.from(editor.children)
    .filter((element) => ["P", "H1", "H2", "H3"].includes(element.tagName))
    .map((element) => ({
      element,
      styleId: element.dataset.styleId || styleIdFromTag(element.tagName),
      text: element.textContent || "",
      wasEmpty: !(element.textContent || "").trim(),
    }));
}

function restoreStyleAfterDeletedHeading(snapshot) {
  if (!snapshot?.length) return;
  for (let index = 0; index < snapshot.length - 1; index += 1) {
    const headingState = snapshot[index];
    const nextState = snapshot[index + 1];
    const headingStyle = getStyleById(headingState.styleId);
    if (headingStyle?.outline_level !== 0) continue;
    if (!headingState.element.isConnected || nextState.element.isConnected) continue;

    const currentText = headingState.element.textContent || "";
    const nextText = nextState.text || "";
    const headingWasFullyRemoved = headingState.wasEmpty
      || (currentText === nextText && currentText !== headingState.text);
    if (!headingWasFullyRemoved || currentText !== nextText) continue;

    const nextStyle = getStyleById(nextState.styleId) || getStyleById("Normal");
    const restored = applyStyleVisuals(headingState.element, nextStyle);
    debugLog("delete:restore-following-style", {
      deletedStyleId: headingState.styleId,
      restoredStyleId: nextStyle?.id || "Normal",
      textLength: currentText.length,
    });
    if (restored) moveCaretToBlockStart(restored);
    break;
  }
}

function shouldRefreshDocumentChrome(event, deleteSnapshot) {
  const inputType = String(event?.inputType || "");
  if (deleteSnapshot || inputType === "insertParagraph" || inputType === "insertFromPaste" || inputType === "insertFromDrop") {
    return true;
  }
  const block = currentBlockElement();
  return isHeadingStyleBlock(block) || Boolean(numberingFromElement(block));
}

function handleEditorInput(event) {
  const deleteSnapshot = pendingDeleteStructure;
  pendingDeleteStructure = null;
  if (deleteSnapshot) restoreStyleAfterDeletedHeading(deleteSnapshot);
  if (shouldRefreshDocumentChrome(event, deleteSnapshot)) {
    scheduleEditorRefresh();
  }
  markDirty();
}

function scheduleEditorRefresh() {
  if (editorRefreshTimer) window.clearTimeout(editorRefreshTimer);
  editorRefreshTimer = window.setTimeout(() => {
    editorRefreshTimer = null;
    refreshOutline();
  }, EDITOR_REFRESH_DEBOUNCE_MS);
}

function slugify(text, fallback) {
  const slug = text
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
  return slug || fallback;
}

function flashOutlineTarget(target) {
  if (!target) return;
  target.classList.remove("outline-target-flash");
  void target.offsetWidth;
  target.classList.add("outline-target-flash");
  window.setTimeout(() => {
    target.classList.remove("outline-target-flash");
  }, 1600);
}

function outlineLevelForElement(element) {
  if (!element || !["P", "H1", "H2", "H3"].includes(element.tagName)) return null;
  const styleId = element.dataset.styleId || styleIdFromTag(element.tagName);
  const style = getStyleById(styleId);
  const fallback = /^H[1-3]$/.test(element.tagName) ? Number(element.tagName.slice(1)) - 1 : null;
  const level = style?.outline_level;
  return level !== null && level !== undefined && Number.isFinite(Number(level)) ? Number(level) : fallback;
}

function isOutlineHeading(element) {
  return Number.isFinite(outlineLevelForElement(element));
}

function directEditorChildForNode(node) {
  let element = node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement;
  while (element && element.parentElement !== editor) element = element.parentElement;
  return element?.parentElement === editor ? element : null;
}

function collapsedOutlineAncestorsForElement(element) {
  const directChild = directEditorChildForNode(element);
  if (!directChild) return [];
  const children = Array.from(editor.children);
  const targetIndex = children.indexOf(directChild);
  if (targetIndex < 0) return [];
  const ancestors = [];
  for (let index = 0; index <= targetIndex; index += 1) {
    const candidate = children[index];
    const level = outlineLevelForElement(candidate);
    if (!Number.isFinite(level)) continue;
    while (ancestors.length && ancestors[ancestors.length - 1].level >= level) ancestors.pop();
    if (candidate.dataset.chapterCollapsed === "true") ancestors.push({ element: candidate, level });
  }
  return ancestors.map((ancestor) => ancestor.element);
}

function outlineHeadingForElement(element) {
  const directChild = directEditorChildForNode(element);
  if (!directChild) return null;
  const children = Array.from(editor.children);
  const targetIndex = children.indexOf(directChild);
  if (targetIndex < 0) return null;
  const headings = [];
  for (let index = 0; index <= targetIndex; index += 1) {
    const candidate = children[index];
    const level = outlineLevelForElement(candidate);
    if (!Number.isFinite(level)) continue;
    while (headings.length && headings[headings.length - 1].level >= level) headings.pop();
    headings.push({ element: candidate, level });
  }
  return headings[headings.length - 1]?.element || null;
}

function createChapterFoldToggle(heading) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "chapter-fold-toggle";
  button.dataset.editorUi = "true";
  button.tabIndex = -1;
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const collapsing = heading.dataset.chapterCollapsed !== "true";
    if (collapsing) {
      const selectionChild = directEditorChildForNode(window.getSelection()?.anchorNode);
      if (selectionChild && outlineHeadingForElement(selectionChild) === heading && selectionChild !== heading) {
        moveCaretToBlockStart(heading);
      }
    }
    heading.dataset.chapterCollapsed = collapsing ? "true" : "false";
    applyChapterFolding();
    setStatus(collapsing ? "已折叠章节。" : "已展开章节。");
  });
  const collapsed = heading.dataset.chapterCollapsed === "true";
  button.setAttribute("aria-label", collapsed ? "展开章节" : "折叠章节");
  button.title = collapsed ? "展开章节" : "折叠章节";
  button.setAttribute("aria-expanded", collapsed ? "false" : "true");
  button.classList.toggle("is-collapsed", collapsed);
  return button;
}

function removeLegacyChapterFoldControls() {
  editor.querySelectorAll(".chapter-fold-toggle, [data-editor-ui='true']").forEach((control) => {
    const heading = directEditorChildForNode(control);
    control.remove();
    if (!heading || heading.textContent.trim()) return;
    const breaks = Array.from(heading.children).filter((child) => child.tagName === "BR");
    if (breaks.length <= 1) return;
    breaks.slice(1).forEach((lineBreak) => lineBreak.remove());
    breaks[0].dataset.editorPlaceholder = "true";
  });
}

function normalizeEmptyChapterHeading(heading) {
  if (heading.textContent.trim()) return;
  const breaks = Array.from(heading.children).filter((child) => child.tagName === "BR");
  if (breaks.length <= 1) return;
  breaks.slice(1).forEach((lineBreak) => lineBreak.remove());
  breaks[0].dataset.editorPlaceholder = "true";
}

function positionChapterFoldOverlay(headings) {
  if (!chapterFoldOverlay) return;
  chapterFoldOverlay.replaceChildren();
  const stageRect = pageStage.getBoundingClientRect();
  const editorRect = editor.getBoundingClientRect();
  const editorPaddingLeft = Number.parseFloat(window.getComputedStyle(editor).paddingLeft) || 0;
  const fixedButtonLeft = editorRect.left - stageRect.left + pageStage.scrollLeft + editorPaddingLeft - 30;
  // Measure every heading before modifying the overlay.  Appending one button
  // at a time between getBoundingClientRect calls forces repeated reflow on
  // large documents.
  const placements = headings.map((heading) => {
    const headingRect = heading.getBoundingClientRect();
    return {
      heading,
      left: fixedButtonLeft,
      top: headingRect.top - stageRect.top + pageStage.scrollTop + (headingRect.height / 2),
    };
  });
  const fragment = document.createDocumentFragment();
  placements.forEach(({ heading, left, top }) => {
    const button = createChapterFoldToggle(heading);
    button.style.left = `${left}px`;
    button.style.top = `${top}px`;
    fragment.appendChild(button);
  });
  chapterFoldOverlay.appendChild(fragment);
}

function scheduleChapterFoldOverlay(headings) {
  pendingFoldOverlayHeadings = headings;
  if (chapterFoldOverlayFrame !== null) return;
  chapterFoldOverlayFrame = window.requestAnimationFrame(() => {
    chapterFoldOverlayFrame = null;
    positionChapterFoldOverlay(pendingFoldOverlayHeadings || []);
    pendingFoldOverlayHeadings = null;
  });
}

function applyChapterFolding({ deferOverlay = false, normalizeHeadings = true } = {}) {
  removeLegacyChapterFoldControls();
  const collapsedAncestors = [];
  const headings = [];
  Array.from(editor.children).forEach((child) => {
    const level = outlineLevelForElement(child);
    if (Number.isFinite(level)) {
      while (collapsedAncestors.length && collapsedAncestors[collapsedAncestors.length - 1].level >= level) {
        collapsedAncestors.pop();
      }
      const isHidden = collapsedAncestors.length > 0;
      if (normalizeHeadings) normalizeEmptyChapterHeading(child);
      child.classList.toggle("chapter-folded-content", isHidden);
      child.classList.toggle("chapter-is-collapsed", child.dataset.chapterCollapsed === "true");
      if (!isHidden) headings.push(child);
      if (child.dataset.chapterCollapsed === "true") collapsedAncestors.push({ element: child, level });
      return;
    }
    child.classList.toggle("chapter-folded-content", collapsedAncestors.length > 0);
  });
  if (deferOverlay) {
    scheduleChapterFoldOverlay(headings);
  } else {
    positionChapterFoldOverlay(headings);
  }
}

function visibleOutlineHeadings() {
  return Array.from(editor.children).filter((child) => isOutlineHeading(child) && !child.classList.contains("chapter-folded-content"));
}

function setAllChaptersCollapsed(collapsed) {
  const headings = Array.from(editor.children).filter(isOutlineHeading);
  headings.forEach((heading) => {
    heading.dataset.chapterCollapsed = collapsed ? "true" : "false";
  });
  if (collapsed) {
    const selectedChild = directEditorChildForNode(window.getSelection()?.anchorNode);
    const selectedHeading = outlineHeadingForElement(selectedChild);
    if (selectedHeading && selectedChild !== selectedHeading) moveCaretToBlockStart(selectedHeading);
  }
  applyChapterFolding({ deferOverlay: true, normalizeHeadings: false });
  setStatus(collapsed ? `已折叠 ${headings.length} 个标题层级。` : `已展开 ${headings.length} 个标题层级。`);
}

function expandChapterContaining(element) {
  const collapsedHeadings = collapsedOutlineAncestorsForElement(element);
  if (!collapsedHeadings.length) return;
  collapsedHeadings.forEach((heading) => { heading.dataset.chapterCollapsed = "false"; });
  applyChapterFolding();
}

function scrollOutlineTargetIntoView(target) {
  if (!target) return;
  const stageRect = pageStage.getBoundingClientRect();
  const targetRect = target.getBoundingClientRect();
  const currentTop = pageStage.scrollTop + (targetRect.top - stageRect.top);
  const topPadding = 48;
  const bottomPadding = 96;
  const visibleTop = pageStage.scrollTop;
  const visibleBottom = visibleTop + pageStage.clientHeight;
  const targetTop = currentTop;
  const targetBottom = currentTop + Math.max(targetRect.height, 1);

  let nextTop = null;
  if (targetTop < visibleTop + topPadding) {
    nextTop = Math.max(targetTop - topPadding, 0);
  } else if (targetBottom > visibleBottom - bottomPadding) {
    nextTop = Math.max(targetTop - topPadding, 0);
  }

  if (nextTop !== null) {
    pageStage.scrollTo({ top: nextTop, behavior: "smooth" });
  }
}

function focusOutlineTarget(target) {
  if (!target) return;
  expandChapterContaining(target);
  scrollOutlineTargetIntoView(target);
  flashOutlineTarget(target);
  const range = document.createRange();
  range.selectNodeContents(target);
  range.collapse(false);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  editor.focus();
  syncParagraphStyleSelect();
}

function directOutlineItems() {
  return allBlockElements()
    .map(outlineItemFromBlock)
    .filter((item) => directEditorChildForNode(item.element) === item.element && Number.isFinite(item.level));
}

function outlineParentElement(items, item) {
  const index = items.findIndex((candidate) => candidate.element === item.element);
  if (index < 0) return null;
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    if (items[cursor].level < item.level) return items[cursor].element;
  }
  return null;
}

function outlineBranchNodes(items, item) {
  const index = items.findIndex((candidate) => candidate.element === item.element);
  const start = directEditorChildForNode(item.element);
  if (index < 0 || !start) return [];
  let end = null;
  for (let cursor = index + 1; cursor < items.length; cursor += 1) {
    if (items[cursor].level <= item.level) {
      end = directEditorChildForNode(items[cursor].element);
      break;
    }
  }
  const children = Array.from(editor.children);
  const startIndex = children.indexOf(start);
  const endIndex = end ? children.indexOf(end) : children.length;
  return startIndex < 0 ? [] : children.slice(startIndex, endIndex < 0 ? children.length : endIndex);
}

function canMoveOutlineItem(source, target, scope) {
  if (!source || !target || source.element === target.element || source.level !== target.level) return false;
  const items = directOutlineItems();
  if (!items.some((item) => item.element === source.element) || !items.some((item) => item.element === target.element)) return false;
  if (source.level === 0) return true;
  // Nested headings can only be reordered among siblings, so a drag never
  // changes the heading hierarchy by accident.
  return outlineParentElement(items, source) === outlineParentElement(items, target);
}

function moveOutlineBranch(source, target, after, scope) {
  if (!canMoveOutlineItem(source, target, scope)) return false;
  const items = directOutlineItems();
  const sourceNodes = outlineBranchNodes(items, source);
  const targetNodes = outlineBranchNodes(items, target);
  if (!sourceNodes.length || !targetNodes.length) return false;

  recordUndoSnapshot();
  const sourceSet = new Set(sourceNodes);
  const remaining = Array.from(editor.children).filter((node) => !sourceSet.has(node));
  const targetStart = remaining.indexOf(targetNodes[0]);
  const targetEnd = remaining.indexOf(targetNodes[targetNodes.length - 1]);
  if (targetStart < 0 || targetEnd < 0) return false;
  const reference = after ? remaining[targetEnd + 1] : remaining[targetStart];
  const fragment = document.createDocumentFragment();
  sourceNodes.forEach((node) => fragment.appendChild(node));
  editor.insertBefore(fragment, reference || null);

  activePrimaryOutlineElement = source.element;
  activePrimaryOutlineBlockIndex = null;
  activeSecondaryOutlineElement = source.element;
  activeSecondaryOutlineBlockIndex = null;
  markDirty();
  refreshOutline();
  setStatus(`已移动标题：${source.textContent.trim() || "未命名标题"}`);
  return true;
}

function clearOutlineDropIndicators(container) {
  container?.querySelectorAll(".outline-item.drag-before, .outline-item.drag-after, .outline-item.is-dragging")
    .forEach((node) => node.classList.remove("drag-before", "drag-after", "is-dragging"));
}

function renderOutlineButtons(container, items, emptyText, options = {}) {
  if (!container) return;
  const previousScrollTop = container.scrollTop;
  const previousActiveButton = container.querySelector(".outline-item.is-active");
  const previousActiveOffset = previousActiveButton
    ? previousActiveButton.offsetTop - previousScrollTop
    : null;
  container.innerHTML = "";
  if (!items.length) {
    const empty = document.createElement("div");
    empty.textContent = emptyText;
    empty.className = "outline-item";
    container.appendChild(empty);
    container.scrollTop = previousScrollTop;
    return;
  }
  let nextActiveButton = null;
  items.forEach((item, index) => {
    const text = item.textContent.trim() || `标题 ${index + 1}`;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "outline-item";
    button.textContent = text;
    button.draggable = Boolean(options.dragScope);
    const indentLevel = Math.max(0, Number(item.level));
    button.classList.add(`level-${indentLevel}`);
    // Navigation indentation is structural: it reflects the outline level,
    // rather than whitespace inserted into the heading text.
    button.style.paddingLeft = `${indentLevel * 12}px`;
    if (item.element === options.activeElement) {
      button.classList.add("is-active");
      nextActiveButton = button;
    }
    button.addEventListener("click", () => {
      container.querySelectorAll(".outline-item.is-active").forEach((node) => node.classList.remove("is-active"));
      button.classList.add("is-active");
      options.onItemClick?.(item);
      focusOutlineTarget(item.element);
    });
    if (options.dragScope) {
      button.addEventListener("dragstart", (event) => {
        outlineDragItem = item;
        button.classList.add("is-dragging");
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", String(item.blockIndex));
      });
      button.addEventListener("dragover", (event) => {
        if (!canMoveOutlineItem(outlineDragItem, item, options.dragScope)) return;
        event.preventDefault();
        clearOutlineDropIndicators(container);
        const after = event.clientY > button.getBoundingClientRect().top + (button.getBoundingClientRect().height / 2);
        button.classList.add(after ? "drag-after" : "drag-before");
        event.dataTransfer.dropEffect = "move";
      });
      button.addEventListener("dragleave", () => button.classList.remove("drag-before", "drag-after"));
      button.addEventListener("drop", (event) => {
        if (!canMoveOutlineItem(outlineDragItem, item, options.dragScope)) return;
        event.preventDefault();
        const rect = button.getBoundingClientRect();
        moveOutlineBranch(outlineDragItem, item, event.clientY > rect.top + (rect.height / 2), options.dragScope);
        clearOutlineDropIndicators(container);
      });
      button.addEventListener("dragend", () => {
        outlineDragItem = null;
        clearOutlineDropIndicators(container);
      });
    }
    container.appendChild(button);
  });
  if (nextActiveButton && previousActiveOffset !== null) {
    container.scrollTop = Math.max(nextActiveButton.offsetTop - previousActiveOffset, 0);
  } else {
    container.scrollTop = previousScrollTop;
  }
}

function outlineItemFromBlock(block, blockIndex) {
  const styleId = block.dataset.styleId || styleIdFromTag(block.tagName);
  const style = getStyleById(styleId);
  const fallbackLevel = /^H[1-3]$/.test(block.tagName) ? Number(block.tagName.slice(1)) - 1 : NaN;
  const rawLevel = style?.outline_level;
  const level = rawLevel !== null && rawLevel !== undefined && rawLevel !== "" && Number.isFinite(Number(rawLevel))
    ? Number(rawLevel)
    : fallbackLevel;
  return {
    element: block,
    blockIndex,
    level,
    textContent: block.textContent || "",
  };
}

function displayOutlineLevel(item) {
  return Number(item?.level);
}

function availableOutlineLevelCeiling(allItems = []) {
  const documentLevels = allItems.map(displayOutlineLevel).filter(Number.isFinite);
  const styleLevels = currentStyles.paragraph
    .map((style) => Number(style.outline_level))
    .filter((level) => Number.isFinite(level) && level >= 0);
  return Math.max(5, ...documentLevels, ...styleLevels);
}

function normalizeOutlineConfig(maxLevel) {
  const ceiling = Math.max(Number(maxLevel) || 5, 1);
  const maxPrimaryLevel = Math.max(ceiling - 1, 0);
  outlineConfig.primaryMin = Math.min(Math.max(Number(outlineConfig.primaryMin) || 0, 0), maxPrimaryLevel);
  outlineConfig.primaryMax = Math.min(
    Math.max(Number(outlineConfig.primaryMax) || outlineConfig.primaryMin, outlineConfig.primaryMin),
    maxPrimaryLevel,
  );
  outlineConfig.secondaryMax = Math.min(
    Math.max(Number(outlineConfig.secondaryMax) || ceiling, outlineConfig.primaryMax + 1),
    ceiling,
  );
  for (let level = outlineConfig.primaryMin; level <= outlineConfig.primaryMax; level += 1) {
    if (outlineFilter[level] === undefined) outlineFilter[level] = true;
  }
  return ceiling;
}

function renderPrimaryOutlineLevelToggles() {
  if (!primaryOutlineLevelToggles) return;
  primaryOutlineLevelToggles.innerHTML = "";
  const otherLabel = document.createElement("label");
  otherLabel.className = "outline-other-toggle";
  const otherInput = document.createElement("input");
  otherInput.type = "checkbox";
  otherInput.checked = showOtherOutlineBranches;
  otherInput.addEventListener("change", () => {
    showOtherOutlineBranches = otherInput.checked;
    persistOutlineFilter();
    refreshOutline();
  });
  otherLabel.append(otherInput, " Other");
  primaryOutlineLevelToggles.appendChild(otherLabel);

  for (let level = outlineConfig.primaryMin; level <= outlineConfig.primaryMax; level += 1) {
    const levelLabel = document.createElement("label");
    const levelInput = document.createElement("input");
    levelInput.type = "checkbox";
    levelInput.checked = outlineFilter[level] !== false;
    levelInput.dataset.outlineLevel = String(level);
    levelInput.addEventListener("change", () => {
      outlineFilter[level] = levelInput.checked;
      persistOutlineFilter();
      refreshOutline();
    });
    levelLabel.append(levelInput, ` L${level}`);
    primaryOutlineLevelToggles.appendChild(levelLabel);
  }
}

function syncOutlineConfigControls(maxLevel, renderToggles = true) {
  const ceiling = normalizeOutlineConfig(maxLevel);
  if (primaryOutlineMinSlider) {
    primaryOutlineMinSlider.min = "0";
    primaryOutlineMinSlider.max = String(Math.max(ceiling - 1, 0));
    primaryOutlineMinSlider.value = String(outlineConfig.primaryMin);
  }
  if (primaryOutlineMaxSlider) {
    primaryOutlineMaxSlider.min = String(outlineConfig.primaryMin);
    primaryOutlineMaxSlider.max = String(Math.max(ceiling - 1, 0));
    primaryOutlineMaxSlider.value = String(outlineConfig.primaryMax);
  }
  if (secondaryOutlineMaxSlider) {
    secondaryOutlineMaxSlider.min = String(outlineConfig.primaryMax + 1);
    secondaryOutlineMaxSlider.max = String(ceiling);
    secondaryOutlineMaxSlider.value = String(outlineConfig.secondaryMax);
  }
  if (primaryOutlineMinText) primaryOutlineMinText.textContent = `L${outlineConfig.primaryMin}`;
  if (primaryOutlineMaxText) primaryOutlineMaxText.textContent = `L${outlineConfig.primaryMax}`;
  if (secondaryOutlineMaxText) secondaryOutlineMaxText.textContent = `L${outlineConfig.secondaryMax}`;
  if (renderToggles) renderPrimaryOutlineLevelToggles();
}

function initOutlineConfigControls() {
  const updateFromSliders = (source) => {
    const ceiling = availableOutlineLevelCeiling();
    if (source === "min") {
      outlineConfig.primaryMin = Number(primaryOutlineMinSlider.value);
      if (outlineConfig.primaryMax < outlineConfig.primaryMin) outlineConfig.primaryMax = outlineConfig.primaryMin;
    } else if (source === "max") {
      outlineConfig.primaryMax = Number(primaryOutlineMaxSlider.value);
      if (outlineConfig.primaryMin > outlineConfig.primaryMax) outlineConfig.primaryMin = outlineConfig.primaryMax;
      if (outlineConfig.secondaryMax <= outlineConfig.primaryMax) outlineConfig.secondaryMax = outlineConfig.primaryMax + 1;
    } else if (source === "secondary") {
      outlineConfig.secondaryMax = Number(secondaryOutlineMaxSlider.value);
    }
    syncOutlineConfigControls(ceiling);
    persistOutlineFilter();
    refreshOutline();
  };
  primaryOutlineMinSlider?.addEventListener("input", () => updateFromSliders("min"));
  primaryOutlineMaxSlider?.addEventListener("input", () => updateFromSliders("max"));
  secondaryOutlineMaxSlider?.addEventListener("input", () => updateFromSliders("secondary"));
  syncOutlineConfigControls(availableOutlineLevelCeiling());
}

function nearestOutlineItem(items, blockIndex) {
  if (!items.length || !Number.isFinite(blockIndex)) return null;
  return items.reduce((nearest, item) => {
    if (!nearest) return item;
    const itemDistance = Math.abs(item.blockIndex - blockIndex);
    const nearestDistance = Math.abs(nearest.blockIndex - blockIndex);
    if (itemDistance < nearestDistance) return item;
    if (itemDistance > nearestDistance) return nearest;
    return item.blockIndex <= blockIndex && nearest.blockIndex > blockIndex ? item : nearest;
  }, null);
}

function rememberPrimaryOutlineItem(item, manuallySelected = false) {
  if (!item) return;
  activePrimaryOutlineElement = item.element;
  activePrimaryOutlineBlockIndex = item.blockIndex;
  if (manuallySelected) primaryOutlineWasManuallySelected = true;
}

function rememberSecondaryOutlineItem(item, manuallySelected = false) {
  if (!item) return;
  activeSecondaryOutlineElement = item.element;
  activeSecondaryOutlineBlockIndex = item.blockIndex;
  if (manuallySelected) secondaryOutlineWasManuallySelected = true;
}

function resetOutlineNavigationState() {
  activePrimaryOutlineElement = null;
  activePrimaryOutlineBlockIndex = null;
  primaryOutlineWasManuallySelected = false;
  activeSecondaryOutlineElement = null;
  activeSecondaryOutlineBlockIndex = null;
  secondaryOutlineWasManuallySelected = false;
  if (outline) outline.scrollTop = 0;
  if (secondaryOutline) secondaryOutline.scrollTop = 0;
}

function resolvePrimaryOutlineItem(allItems, visiblePrimaryItems) {
  const eligibleItems = allItems.filter((item) => {
    const level = displayOutlineLevel(item);
    return level >= outlineConfig.primaryMin && level <= outlineConfig.primaryMax && item.textContent.trim();
  });
  const currentItem = eligibleItems.find((item) => item.element === activePrimaryOutlineElement);
  if (currentItem) {
    rememberPrimaryOutlineItem(currentItem);
    return currentItem;
  }

  const nearbyItem = nearestOutlineItem(eligibleItems, activePrimaryOutlineBlockIndex);
  if (nearbyItem) {
    rememberPrimaryOutlineItem(nearbyItem);
    return nearbyItem;
  }

  if (!primaryOutlineWasManuallySelected && visiblePrimaryItems.length) {
    rememberPrimaryOutlineItem(visiblePrimaryItems[0]);
    return visiblePrimaryItems[0];
  }

  activePrimaryOutlineElement = null;
  return null;
}

function secondaryItemsForAnchor(allItems, anchorElement) {
  const anchorIndex = allItems.findIndex((item) => item.element === anchorElement);
  if (anchorIndex < 0) return [];
  const anchorLevel = allItems[anchorIndex].level;
  const descendants = [];
  for (let index = anchorIndex + 1; index < allItems.length; index += 1) {
    const item = allItems[index];
    if (item.level <= anchorLevel) break;
    const displayLevel = displayOutlineLevel(item);
    if (
      item.level > anchorLevel
      && displayLevel > outlineConfig.primaryMax
      && displayLevel <= outlineConfig.secondaryMax
      && item.textContent.trim()
    ) {
      descendants.push(item);
    }
  }
  return descendants;
}

function outlineAncestorIndex(allItems, anchorIndex, level) {
  if (anchorIndex < 0) return -1;
  for (let index = anchorIndex; index >= 0; index -= 1) {
    const itemLevel = displayOutlineLevel(allItems[index]);
    if (itemLevel === level) return index;
    if (itemLevel < level) return -1;
  }
  return -1;
}

function primarySiblingBranchBounds(allItems, anchorElement) {
  if (showOtherOutlineBranches) return null;
  const anchorIndex = allItems.findIndex((item) => item.element === anchorElement);
  if (anchorIndex < 0) return null;
  const anchorLevel = displayOutlineLevel(allItems[anchorIndex]);
  if (!Number.isFinite(anchorLevel)) return null;
  const parentLevel = anchorLevel - 1;
  if (parentLevel < outlineConfig.primaryMin) return { start: 0, end: allItems.length, level: anchorLevel };

  const parentStart = outlineAncestorIndex(allItems, anchorIndex, parentLevel);
  if (parentStart < 0) return null;
  let parentEnd = allItems.length;
  for (let index = parentStart + 1; index < allItems.length; index += 1) {
    if (displayOutlineLevel(allItems[index]) <= parentLevel) {
      parentEnd = index;
      break;
    }
  }
  return { start: parentStart, end: parentEnd, level: anchorLevel };
}

function renderSecondaryOutline(allItems) {
  const availableItems = secondaryItemsForAnchor(allItems, activePrimaryOutlineElement);
  const visibleItems = availableItems;
  if (secondaryOutlineWasManuallySelected) {
    const currentItem = availableItems.find((item) => item.element === activeSecondaryOutlineElement);
    const resolvedItem = currentItem || nearestOutlineItem(availableItems, activeSecondaryOutlineBlockIndex);
    if (resolvedItem) {
      rememberSecondaryOutlineItem(resolvedItem);
    } else {
      activeSecondaryOutlineElement = null;
    }
  }

  if (secondaryOutlineLevelText) {
    secondaryOutlineLevelText.textContent = `L${outlineConfig.primaryMax + 1}–L${outlineConfig.secondaryMax}`;
  }
  renderOutlineButtons(secondaryOutline, visibleItems, "当前标题下暂无子节", {
    activeElement: activeSecondaryOutlineElement,
    dragScope: "secondary",
    onItemClick: (item) => rememberSecondaryOutlineItem(item, true),
  });
}

function refreshOutline() {
  normalizeEditorStructure();
  applyChapterFolding();
  refreshNumberingVisuals();
  const allItems = allBlockElements()
    .map(outlineItemFromBlock)
    .filter((item) => Number.isFinite(item.level) && item.textContent.trim());
  syncOutlineConfigControls(availableOutlineLevelCeiling(allItems), false);
  const primaryItems = allItems.filter((item) => {
    const displayLevel = displayOutlineLevel(item);
    return displayLevel >= outlineConfig.primaryMin
      && displayLevel <= outlineConfig.primaryMax
      && outlineFilter[displayLevel] !== false;
  });

  const activeItem = resolvePrimaryOutlineItem(allItems, primaryItems);
  const branchBounds = primarySiblingBranchBounds(allItems, activeItem?.element || activePrimaryOutlineElement);
  const visiblePrimaryItems = branchBounds
    ? primaryItems.filter((item) => item.blockIndex >= allItems[branchBounds.start].blockIndex
      && item.blockIndex < (allItems[branchBounds.end]?.blockIndex ?? Number.POSITIVE_INFINITY)
      && displayOutlineLevel(item) === branchBounds.level)
    : primaryItems;
  const safeVisiblePrimaryItems = visiblePrimaryItems.length ? visiblePrimaryItems : primaryItems;

  renderOutlineButtons(outline, safeVisiblePrimaryItems, "暂无标题导航", {
    activeElement: activePrimaryOutlineElement,
    dragScope: "primary",
    onItemClick: (item) => {
      rememberPrimaryOutlineItem(item, true);
      activeSecondaryOutlineElement = null;
      activeSecondaryOutlineBlockIndex = null;
      secondaryOutlineWasManuallySelected = false;
      if (secondaryOutline) secondaryOutline.scrollTop = 0;
      refreshOutline();
    },
  });
  renderSecondaryOutline(allItems);
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
  currentFilePath = null;
  currentFileName = file.name || "mini-docx.docx";
  await openDocx(file);
  try {
    await recordRecentFile(fileHandle);
  } catch {
    // Ignore storage errors for recent list.
  }
  updateFileUiState();
  updateCurrentFileStatus();
}

async function openDocxPicker() {
  if (!confirmDiscardChanges()) return;
  const response = await fetch("/api/pick-open-docx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: "{}",
  });
  const result = await response.json();
  if (result.cancelled) {
    setStatus("操作已取消");
    return;
  }
  if (!response.ok || !result.ok) {
    throw new Error(result.error || "打开文件失败。");
  }
  loadDocument(result.document);
  currentFileHandle = null;
  currentFilePath = windowsPath(result.path);
  currentFileName = result.name || "mini-docx.docx";
  updateFileUiState();
  updateCurrentFileStatus();
  setStatus(`已打开：${currentFilePath}`);
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

function isSoleEmptyBlockBreak(node) {
  const block = node?.closest?.("p, h1, h2, h3");
  if (!block || block.querySelector("img") || (block.textContent || "").length) return false;
  return block.querySelectorAll("br").length === 1;
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
  if (node.dataset.editorUi === "true" || node.classList.contains("chapter-fold-toggle")) {
    return;
  }
  if (node.tagName === "IMG") {
    return;
  }
  if (node.tagName === "BR") {
    if (node.dataset.editorPlaceholder !== "true" && !isSoleEmptyBlockBreak(node)) {
      bucket.push({ text: "\n", descriptor: inheritedStyle });
    }
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
      indent_level: indentLevelFromElement(element),
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
        ...(img.dataset.mediaToken
          ? { media_token: img.dataset.mediaToken }
          : {
            mime: (img.src.match(/^data:([^;]+);base64,/) || [])[1] || "image/png",
            data_url: img.src,
          }),
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
    meta.page_dirty = Boolean(pageDirty);
    payload._docx_meta = meta;
  }
  return payload;
}

function appendRuns(parent, runs) {
  const fragment = document.createDocumentFragment();
  (runs || []).forEach((run) => {
    const [family, size, bold, italic, underline, background] = cloneDescriptor(run.descriptor);
    const span = document.createElement("span");
    const lines = String(run.text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    span.style.fontFamily = family;
    span.style.fontSize = `${Math.max(Number(size) / 0.75, 1)}px`;
    span.style.fontWeight = bold ? "700" : "400";
    span.style.fontStyle = italic ? "italic" : "normal";
    span.style.textDecoration = underline ? "underline" : "none";
    if (background) {
      span.style.backgroundColor = background;
    }
    lines.forEach((line, index) => {
      if (index) span.appendChild(document.createElement("br"));
      span.appendChild(document.createTextNode(line));
    });
    fragment.appendChild(span);
  });
  parent.appendChild(fragment);
}

function renderParagraphBlock(block, parent) {
  const style = getStyleById(block.style_id || styleIdFromBlockStyleKey(block.style));
  const el = document.createElement(tagNameFromStyle(style).toLowerCase());
  el.dataset.styleId = (style && style.id) || "Normal";
  el.style.textAlign = { align_left: "left", align_center: "center", align_right: "right", align_justify: "justify" }[block.alignment] || (style ? style.alignment : "left");
  appendRuns(el, block.runs);
  applyStyleVisuals(el, getStyleById(el.dataset.styleId));
  el.style.textAlign = { align_left: "left", align_center: "center", align_right: "right", align_justify: "justify" }[block.alignment] || el.style.textAlign;
  applyParagraphMetrics(el, {
    lineSpacing: block.line_spacing || DEFAULT_LINE_SPACING,
    spaceBefore: block.space_before ?? DEFAULT_PARAGRAPH_SPACING,
    spaceAfter: block.space_after ?? 0,
  });
  applyBlockIndent(el, block.indent_level || 0);
  setNumberingData(el, block.numbering);
  if (!el.textContent.trim()) {
    setEmptyBlockPlaceholder(el);
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
  resetOutlineNavigationState();
  clearFormatPainter();
  docxMeta = cloneDocxMeta(documentData._docx_meta);
  stylesDirty = false;
  numberingDirty = false;
  pageDirty = false;
  currentStyles = normalizeStyles(documentData.styles);
  populateStyleSelect();
  pageSize = {
    widthTwips: Number(documentData.page?.width_twips) || mmToTwips(DEFAULT_DOCUMENT_WIDTH_MM),
    heightTwips: Number(documentData.page?.height_twips) || mmToTwips(DEFAULT_DOCUMENT_HEIGHT_MM),
  };
  applyPageSizeToEditor();
  syncPageSizeControls();
  const fragment = document.createDocumentFragment();
  (documentData.blocks || []).forEach((block) => {
    if (block.type === "table") {
      fragment.appendChild(renderTableBlock(block));
      return;
    }
    if (block.type === "image") {
      const p = document.createElement("p");
      const img = document.createElement("img");
      p.dataset.styleId = "Normal";
      img.src = block.media_token ? `/api/media/${encodeURIComponent(block.media_token)}` : block.data_url;
      img.alt = block.name || "image";
      img.dataset.name = block.name || "image.png";
      if (block.media_token) img.dataset.mediaToken = block.media_token;
      if (block.width_px) img.style.width = `${block.width_px}px`;
      if (block.height_px) img.style.height = "auto";
      p.appendChild(img);
      fragment.appendChild(p);
      return;
    }
    renderParagraphBlock(block, fragment);
  });
  editor.replaceChildren(fragment);
  ensureStarterContent();
  refreshOutline();
  isLoadingDocument = false;
  resetEditorHistory();
  markClean();
}

function applyParagraphStyle(styleId) {
  // A toolbar control takes focus away from the editor. Only restore the
  // cached range in that case: restoring it unconditionally can revive an
  // old multi-paragraph selection after the user has clicked a single line.
  if (!selectionInsideEditor()) {
    restoreEditorSelection();
  }
  const style = getStyleById(styleId);
  if (!style) return;
  // A plain click means "this paragraph only".  Keep that intent even after
  // the style selector has taken focus, rather than reviving an older range.
  const blocks = lastClickedParagraph && nodeInEditor(lastClickedParagraph)
    ? [lastClickedParagraph]
    : selectedBlockElements();
  if (!blocks.length) {
    setStatus("请先把光标放在要应用样式的段落中。");
    paragraphStyleSelect.value = styleId;
    return;
  }
  recordUndoSnapshot();
  const updatedBlocks = blocks.map((block) => {
    const previousStyleId = block.dataset.styleId || styleIdFromTag(block.tagName);
    const previousStyle = getStyleById(previousStyleId);
    const previousDescriptor = cloneDescriptor(previousStyle?.descriptor);
    const updatedBlock = applyStyleVisuals(block, style);
    syncStyledRunsForStyleUpdate(updatedBlock, previousDescriptor, style.descriptor);
    return updatedBlock;
  });
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
    const configRaw = window.localStorage.getItem(OUTLINE_CONFIG_KEY);
    const config = configRaw ? JSON.parse(configRaw) : null;
    if (config && typeof config === "object") {
      outlineConfig = {
        primaryMin: Math.max(Number(config.primaryMin) || 0, 0),
        primaryMax: Math.max(Number(config.primaryMax) || 3, 1),
        secondaryMax: Math.max(Number(config.secondaryMax) || 5, 1),
      };
      const visibleLevels = config.visibleLevels && typeof config.visibleLevels === "object"
        ? config.visibleLevels
        : {};
      outlineFilter = Object.fromEntries(
        Object.entries(visibleLevels).map(([level, visible]) => [Number(level), Boolean(visible)]),
      );
      showOtherOutlineBranches = config.showOtherBranches !== false;
      return;
    }

    const legacyConfigRaw = window.localStorage.getItem(LEGACY_OUTLINE_CONFIG_KEY);
    const legacyConfig = legacyConfigRaw ? JSON.parse(legacyConfigRaw) : null;
    if (legacyConfig && typeof legacyConfig === "object") {
      outlineConfig = {
        primaryMin: Math.max((Number(legacyConfig.primaryMin) || 1) - 1, 0),
        primaryMax: Math.max((Number(legacyConfig.primaryMax) || 3) - 1, 0),
        secondaryMax: Math.max((Number(legacyConfig.secondaryMax) || 6) - 1, 1),
      };
      const visibleLevels = legacyConfig.visibleLevels && typeof legacyConfig.visibleLevels === "object"
        ? legacyConfig.visibleLevels
        : {};
      outlineFilter = Object.fromEntries(
        Object.entries(visibleLevels).map(([level, visible]) => [Math.max(Number(level) - 1, 0), Boolean(visible)]),
      );
      return;
    }

    const legacyRaw = window.localStorage.getItem(OUTLINE_FILTER_KEY);
    const parsed = legacyRaw ? JSON.parse(legacyRaw) : null;
    if (parsed && typeof parsed === "object") {
      outlineFilter = {
        0: Boolean(parsed[1]),
        1: Boolean(parsed[2]),
        2: Boolean(parsed[3]),
      };
    }
  } catch {
    // Ignore storage errors.
  }
}

function persistOutlineFilter() {
  try {
    window.localStorage.setItem(OUTLINE_CONFIG_KEY, JSON.stringify({
      ...outlineConfig,
      visibleLevels: outlineFilter,
      showOtherBranches: showOtherOutlineBranches,
    }));
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
  nextDescriptor[5] = previousDescriptor[5] || "";
  const nextOutlineLevel = (() => {
    if (style.id === "Normal") return null;
    if (block.tagName === "H1") return 0;
    if (block.tagName === "H2") return 1;
    if (block.tagName === "H3") return 2;
    return style.outline_level ?? null;
  })();
  style.descriptor = nextDescriptor;
  style.outline_level = nextOutlineLevel;
  style.alignment = { align_left: "left", align_center: "center", align_right: "right", align_justify: "justify" }[blockAlignment(block)] || "left";
  style.line_spacing = Number(metrics.lineSpacing) || style.line_spacing || DEFAULT_LINE_SPACING;
  style.space_before = Number(metrics.spaceBefore) || DEFAULT_PARAGRAPH_SPACING;
  style.space_after = Number(metrics.spaceAfter) || DEFAULT_PARAGRAPH_SPACING;

  const targets = allBlockElements().filter((target) => {
    const targetStyleId = target.dataset.styleId || styleIdFromTag(target.tagName);
    return targetStyleId === style.id;
  });
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
  const response = await fetch("/api/import-docx", {
    method: "POST",
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "X-Filename": encodeURIComponent(file.name || "mini-docx.docx"),
    },
    body: file,
  });
  const result = await response.json();
  if (!response.ok) {
    throw new Error(result.error || "打开文件失败。");
  }
  loadDocument(result.document);
  currentFilePath = null;
  currentFileName = file.name || currentFileName;
  updateFileUiState();
  updateCurrentFileStatus();
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
    throw new Error(result.error || "导出 DOCX 失败。");
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
  const targetName = currentFileName || "mini-docx.docx";
  try {
    if (!currentFilePath) {
      if (!interactive && !allowPicker) {
        setStatus("当前文件尚未选择保存位置。");
        return;
      }
      const pathResponse = await fetch("/api/pick-save-path", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ suggested_name: targetName, current_path: currentFilePath }),
      });
      const pathResult = await pathResponse.json();
      if (pathResult.cancelled) {
        setStatus("操作已取消");
        return;
      }
      if (!pathResponse.ok || !pathResult.ok) {
        throw new Error(pathResult.error || "选择保存位置失败。");
      }
      currentFilePath = windowsPath(pathResult.path);
      currentFileName = pathResult.name || targetName;
      currentFileHandle = null;
      updateFileUiState();
      updateCurrentFileStatus();
    }

    setStatus(`正在保存：${currentFilePath}`);
    const response = await fetch("/api/save-docx-path", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: currentFilePath, document: editorToDocument() }),
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.error || "保存文件失败。");
    }
    currentFilePath = windowsPath(result.path);
    currentFileName = result.name || currentFileName;
    markClean();
    updateCurrentFileStatus();
    setStatus(`已保存：${currentFilePath}`);
  } catch (error) {
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
  resetOutlineNavigationState();
  docxMeta = null;
  stylesDirty = false;
  numberingDirty = false;
  currentStyles = normalizeStyles(defaultStyles());
  currentFileHandle = null;
  currentFilePath = null;
  currentFileName = "mini-docx.docx";
  populateStyleSelect("Normal");
  editor.innerHTML = "";
  ensureStarterContent();
  resetEditorHistory();
  markClean();
  updateCurrentFileStatus();
  setStatus("已新建文档");
});

openBtn.addEventListener("click", async () => {
  try {
    await openDocxPicker();
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

fontFamily?.addEventListener("change", () => {
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

fontSize?.addEventListener("change", () => {
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

if (applyPageSizeBtn && pageWidthInput && pageHeightInput) {
  applyPageSizeBtn.addEventListener("click", () => {
    setPageSizeFromMm(pageWidthInput.value, pageHeightInput.value);
  });
}

paragraphStyleSelect.addEventListener("change", () => applyParagraphStyle(paragraphStyleSelect.value));
lineSpacingSelect.addEventListener("change", applyCurrentParagraphMetrics);
spaceBeforeInput.addEventListener("change", applyCurrentParagraphMetrics);
spaceAfterInput.addEventListener("change", applyCurrentParagraphMetrics);
toggleNumberingBtn?.addEventListener("click", toggleNumbering);
numberFormatSelect?.addEventListener("change", applyNumberFormatToSelection);
clearFormatBtn?.addEventListener("click", () => {
  exec("removeFormat");
  setStatus("已清除选区文字格式。");
});
saveStyleBtn.addEventListener("click", saveCurrentStyle);
updateStyleBtn.addEventListener("click", updateStyleFromSelection);
formatPainterBtn?.addEventListener("click", () => {
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
closeFindReplaceBtn?.addEventListener("click", closeFindReplaceModal);
findReplaceModal?.addEventListener("click", (event) => {
  if (event.target.dataset.closeModal === "true") closeFindReplaceModal();
});
document.addEventListener("pointerdown", (event) => {
  if (!activeFindHighlightElement) return;
  if (findReplaceModal?.contains(event.target)) return;
  clearPersistentFindHighlight();
}, true);
findTextInput?.addEventListener("input", () => {
  window.clearTimeout(findInputRefreshTimer);
  updateFindMatchStatus("正在查找…");
  findInputRefreshTimer = window.setTimeout(() => refreshFindMatches(), 120);
});
findCaseSensitive?.addEventListener("change", () => refreshFindMatches());
findPreviousBtn?.addEventListener("click", () => moveFindMatch(-1));
findNextBtn?.addEventListener("click", () => moveFindMatch(1));
replaceCurrentBtn?.addEventListener("click", replaceCurrentFindMatch);
replaceAllBtn?.addEventListener("click", replaceAllFindMatches);
[findTextInput, replaceTextInput].forEach((input) => {
  input?.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      event.preventDefault();
      closeFindReplaceModal();
      return;
    }
    if (input === findTextInput && event.key === "Enter" && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      moveFindMatch(event.shiftKey ? -1 : 1);
    }
  });
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

editor.addEventListener("input", handleEditorInput);
editor.addEventListener("beforeinput", (event) => {
  if (suppressEditorHistory || isLoadingDocument) return;
  if (event.inputType === "historyUndo" || event.inputType === "historyRedo") return;
  pendingDeleteStructure = (event.inputType || "").startsWith("delete") ? captureDeleteStructure() : null;
  recordUndoSnapshot({ coalesceInput: true });
});
editor.addEventListener("click", () => {
  // Capture synchronously so the next toolbar operation uses this click's
  // collapsed caret range rather than an earlier multi-paragraph selection.
  captureEditorSelection();
  const selection = window.getSelection();
  lastClickedParagraph = selection?.rangeCount && selection.getRangeAt(0).collapsed
    ? currentBlockElement()
    : null;
  syncParagraphStyleSelect();
});
editor.addEventListener("click", (event) => {
  const block = event.target.closest && event.target.closest("p, h1, h2, h3");
  if (formatPainterPayload && block) {
    applyFormatPainterToBlock(block);
  }
});
editor.addEventListener("paste", async (event) => {
  const clipboard = event.clipboardData;
  if (!clipboard) {
    window.setTimeout(scheduleEditorRefresh, 20);
    return;
  }
  const items = Array.from(clipboard.items || []);
  const imageItems = items.filter((item) => item.type && item.type.startsWith("image/"));
  if (!imageItems.length) {
    window.setTimeout(scheduleEditorRefresh, 20);
    return;
  }
  event.preventDefault();
  setStatus("插入图片功能已移除。");
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
  const currentStyleId = current?.dataset?.styleId || styleIdFromTag(current?.tagName || "");
  const currentNumbering = numberingFromElement(current);
  const hasText = (current?.textContent || "").trim().length > 0;
  const isHeadingBlock = ["Heading1", "Heading2", "Heading3"].includes(currentStyleId);
  if (!current) return;
  if (isHeadingBlock) {
    event.preventDefault();
    recordUndoSnapshot();
    const nextBlock = document.createElement("p");
    nextBlock.dataset.styleId = "Normal";
    setEmptyBlockPlaceholder(nextBlock);
    applyStyleVisuals(nextBlock, getStyleById("Normal"));
    current.insertAdjacentElement("afterend", nextBlock);
    moveCaretToBlockStart(nextBlock);
    syncParagraphStyleSelect();
    refreshOutline();
    markDirty();
    return;
  }
  window.setTimeout(() => {
    const nextBlock = currentBlockElement();
    if (!nextBlock || nextBlock === current) return;

    if (!currentNumbering || !hasText || numberingFromElement(nextBlock)) {
      return;
    }
    setNumberingData(nextBlock, currentNumbering);
    refreshOutline();
    markDirty();
  }, 0);
});
document.addEventListener("selectionchange", () => {
  if (document.activeElement === editor || editor.contains(document.activeElement) || editor.contains(window.getSelection()?.anchorNode)) {
    const selection = window.getSelection();
    if (selection?.rangeCount && !selection.getRangeAt(0).collapsed) {
      lastClickedParagraph = null;
    }
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

window.addEventListener("keydown", (event) => {
  if (handleFindShortcut(event)) {
    return;
  }
  if (handleGlobalSaveShortcut(event, "window-capture")) {
    return;
  }
  if (handlePrefixShortcut(event, "window-capture")) {
    return;
  }
}, true);

document.addEventListener("keydown", (event) => {
  if (handleFindShortcut(event)) {
    return;
  }
  if (handleGlobalSaveShortcut(event, "document")) {
    return;
  }
  const target = event.target;
  if (handlePrefixShortcut(event, "document")) {
    return;
  }

  if (event.ctrlKey || event.metaKey) {
    if (event.key === "]") {
      if (selectionInsideEditor() || document.activeElement === editor || editor.contains(document.activeElement)) {
        event.preventDefault();
        updateBlockIndent(1);
        return;
      }
    }
    if (event.key === "[") {
      if (selectionInsideEditor() || document.activeElement === editor || editor.contains(document.activeElement)) {
        event.preventDefault();
        updateBlockIndent(-1);
        return;
      }
    }
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

window.addEventListener("resize", () => {
  requestAnimationFrame(() => {
    positionChapterFoldOverlay(visibleOutlineHeadings());
  });
});

window.addEventListener("beforeunload", (event) => {
  if (!isDirty) return;
  event.preventDefault();
  event.returnValue = "";
});

currentStyles = normalizeStyles(defaultStyles());
loadOutlineFilter();
initOutlineConfigControls();
loadShortcuts();
initLayoutToggles();
initAdvancedToolGroups();
if (saveStyleBtn) {
  saveStyleBtn.title = "保存到当前选中的段落样式";
}
cleanResourcesBtn?.addEventListener("click", cleanResources);
collapseAllChaptersBtn?.addEventListener("click", () => setAllChaptersCollapsed(true));
expandAllChaptersBtn?.addEventListener("click", () => setAllChaptersCollapsed(false));
let resourceRefreshTimer = null;
function scheduleResourceRefresh() {
  if (resourceRefreshTimer) window.clearInterval(resourceRefreshTimer);
  resourceRefreshTimer = null;
  if (document.hidden) return;
  refreshResourceStats();
  resourceRefreshTimer = window.setInterval(refreshResourceStats, RESOURCE_REFRESH_MS);
}
document.addEventListener("visibilitychange", scheduleResourceRefresh);
scheduleResourceRefresh();
populateStyleSelect("Normal");
renderShortcutInputs();
setServerAddress();
updateFileUiState();
updateCurrentFileStatus();
applyEditorZoom();
ensureStarterContent();
applyPageSizeToEditor();
syncPageSizeControls();
resetEditorHistory();
markClean();
