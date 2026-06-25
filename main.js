'use strict';
const { app, BrowserWindow, shell, Menu, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const { pathToFileURL } = require('url');
const mammoth = require('mammoth');
const preloadPath = path.join(__dirname, 'preload.js');

process.env.LANG = 'ko_KR.UTF-8';

if (process.platform === 'win32') {
  try {
    app.setAppUserModelId("com.web2electron.mdpro");
  } catch (_) {}
}

const HELP_FILE = null;
const RHWP_URL = 'https://edwardkim.github.io/rhwp/';
const LIBREOFFICE_DOWNLOAD_URL = 'https://www.libreoffice.org/download/';
let helpWindow = null;
let activePptxWindow = null;
let externalOpenPrefs = null;

const DEFAULT_EXTERNAL_OPEN_PREFS = {
  pdf: false,
  doc: false,
  docx: false,
  csv: false,
  xls: true,
  xlsx: true,
  ppt: true,
  pptx: true,
  hwp: false,
  hwpx: false,
  alwaysNewWindowPptx: false,
};

const OPEN_DIALOG_EXTENSIONS = [
  'md', 'markdown', 'mdown', 'txt', 'csv', 'html', 'htm', 'json', 'mpv', 'mpp', 'mdd',
  'doc', 'docx', 'hwp', 'hwpx', 'xls', 'xlsx', 'ppt', 'pptx', 'pps', 'ppsx', 'pdf',
  'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'ico', 'avif',
];

function getExternalOpenPrefsPath() {
  return path.join(app.getPath('userData'), 'external-open-prefs.json');
}

function loadExternalOpenPrefs() {
  if (externalOpenPrefs) return externalOpenPrefs;
  externalOpenPrefs = { ...DEFAULT_EXTERNAL_OPEN_PREFS };
  try {
    const p = getExternalOpenPrefsPath();
    if (fs.existsSync(p)) {
      const parsed = JSON.parse(fs.readFileSync(p, 'utf8'));
      if (parsed && typeof parsed === 'object') externalOpenPrefs = { ...externalOpenPrefs, ...parsed };
    }
  } catch (_) {}
  return externalOpenPrefs;
}

function saveExternalOpenPrefs() {
  try {
    fs.mkdirSync(path.dirname(getExternalOpenPrefsPath()), { recursive: true });
    fs.writeFileSync(getExternalOpenPrefsPath(), JSON.stringify(loadExternalOpenPrefs(), null, 2), 'utf8');
  } catch (_) {}
}

function setExternalOpenPref(extKey, checked) {
  const prefs = loadExternalOpenPrefs();
  prefs[extKey] = !!checked;
  saveExternalOpenPrefs();
  createAppMenu();
}

function openHelpDocument() {
  if (!HELP_FILE) return;
  const docPath = path.join(__dirname, 'docs', HELP_FILE);
  if (!fs.existsSync(docPath)) return;
  const ext = path.extname(docPath).toLowerCase();
  if (ext === '.html' || ext === '.htm') {
    if (helpWindow && !helpWindow.isDestroyed()) {
      helpWindow.focus();
      return;
    }
    helpWindow = new BrowserWindow({
      width: 880,
      height: 720,
      minWidth: 400,
      minHeight: 320,
      title: 'Help',
      resizable: true,
      maximizable: true,
      minimizable: true,
      fullscreenable: true,
      autoHideMenuBar: false,
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: false,
        preload: preloadPath,
      },
    });
    helpWindow.loadFile(docPath);
    helpWindow.on('closed', () => { helpWindow = null; });
  } else {
    shell.openPath(docPath);
  }
}

function createAppMenu() {
  const isMac = process.platform === 'darwin';
  const prefs = loadExternalOpenPrefs();
  const sendMenuCommand = (browserWindow, command) => {
    const targetWindow = browserWindow || firstMainWindow;
    if (!targetWindow || targetWindow.isDestroyed()) return;
    targetWindow.webContents.send('web2electron:menu-command', command);
  };
  const externalTypeItems = [
    ['pdf', 'PDF'],
    ['doc', 'DOC'],
    ['docx', 'DOCX'],
    ['csv', 'CSV'],
    ['xls', 'XLS'],
    ['xlsx', 'XLSX'],
    ['ppt', 'PPT'],
    ['pptx', 'PPTX'],
    ['hwp', 'HWP'],
    ['hwpx', 'HWPX'],
  ].map(([key, label]) => ({
    label,
    type: 'checkbox',
    checked: !!prefs[key],
    click: (menuItem) => setExternalOpenPref(key, menuItem.checked),
  }));
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'File',
      submenu: [
        {
          label: '새 파일',
          accelerator: 'Ctrl+N',
          click: (_menuItem, browserWindow) => sendMenuCommand(browserWindow, 'new-file'),
        },
        {
          label: '파일 열기...',
          accelerator: 'Ctrl+O',
          click: async (_menuItem, browserWindow) => {
            const targetWindow = browserWindow || firstMainWindow;
            if (!targetWindow || targetWindow.isDestroyed()) return;
            const result = await dialog.showOpenDialog(targetWindow, {
              title: '파일 열기',
              properties: ['openFile'],
              filters: [
                { name: 'Supported documents', extensions: OPEN_DIALOG_EXTENSIONS },
                { name: 'All files', extensions: ['*'] },
              ],
            });
            if (result.canceled || !result.filePaths || !result.filePaths[0]) return;
            const filePath = path.resolve(result.filePaths[0]);
            await openSupportedFileFromDialog(targetWindow, filePath);
          },
        },
        { type: 'separator' },
        {
          label: '저장',
          accelerator: 'Ctrl+S',
          click: (_menuItem, browserWindow) => sendMenuCommand(browserWindow, 'save-file'),
        },
        {
          label: '다른 이름으로 저장...',
          accelerator: 'Ctrl+Shift+S',
          click: (_menuItem, browserWindow) => sendMenuCommand(browserWindow, 'save-file-as'),
        },
        { type: 'separator' },
        {
          label: '외부앱으로 파일 열기...',
          click: async (_menuItem, browserWindow) => {
            const targetWindow = browserWindow || firstMainWindow;
            if (!targetWindow || targetWindow.isDestroyed()) return;
            await openFileWithExternalDialog(targetWindow);
          },
        },
        {
          label: '외부앱으로 열 파일 형식',
          submenu: externalTypeItems,
        },
        {
          label: 'PPTX 열기 시 매번 새창으로 띄우기',
          type: 'checkbox',
          checked: !!prefs.alwaysNewWindowPptx,
          click: (menuItem) => {
            const currentPrefs = loadExternalOpenPrefs();
            currentPrefs.alwaysNewWindowPptx = menuItem.checked;
            saveExternalOpenPrefs();
            createAppMenu();
          },
        },
        {
          label: 'LibreOffice 다운로드...',
          click: () => shell.openExternal(LIBREOFFICE_DOWNLOAD_URL),
        },
        { type: 'separator' },
        {
          label: '폴더 열기...    Ctrl+K Ctrl+O',
          click: async (_menuItem, browserWindow) => {
            const targetWindow = browserWindow || firstMainWindow;
            if (!targetWindow || targetWindow.isDestroyed()) return;
            const result = await dialog.showOpenDialog(targetWindow, {
              title: '작업 폴더 열기',
              properties: ['openDirectory'],
            });
            if (result.canceled || !result.filePaths || !result.filePaths[0]) return;
            const folderPath = path.resolve(result.filePaths[0]);
            targetWindow.webContents.send('web2electron:opened-folder', {
              path: folderPath,
              parentPath: getParentDirectory(folderPath),
              name: getNameFromPath(folderPath),
            });
          },
        },
        { type: 'separator' },
        isMac
          ? { label: 'Close', role: 'close' }
          : {
              label: '종료 (Exit)',
              accelerator: 'Ctrl+Q',
              click: () => app.quit(),
            },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'forceReload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        {
          label: '한 페이지 스크롤',
          type: 'radio',
          checked: true,
          click: (menuItem, browserWindow) => {
            if (browserWindow) {
              browserWindow.webContents.send('set-view-mode', 'single');
            }
          }
        },
        {
          label: '두 페이지 보기',
          type: 'radio',
          click: (menuItem, browserWindow) => {
            if (browserWindow) {
              browserWindow.webContents.send('set-view-mode', '2-page');
            }
          }
        },
        {
          label: '매트릭스 보기',
          type: 'radio',
          click: (menuItem, browserWindow) => {
            if (browserWindow) {
              browserWindow.webContents.send('set-view-mode', 'matrix');
            }
          }
        },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
      ],
    },
    {
      label: 'Window',
      submenu: [
        { role: 'minimize' },
        ...(isMac ? [{ role: 'zoom' }, { type: 'separator' }, { role: 'front' }] : []),
        { role: 'close' },
      ],
    },
  ];
  template.push({
    label: 'Help',
    submenu: [
      {
        label: HELP_FILE ? '도움말 문서' : '도움말 문서 (미번들)',
        click: () => openHelpDocument(),
        enabled: Boolean(HELP_FILE),
      },
    ],
  });
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

let firstMainWindow = null;
let pendingOpenPath = null;
let ipcInitialPath = null;
let appLaunchHandled = false;

const MAX_OPEN_FILE_BYTES = 32 * 1024 * 1024;
const MAX_WORKSPACE_FILE_BYTES = 32 * 1024 * 1024;
const MAX_WORKSPACE_ENTRIES = 2000;
const WORKSPACE_TEXT_EXTS = new Set([
  '.md', '.markdown', '.mdown', '.txt', '.csv', '.html', '.htm', '.json', '.mpv', '.mpp', '.mdd', '.docx'
]);
const WORKSPACE_VIEW_EXTS = new Set([
  '.pdf',
  '.doc', '.docx',
  '.hwp', '.hwpx',
  '.xls', '.xlsx',
  '.ppt', '.pptx', '.pps', '.ppsx',
  '.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.svg', '.ico', '.avif'
]);
const WORKSPACE_LIST_EXTS = new Set([
  ...WORKSPACE_TEXT_EXTS,
  ...WORKSPACE_VIEW_EXTS,
]);

function tryReadFileTextUtf8(p) {
  if (!p) return null;
  try {
    if (!fs.existsSync(p)) return null;
    const st = fs.statSync(p);
    if (!st.isFile() || st.size > MAX_OPEN_FILE_BYTES) return null;
    return fs.readFileSync(p, 'utf8');
  } catch (_) {
    return null;
  }
}

function getNameFromPath(p) {
  try {
    return path.basename(String(p || ''));
  } catch (_) {
    return '';
  }
}

function normalizeExistingDirectory(p) {
  if (!p) return null;
  try {
    const resolved = path.resolve(String(p));
    if (fs.existsSync(resolved) && fs.statSync(resolved).isDirectory()) return resolved;
  } catch (_) {}
  return null;
}

function isPathInside(parent, child) {
  try {
    const rel = path.relative(parent, child);
    return rel === '' || (!!rel && !rel.startsWith('..') && !path.isAbsolute(rel));
  } catch (_) {
    return false;
  }
}

function getParentDirectory(p) {
  try {
    const resolved = normalizeExistingDirectory(p);
    if (!resolved) return null;
    const parent = path.dirname(resolved);
    if (parent && parent !== resolved && fs.existsSync(parent) && fs.statSync(parent).isDirectory()) return parent;
  } catch (_) {}
  return null;
}

function listWorkspaceDirectory(rootPath, dirPath) {
  const root = normalizeExistingDirectory(rootPath);
  const dir = normalizeExistingDirectory(dirPath || rootPath);
  if (!root || !dir || !isPathInside(root, dir)) return { error: 'Invalid workspace path.' };

  try {
    const entries = fs.readdirSync(dir, { withFileTypes: true }).slice(0, MAX_WORKSPACE_ENTRIES);
    const items = entries
      .filter((entry) => {
        const name = entry.name || '';
        if (!name || name === '.' || name === '..') return false;
        if (name === 'node_modules' || name === '.git' || name === 'dist') return false;
        if (entry.isDirectory()) return true;
        return WORKSPACE_LIST_EXTS.has(path.extname(name).toLowerCase());
      })
      .map((entry) => {
        const fullPath = path.join(dir, entry.name);
        const isDirectory = entry.isDirectory();
        let size = 0;
        let mtimeMs = 0;
        try {
          const st = fs.statSync(fullPath);
          size = st.size || 0;
          mtimeMs = st.mtimeMs || 0;
        } catch (_) {}
        return {
          name: entry.name,
          path: fullPath,
          relativePath: path.relative(root, fullPath),
          type: isDirectory ? 'directory' : 'file',
          kind: isDirectory ? 'directory' : (WORKSPACE_VIEW_EXTS.has(path.extname(entry.name).toLowerCase()) ? 'viewer' : 'text'),
          size,
          mtimeMs,
        };
      })
      .sort((a, b) => {
        if (a.type !== b.type) return a.type === 'directory' ? -1 : 1;
        return String(a.name).localeCompare(String(b.name), 'ko', { numeric: true, sensitivity: 'base' });
      });
    return { rootPath: root, dirPath: dir, parentPath: getParentDirectory(dir), items };
  } catch (err) {
    return { error: err && err.message ? err.message : String(err) };
  }
}

function getWorkspaceBoundary(rootPath, displayRootPath) {
  const root = normalizeExistingDirectory(rootPath);
  if (!root) return null;
  const displayRoot = normalizeExistingDirectory(displayRootPath);
  if (displayRoot && isPathInside(displayRoot, root)) return displayRoot;
  return root;
}

function listWorkspaceDisplayDirectory(rootPath, displayRootPath, dirPath) {
  const boundary = getWorkspaceBoundary(rootPath, displayRootPath);
  const dir = normalizeExistingDirectory(dirPath || boundary);
  if (!boundary || !dir || !isPathInside(boundary, dir)) return { error: 'Invalid workspace path.' };
  return listWorkspaceDirectory(boundary, dir);
}

async function convertDocxToHtmlContent(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  const converted = await mammoth.convertToHtml(
    { path: resolved },
    {
      convertImage: mammoth.images.imgElement(async (image) => {
        const buffer = await image.read('base64');
        return { src: 'data:' + image.contentType + ';base64,' + buffer };
      }),
    }
  );
  const warnings = Array.isArray(converted.messages) && converted.messages.length
    ? '\n\n<hr>\n\n<details><summary>DOCX 변환 알림</summary>\n\n```\n' + converted.messages.map((m) => m.message || String(m)).join('\n') + '\n```\n\n</details>\n'
    : '';
  return '<div class="docx-imported-document">\n' + (converted.value || '') + '\n</div>' + warnings;
}

async function readWorkspaceFile(rootPath, filePath) {
  const root = normalizeExistingDirectory(rootPath);
  if (!root || !filePath) return { error: 'Invalid workspace path.' };
  try {
    const resolved = path.resolve(String(filePath));
    if (!isPathInside(root, resolved) || !fs.existsSync(resolved)) return { error: 'File is outside workspace.' };
    const st = fs.statSync(resolved);
    if (!st.isFile()) return { error: 'Not a file.' };
    if (st.size > MAX_WORKSPACE_FILE_BYTES) return { error: 'File is too large.' };
    const ext = path.extname(resolved).toLowerCase();
    if (!WORKSPACE_TEXT_EXTS.has(ext)) return { error: 'Unsupported file type.' };
    if (ext === '.docx') {
      return {
        path: resolved,
        fileName: getNameFromPath(resolved).replace(/\.docx$/i, '.md'),
        originalFileName: getNameFromPath(resolved),
        sourceFormat: 'docx',
        text: await convertDocxToHtmlContent(resolved),
      };
    }
    return {
      path: resolved,
      fileName: getNameFromPath(resolved),
      text: fs.readFileSync(resolved, 'utf8'),
    };
  } catch (err) {
    return { error: err && err.message ? err.message : String(err) };
  }
}

async function readDirectSupportedTextFile(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  if (!fs.existsSync(resolved)) return { error: 'File does not exist.' };
  const st = fs.statSync(resolved);
  if (!st.isFile()) return { error: 'Not a file.' };
  if (st.size > MAX_WORKSPACE_FILE_BYTES) return { error: 'File is too large.' };
  const ext = path.extname(resolved).toLowerCase();
  if (!WORKSPACE_TEXT_EXTS.has(ext)) return { error: 'Unsupported file type.' };
  if (ext === '.docx') {
    return {
      path: resolved,
      fileName: getNameFromPath(resolved).replace(/\.docx$/i, '.md'),
      originalFileName: getNameFromPath(resolved),
      sourceFormat: 'docx',
      text: await convertDocxToHtmlContent(resolved),
    };
  }
  return {
    path: resolved,
    fileName: getNameFromPath(resolved),
    text: fs.readFileSync(resolved, 'utf8'),
  };
}

async function openSupportedFileFromDialog(browserWindow, filePath) {
  const resolved = path.resolve(String(filePath || ''));
  const ext = path.extname(resolved).toLowerCase();
  if (WORKSPACE_TEXT_EXTS.has(ext)) {
    const payload = await readDirectSupportedTextFile(resolved);
    if (payload && !payload.error) {
      browserWindow.webContents.send('web2electron:opened-file', payload.path, payload.text, payload);
      return payload;
    }
    dialog.showErrorBox('파일 열기 실패', payload && payload.error ? payload.error : 'Unknown error');
    return payload;
  }
  if (WORKSPACE_VIEW_EXTS.has(ext)) {
    return openReadonlyFilePath(resolved, browserWindow);
  }
  dialog.showErrorBox('지원하지 않는 파일', '이 파일 형식은 MDpro에서 열 수 없습니다.');
  return { error: 'Unsupported file type.' };
}

function createPdfViewerWindow(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  const viewerWindow = new BrowserWindow({
    width: 1100,
    height: 820,
    minWidth: 520,
    minHeight: 360,
    title: getNameFromPath(resolved),
    resizable: true,
    maximizable: true,
    minimizable: true,
    fullscreenable: true,
    autoHideMenuBar: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: preloadPath,
    },
  });
  viewerWindow.loadURL(pathToFileURL(resolved).toString());
  return viewerWindow;
}

function createImageViewerWindow(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  const viewerWindow = new BrowserWindow({
    width: 1100,
    height: 820,
    minWidth: 360,
    minHeight: 260,
    title: getNameFromPath(resolved),
    resizable: true,
    maximizable: true,
    minimizable: true,
    fullscreenable: true,
    autoHideMenuBar: false,
    backgroundColor: '#111827',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: preloadPath,
    },
  });
  const imageUrl = pathToFileURL(resolved).toString();
  const html = '<!doctype html><html><head><meta charset="utf-8"><title>' + escapeHtmlDocumentText(getNameFromPath(resolved)) + '</title>'
    + '<style>html,body{margin:0;width:100%;height:100%;background:#111827;color:#e5e7eb;font-family:"Malgun Gothic",Arial,sans-serif;overflow:hidden;}'
    + '.bar{position:fixed;top:0;left:0;right:0;height:38px;display:flex;align-items:center;padding:0 12px;background:rgba(17,24,39,.88);border-bottom:1px solid rgba(148,163,184,.25);font-size:12px;z-index:2;}'
    + '.name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}'
    + '.stage{box-sizing:border-box;width:100%;height:100%;padding:50px 16px 16px;display:flex;align-items:center;justify-content:center;}'
    + 'img{max-width:100%;max-height:100%;object-fit:contain;box-shadow:0 18px 60px rgba(0,0,0,.35);}</style></head><body>'
    + '<div class="bar"><div class="name">' + escapeHtmlDocumentText(getNameFromPath(resolved)) + '</div></div>'
    + '<div class="stage"><img src="' + imageUrl + '" alt=""></div></body></html>';
  viewerWindow.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  return viewerWindow;
}

function createGenSlideImageHtml(imagePath, index, total) {
  const imageUrl = pathToFileURL(imagePath).toString();
  const title = 'Slide ' + index + ' / ' + total;
  return '<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>html,body{margin:0;width:100%;height:100%;background:#0f172a;}'
    + '.slide{box-sizing:border-box;width:1280px;height:720px;margin:0;background:#0f172a;position:relative;overflow:hidden;}'
    + '.slide img{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#0f172a;}'
    + '.slide .label{position:absolute;right:18px;bottom:14px;padding:4px 8px;border-radius:4px;background:rgba(15,23,42,.65);color:#e5e7eb;font:12px Segoe UI,Arial,sans-serif;}</style></head>'
    + '<body><div class="slide"><img src="' + imageUrl + '" alt=""><div class="label">' + escapeHtmlDocumentText(title) + '</div></div></body></html>';
}

function createGenSlideWindowWithSlides(slideItems, title) {
  const viewerWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 720,
    minHeight: 480,
    title: title || 'GenSlide',
    resizable: true,
    maximizable: true,
    minimizable: true,
    fullscreenable: true,
    autoHideMenuBar: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: preloadPath,
    },
  });
  const genSlidePath = path.join(__dirname, 'web', 'js', 'GenSlide', 'jenaEditor', 'index.html');
  viewerWindow.loadFile(genSlidePath);
  viewerWindow.webContents.once('did-finish-load', () => {
    const payload = JSON.stringify({
      title: title || 'PPTX Image Slides',
      slides: Array.isArray(slideItems) ? slideItems : [],
    });
    const script = `
      (function () {
        var payload = ${payload};
        function apply() {
          if (!Array.isArray(payload.slides) || !payload.slides.length || typeof loadCurrent !== 'function') return false;
          slides = payload.slides.map(function (s) { return { html: String((s && s.html) || '') }; }).filter(function (s) { return s.html.trim(); });
          if (!slides.length) return false;
          cur = 0;
          slideWidth = 1280;
          slideHeight = 720;
          backup = slides[0].html;
          loadCurrent();
          if (typeof renderSidebar === 'function') renderSidebar();
          if (typeof setWorkViewMode === 'function') setWorkViewMode('editor-only');
          return true;
        }
        if (!apply()) setTimeout(apply, 800);
      })();
    `;
    viewerWindow.webContents.executeJavaScript(script).catch(() => {});
  });
  return viewerWindow;
}

function createPptxHtmlViewerWindow(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  const prefs = loadExternalOpenPrefs();

  if (!prefs.alwaysNewWindowPptx && activePptxWindow && !activePptxWindow.isDestroyed()) {
    try {
      activePptxWindow.loadFile(path.join(__dirname, 'web', 'pptx-viewer.html'), {
        query: {
          file: pathToFileURL(resolved).toString(),
          path: resolved,
          title: getNameFromPath(resolved),
        },
      });
      if (activePptxWindow.isMinimized()) activePptxWindow.restore();
      activePptxWindow.focus();
      return activePptxWindow;
    } catch (_) {}
  }

  const viewerWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 720,
    minHeight: 480,
    title: getNameFromPath(resolved),
    resizable: true,
    maximizable: true,
    minimizable: true,
    fullscreenable: true,
    autoHideMenuBar: false,
    backgroundColor: '#0f172a',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: false,
      preload: preloadPath,
    },
  });

  if (!prefs.alwaysNewWindowPptx) {
    activePptxWindow = viewerWindow;
  }

  viewerWindow.loadFile(path.join(__dirname, 'web', 'pptx-viewer.html'), {
    query: {
      file: pathToFileURL(resolved).toString(),
      path: resolved,
      title: getNameFromPath(resolved),
    },
  });
  return viewerWindow;
}

function escapeHtmlDocumentText(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function createDocxViewerWindow(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  const docxHtml = await convertDocxToHtmlContent(resolved);
  const html = '<!doctype html><html><head><meta charset="utf-8">'
    + '<title>' + escapeHtmlDocumentText(getNameFromPath(resolved)) + '</title>'
    + '<style>'
    + 'body{margin:0;background:#eef2f7;color:#111827;font-family:"Malgun Gothic",Arial,sans-serif;}'
    + '.toolbar{position:sticky;top:0;z-index:10;display:flex;align-items:center;gap:10px;padding:9px 14px;background:#fff;border-bottom:1px solid #d8dee8;box-shadow:0 1px 2px rgba(15,23,42,.06);}'
    + '.title{font-size:13px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}'
    + '.page{box-sizing:border-box;width:min(920px,calc(100vw - 32px));min-height:calc(100vh - 88px);margin:16px auto;padding:54px 64px;background:#fff;border:1px solid #d8dee8;box-shadow:0 10px 30px rgba(15,23,42,.12);line-height:1.65;font-size:15px;}'
    + '.page img{max-width:100%;height:auto;} .page table{border-collapse:collapse;max-width:100%;} .page td,.page th{border:1px solid #cbd5e1;padding:4px 7px;}'
    + '.docx-messages{width:min(920px,calc(100vw - 32px));margin:0 auto 16px;color:#64748b;font-size:12px;} pre{white-space:pre-wrap;}'
    + '@media print{body{background:#fff}.toolbar,.docx-messages{display:none}.page{width:auto;min-height:auto;margin:0;padding:0;border:0;box-shadow:none;}}'
    + '</style></head><body>'
    + '<div class="toolbar"><div class="title">' + escapeHtmlDocumentText(getNameFromPath(resolved)) + '</div></div>'
    + '<article class="page">' + docxHtml + '</article>'
    + '</body></html>';

  const viewerWindow = new BrowserWindow({
    width: 1100,
    height: 820,
    minWidth: 520,
    minHeight: 360,
    title: getNameFromPath(resolved),
    resizable: true,
    maximizable: true,
    minimizable: true,
    fullscreenable: true,
    autoHideMenuBar: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: preloadPath,
    },
  });
  viewerWindow.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  return viewerWindow;
}

function findLibreOfficeExecutable() {
  const candidates = [];
  if (process.platform === 'win32') {
    candidates.push(
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
      path.join(__dirname, 'LibreOfficePortable', 'App', 'libreoffice', 'program', 'soffice.exe'),
      path.join(__dirname, 'LibreOfficePortable', 'LibreOfficePortable.exe')
    );
  } else if (process.platform === 'darwin') {
    candidates.push('/Applications/LibreOffice.app/Contents/MacOS/soffice');
  } else {
    candidates.push('/usr/bin/libreoffice', '/usr/bin/soffice');
  }
  return candidates.find((candidate) => {
    try {
      return fs.existsSync(candidate);
    } catch (_) {
      return false;
    }
  }) || null;
}

function convertOfficeFileToPdf(filePath) {
  return new Promise((resolve) => {
    const soffice = findLibreOfficeExecutable();
    if (!soffice) return resolve({ error: 'LibreOffice is not installed.' });

    const source = path.resolve(String(filePath || ''));
    const outDir = path.join(os.tmpdir(), 'mdpro-libreoffice-viewer-' + Date.now());
    fs.mkdirSync(outDir, { recursive: true });
    execFile(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', outDir, source], { windowsHide: true }, (err) => {
      if (err) return resolve({ error: err.message || String(err) });
      const base = path.basename(source, path.extname(source)) + '.pdf';
      const pdfPath = path.join(outDir, base);
      if (!fs.existsSync(pdfPath)) return resolve({ error: 'LibreOffice did not create a PDF preview.' });
      resolve({ pdfPath });
    });
  });
}

function convertMicrosoftOfficeFileToPdf(filePath) {
  return new Promise((resolve) => {
    if (process.platform !== 'win32') return resolve({ error: 'MS Office conversion is only available on Windows.' });

    const source = path.resolve(String(filePath || ''));
    const ext = path.extname(source).toLowerCase();
    const outDir = path.join(os.tmpdir(), 'mdpro-msoffice-viewer-' + Date.now());
    fs.mkdirSync(outDir, { recursive: true });
    const pdfPath = path.join(outDir, path.basename(source, path.extname(source)) + '.pdf');
    let appType = '';
    if (ext === '.doc' || ext === '.docx') appType = 'word';
    else if (ext === '.ppt' || ext === '.pptx' || ext === '.pps' || ext === '.ppsx') appType = 'powerpoint';
    else if (ext === '.xls' || ext === '.xlsx') appType = 'excel';
    else return resolve({ error: 'Unsupported MS Office conversion type.' });

    const script = [
      'param([string]$Source,[string]$Pdf,[string]$Kind)',
      '$ErrorActionPreference = "Stop"',
      '$app = $null; $doc = $null',
      'try {',
      '  if ($Kind -eq "word") {',
      '    $app = New-Object -ComObject Word.Application; $app.Visible = $false',
      '    $doc = $app.Documents.Open($Source, $false, $true)',
      '    $doc.ExportAsFixedFormat($Pdf, 17)',
      '  } elseif ($Kind -eq "powerpoint") {',
      '    $app = New-Object -ComObject PowerPoint.Application',
      '    $doc = $app.Presentations.Open($Source, $true, $true, $false)',
      '    $doc.SaveAs($Pdf, 32)',
      '  } elseif ($Kind -eq "excel") {',
      '    $app = New-Object -ComObject Excel.Application; $app.Visible = $false; $app.DisplayAlerts = $false',
      '    $doc = $app.Workbooks.Open($Source, 3, $true)',
      '    $doc.ExportAsFixedFormat(0, $Pdf)',
      '  }',
      '  Write-Output $Pdf',
      '} finally {',
      '  if ($doc -ne $null) { try { $doc.Close($false) } catch {} }',
      '  if ($app -ne $null) { try { $app.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {} }',
      '}',
    ].join('\n');

    execFile('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script, '-Source', source, '-Pdf', pdfPath, '-Kind', appType], { windowsHide: true, timeout: 120000 }, (err) => {
      if (err) return resolve({ error: err.message || String(err) });
      if (!fs.existsSync(pdfPath)) return resolve({ error: 'MS Office did not create a PDF preview.' });
      resolve({ pdfPath });
    });
  });
}

function convertPowerPointFileToPngSlides(filePath) {
  return new Promise((resolve) => {
    if (process.platform !== 'win32') return resolve({ error: 'PowerPoint export is only available on Windows.' });
    const source = path.resolve(String(filePath || ''));
    const ext = path.extname(source).toLowerCase();
    if (!['.ppt', '.pptx', '.pps', '.ppsx'].includes(ext)) return resolve({ error: 'Not a PowerPoint file.' });
    const outDir = path.join(os.tmpdir(), 'mdpro-genslide-pptx-' + Date.now());
    fs.mkdirSync(outDir, { recursive: true });
    const script = [
      'param([string]$Source,[string]$OutDir)',
      '$ErrorActionPreference = "Stop"',
      '$app = $null; $pres = $null',
      'try {',
      '  $app = New-Object -ComObject PowerPoint.Application',
      '  $pres = $app.Presentations.Open($Source, $true, $true, $false)',
      '  $pres.Export($OutDir, "PNG", 1280, 720)',
      '  Write-Output $OutDir',
      '} finally {',
      '  if ($pres -ne $null) { try { $pres.Close() } catch {} }',
      '  if ($app -ne $null) { try { $app.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {} }',
      '}',
    ].join('\n');
    execFile('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script, '-Source', source, '-OutDir', outDir], { windowsHide: true, timeout: 180000 }, (err) => {
      if (err) return resolve({ error: err.message || String(err) });
      let images = [];
      try {
        images = fs.readdirSync(outDir)
          .filter((name) => /\.(png|jpg|jpeg)$/i.test(name))
          .map((name) => path.join(outDir, name))
          .sort((a, b) => String(path.basename(a)).localeCompare(String(path.basename(b)), undefined, { numeric: true, sensitivity: 'base' }));
      } catch (_) {
        images = [];
      }
      if (!images.length) return resolve({ error: 'PowerPoint did not export slide images.' });
      resolve({ imagePaths: images, outDir });
    });
  });
}

async function openPowerPointInGenSlide(win, filePath) {
  const converted = await convertPowerPointFileToPngSlides(filePath);
  if (!converted || !Array.isArray(converted.imagePaths) || !converted.imagePaths.length) {
    return openWithOfficePreviewOrGuide(win, filePath);
  }
  const total = converted.imagePaths.length;
  const slidesForGenSlide = converted.imagePaths.map((imagePath, index) => ({
    html: createGenSlideImageHtml(imagePath, index + 1, total),
  }));
  createGenSlideWindowWithSlides(slidesForGenSlide, getNameFromPath(filePath) + ' - GenSlide');
  return { opened: true, mode: 'genslide-image-slides', filePath, fileName: getNameFromPath(filePath), slideCount: total };
}

async function openPowerPointInThumbnailViewer(filePath) {
  const resolved = path.resolve(String(filePath || ''));
  createPptxHtmlViewerWindow(resolved);
  return { opened: true, mode: 'pptx-thumbnail', filePath: resolved, fileName: getNameFromPath(resolved) };
}

async function openReadonlyFilePath(filePath, win) {
  const resolved = path.resolve(String(filePath || ''));
  if (!fs.existsSync(resolved) || !fs.statSync(resolved).isFile()) return { error: 'Not a file.' };
  const ext = path.extname(resolved).toLowerCase();
  if (!WORKSPACE_VIEW_EXTS.has(ext)) return { error: 'Unsupported viewer file type.' };
  const extKey = ext.replace(/^\./, '');
  const prefs = loadExternalOpenPrefs();

  if (ext === '.hwp' || ext === '.hwpx') {
    shell.openExternal(RHWP_URL);
    return { opened: true, mode: 'rhwp', filePath: resolved, fileName: getNameFromPath(resolved) };
  }

  if (ext === '.pdf') {
    createPdfViewerWindow(resolved);
    return { opened: true, mode: 'internal', filePath: resolved, fileName: getNameFromPath(resolved) };
  }

  if (['.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.svg', '.ico', '.avif'].includes(ext)) {
    createImageViewerWindow(resolved);
    return { opened: true, mode: 'image-internal', filePath: resolved, fileName: getNameFromPath(resolved) };
  }

  if (ext === '.docx') {
    await createDocxViewerWindow(resolved);
    return { opened: true, mode: 'docx-internal', filePath: resolved, fileName: getNameFromPath(resolved) };
  }

  if (ext === '.doc') return openWithOfficePreviewOrGuide(win, resolved);
  if (ext === '.ppt' || ext === '.pptx' || ext === '.pps' || ext === '.ppsx') return openPowerPointInThumbnailViewer(resolved);
  if (prefs[extKey]) return openExternalFilePath(resolved, win);
  return openExternalFilePath(resolved, win);
}

async function openWithOfficePreviewOrGuide(win, filePath) {
  const msConverted = await convertMicrosoftOfficeFileToPdf(filePath);
  if (msConverted && msConverted.pdfPath) {
    createPdfViewerWindow(msConverted.pdfPath);
    return { opened: true, mode: 'msoffice-pdf', filePath, fileName: getNameFromPath(filePath) };
  }

  const converted = await convertOfficeFileToPdf(filePath);
  if (converted && converted.pdfPath) {
    createPdfViewerWindow(converted.pdfPath);
    return { opened: true, mode: 'libreoffice-pdf', filePath, fileName: getNameFromPath(filePath) };
  }

  const result = await dialog.showMessageBox(win || firstMainWindow || undefined, {
    type: 'info',
    buttons: ['LibreOffice 다운로드', 'MS Office/외부앱으로 열기', '취소'],
    defaultId: 0,
    cancelId: 2,
    title: '문서 변환기가 필요합니다',
    message: '이 문서를 MDpro 안에서 PDF 미리보기로 변환하지 못했습니다.',
    detail: 'MS Office가 설치되어 있으면 Word/PowerPoint/Excel을 사용해 PDF로 변환합니다. 실패한 경우 외부앱으로 직접 열 수 있습니다.\nLibreOffice가 설치되어 있으면 MDpro가 대체 변환기로 사용합니다.\n설치 주소: ' + LIBREOFFICE_DOWNLOAD_URL,
  });
  if (result.response === 0) {
    shell.openExternal(LIBREOFFICE_DOWNLOAD_URL);
    return { opened: true, mode: 'libreoffice-download' };
  }
  if (result.response === 1) {
    const opened = await shell.openPath(filePath);
    return opened ? { error: opened } : { opened: true, mode: 'external', filePath, fileName: getNameFromPath(filePath) };
  }
  return { canceled: true };
}

async function openExternalFilePath(filePath, win) {
  const resolved = path.resolve(String(filePath || ''));
  if (!fs.existsSync(resolved) || !fs.statSync(resolved).isFile()) return { error: 'Not a file.' };
  const opened = await shell.openPath(resolved);
  if (!opened) return { opened: true, mode: 'external', filePath: resolved, fileName: getNameFromPath(resolved) };
  return openWithOfficePreviewOrGuide(win, resolved);
}

async function openFileWithExternalDialog(win) {
  const result = await dialog.showOpenDialog(win || undefined, {
    title: '외부앱으로 파일 열기',
    properties: ['openFile'],
    filters: [
      { name: 'Documents', extensions: ['pdf', 'doc', 'docx', 'hwp', 'hwpx', 'csv', 'xls', 'xlsx', 'ppt', 'pptx', 'pps', 'ppsx'] },
      { name: 'All files', extensions: ['*'] },
    ],
  });
  if (result.canceled || !result.filePaths || !result.filePaths[0]) return { canceled: true };
  return openExternalFilePath(result.filePaths[0], win);
}

async function openWorkspaceReadonlyFile(rootPath, filePath, senderWebContents) {
  const root = normalizeExistingDirectory(rootPath);
  if (!root || !filePath) return { error: 'Invalid workspace path.' };
  try {
    const resolved = path.resolve(String(filePath));
    if (!isPathInside(root, resolved) || !fs.existsSync(resolved)) return { error: 'File is outside workspace.' };
    const st = fs.statSync(resolved);
    if (!st.isFile()) return { error: 'Not a file.' };

    const ext = path.extname(resolved).toLowerCase();
    if (!WORKSPACE_VIEW_EXTS.has(ext)) return { error: 'Unsupported viewer file type.' };
    const win = senderWebContents ? BrowserWindow.fromWebContents(senderWebContents) : firstMainWindow;
    return await openReadonlyFilePath(resolved, win);
  } catch (err) {
    return { error: err && err.message ? err.message : String(err) };
  }
}

function sanitizeSaveFileName(fileName) {
  const name = String(fileName || 'document.md').trim().replace(/[\\/:*?"<>|]+/g, '_') || 'document.md';
  return /\.[a-z0-9]+$/i.test(name) ? name : name + '.md';
}

function writeTextFileUtf8(filePath, content) {
  const resolved = path.resolve(String(filePath || ''));
  if (!resolved) return { error: 'Invalid file path.' };
  fs.writeFileSync(resolved, String(content == null ? '' : content), 'utf8');
  return {
    filePath: resolved,
    fileName: getNameFromPath(resolved),
  };
}

/** Windows: argv에 Chromium 플래그가 섞임. -- 구분 이후 인자만 우선 검사(연결 프로그램으로 연 .md 등). */
function extractFileArg(argv, cwd) {
  const base = cwd && String(cwd).trim() ? String(cwd) : process.cwd();
  if (!argv || argv.length < 2) return null;
  const skipToken = (a) =>
    !a || a === '.' || a === '..' || a === '--' || a.startsWith('-');
  const tryFile = (arg) => {
    if (skipToken(arg)) return null;
    try {
      const p = path.normalize(path.isAbsolute(arg) ? arg : path.resolve(base, arg));
      if (fs.existsSync(p) && fs.statSync(p).isFile()) return p;
    } catch (_) {}
    return null;
  };
  const scan = (list) => {
    let last = null;
    for (let i = 0; i < list.length; i++) {
      const f = tryFile(list[i]);
      if (f) last = f;
    }
    return last;
  };
  const dash = argv.indexOf('--');
  if (dash >= 0) {
    const after = scan(argv.slice(dash + 1));
    if (after) return after;
  }
  return scan(argv.slice(1));
}

function createWindow() {
  const isFirstMain = firstMainWindow === null;
  let openPath = null;
  if (isFirstMain) {
    if (!appLaunchHandled) {
      openPath = extractFileArg(process.argv, process.cwd()) || pendingOpenPath;
      pendingOpenPath = null;
      appLaunchHandled = true;
    } else {
      pendingOpenPath = null;
    }
  }
  ipcInitialPath = openPath;

  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    resizable: true,
    frame: true,
    title: "MDpro",
    autoHideMenuBar: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: preloadPath,
    },
  });

  // Windows/Linux: 메인 창을 닫을 때 PV 등 자식 창이 남으면 프로세스가 안 끝나는 경우가 있어 함께 닫음
  if (firstMainWindow === null) {
    firstMainWindow = win;
    win.on('closed', () => { firstMainWindow = null; });
    if (process.platform !== 'darwin') {
      win.on('close', () => {
        BrowserWindow.getAllWindows().forEach((w) => {
          try {
            if (w !== win && !w.isDestroyed()) w.destroy();
          } catch (_) {}
        });
      });
    }
  }

  // PV/미리보기 등 window.open('about:blank')·내부 URL은 앱 안 새 창, http(s)만 외부 브라우저
  win.webContents.setWindowOpenHandler((details) => {
    const r = (details.url || '').trim();
    if (/^https?:\/\//i.test(r)) {
      shell.openExternal(r);
      return { action: 'deny' };
    }
    if (/^(mailto|tel):/i.test(r)) {
      shell.openExternal(r);
      return { action: 'deny' };
    }
    return {
      action: 'allow',
      overrideBrowserWindowOptions: {
        width: 1200,
        height: 800,
        minWidth: 320,
        minHeight: 240,
        resizable: true,
        maximizable: true,
        minimizable: true,
        fullscreenable: true,
        autoHideMenuBar: false,
        webPreferences: {
          contextIsolation: true,
          nodeIntegration: false,
          sandbox: false,
          preload: preloadPath,
        },
      },
    };
  });

  win.loadFile(path.join(__dirname, 'web', "index.html"));
  if (openPath) {
    win.webContents.once('did-finish-load', () => {
      try {
        const t = tryReadFileTextUtf8(openPath);
        win.webContents.send('web2electron:opened-file', openPath, t);
      } catch (_) {}
    });
  }
}

app.on('open-file', (event, filePath) => {
  event.preventDefault();
  pendingOpenPath = path.resolve(filePath);
  const w = firstMainWindow;
  if (w && !w.isDestroyed()) {
    const t = tryReadFileTextUtf8(pendingOpenPath);
    w.webContents.send('web2electron:opened-file', pendingOpenPath, t);
    w.focus();
  }
});

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv, workingDirectory) => {
    const p = extractFileArg(argv, workingDirectory || process.cwd());
    if (!p) return;
    const w = firstMainWindow;
    if (w && !w.isDestroyed()) {
      const t = tryReadFileTextUtf8(p);
      w.webContents.send('web2electron:opened-file', p, t);
      w.focus();
    } else {
      pendingOpenPath = p;
    }
  });

  app.whenReady().then(() => {
    ipcMain.handle('web2electron:get-opened-file', () => {
      const ret = ipcInitialPath;
      ipcInitialPath = null;
      if (!ret) return { path: null, text: null };
      return { path: ret, text: tryReadFileTextUtf8(ret) };
    });
    ipcMain.handle('web2electron:open-folder-dialog', async (event) => {
      const win = BrowserWindow.fromWebContents(event.sender) || firstMainWindow;
      const result = await dialog.showOpenDialog(win || undefined, {
        title: '작업 폴더 열기',
        properties: ['openDirectory'],
      });
      if (result.canceled || !result.filePaths || !result.filePaths[0]) return { canceled: true };
      const folderPath = path.resolve(result.filePaths[0]);
      return { canceled: false, path: folderPath, parentPath: getParentDirectory(folderPath), name: getNameFromPath(folderPath) };
    });
    ipcMain.handle('web2electron:list-folder', (_event, payload) => {
      const rootPath = payload && payload.rootPath;
      const displayRootPath = payload && payload.displayRootPath;
      const dirPath = payload && payload.dirPath;
      return listWorkspaceDisplayDirectory(rootPath, displayRootPath, dirPath || displayRootPath || rootPath);
    });
    ipcMain.handle('web2electron:read-workspace-file', async (_event, payload) => {
      const rootPath = payload && payload.rootPath;
      const filePath = payload && payload.filePath;
      return await readWorkspaceFile(rootPath, filePath);
    });
    ipcMain.handle('web2electron:create-directory', (_event, payload) => {
      try {
        const rootPath = payload && payload.rootPath;
        const dirPath = payload && payload.dirPath;
        const folderName = payload && payload.folderName;
        if (!rootPath || !dirPath || !folderName) return { error: 'Invalid parameters.' };
        const root = normalizeExistingDirectory(rootPath);
        if (!root) return { error: 'Invalid workspace root.' };
        const baseDir = path.resolve(dirPath);
        if (!isPathInside(root, baseDir)) return { error: 'Directory is outside workspace.' };
        const cleanFolderName = folderName.replace(/[\\/:*?"<>|]+/g, '_');
        const target = path.join(baseDir, cleanFolderName);
        if (!isPathInside(root, target)) return { error: 'Target is outside workspace.' };
        if (fs.existsSync(target)) {
          return { error: 'Folder already exists.' };
        }
        fs.mkdirSync(target, { recursive: true });
        return { success: true, path: target, name: cleanFolderName };
      } catch (err) {
        return { error: err && err.message ? err.message : String(err) };
      }
    });
    ipcMain.handle('web2electron:create-file', (_event, payload) => {
      try {
        const rootPath = payload && payload.rootPath;
        const dirPath = payload && payload.dirPath;
        const fileName = payload && payload.fileName;
        const content = (payload && payload.content) || '';
        if (!rootPath || !dirPath || !fileName) return { error: 'Invalid parameters.' };
        const root = normalizeExistingDirectory(rootPath);
        if (!root) return { error: 'Invalid workspace root.' };
        const baseDir = path.resolve(dirPath);
        if (!isPathInside(root, baseDir)) return { error: 'Directory is outside workspace.' };
        let cleanFileName = fileName.replace(/[\\/:*?"<>|]+/g, '_');
        if (!/\.[a-z0-9]+$/i.test(cleanFileName)) {
          cleanFileName += '.md';
        }
        const target = path.join(baseDir, cleanFileName);
        if (!isPathInside(root, target)) return { error: 'Target is outside workspace.' };
        if (fs.existsSync(target)) {
          return { error: 'File already exists.' };
        }
        fs.writeFileSync(target, content, 'utf8');
        return { success: true, path: target, name: cleanFileName };
      } catch (err) {
        return { error: err && err.message ? err.message : String(err) };
      }
    });
    ipcMain.handle('web2electron:open-readonly-workspace-file', (event, payload) => {
      const rootPath = payload && payload.rootPath;
      const filePath = payload && payload.filePath;
      return openWorkspaceReadonlyFile(rootPath, filePath, event.sender);
    });
    ipcMain.handle('web2electron:open-external-file', async (event, payload) => {
      const filePath = payload && payload.filePath;
      const win = BrowserWindow.fromWebContents(event.sender) || firstMainWindow;
      if (!filePath) return { error: 'Invalid file path.' };
      const resolved = path.resolve(String(filePath));
      if (!fs.existsSync(resolved) || !fs.statSync(resolved).isFile()) return { error: 'Not a file.' };
      const opened = await shell.openPath(resolved);
      return opened ? { error: opened } : { opened: true, mode: 'external', filePath: resolved, fileName: getNameFromPath(resolved) };
    });
    ipcMain.handle('save-current-file', async (event, payload) => {
      try {
        const data = payload || {};
        let filePath = data.filePath && String(data.filePath).trim() ? path.resolve(String(data.filePath)) : null;
        if (!filePath) {
          const win = BrowserWindow.fromWebContents(event.sender) || firstMainWindow;
          const result = await dialog.showSaveDialog(win || undefined, {
            title: '저장',
            defaultPath: sanitizeSaveFileName(data.fileName),
            filters: [
              { name: 'Markdown', extensions: ['md'] },
              { name: 'Text', extensions: ['txt'] },
              { name: 'All files', extensions: ['*'] },
            ],
          });
          if (result.canceled || !result.filePath) return { canceled: true };
          filePath = result.filePath;
        }
        return writeTextFileUtf8(filePath, data.content);
      } catch (err) {
        return { error: err && err.message ? err.message : String(err) };
      }
    });
    ipcMain.handle('save-file-as', async (event, payload) => {
      try {
        const data = payload || {};
        const win = BrowserWindow.fromWebContents(event.sender) || firstMainWindow;
        const result = await dialog.showSaveDialog(win || undefined, {
          title: '다른 이름으로 저장',
          defaultPath: sanitizeSaveFileName(data.fileName),
          filters: [
            { name: 'Markdown', extensions: ['md'] },
            { name: 'Text', extensions: ['txt'] },
            { name: 'HTML', extensions: ['html', 'htm'] },
            { name: 'JSON', extensions: ['json'] },
            { name: 'All files', extensions: ['*'] },
          ],
        });
        if (result.canceled || !result.filePath) return { canceled: true };
        return writeTextFileUtf8(result.filePath, data.content);
      } catch (err) {
        return { error: err && err.message ? err.message : String(err) };
      }
    });
    createAppMenu();
    createWindow();
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
  });
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
