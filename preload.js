'use strict';

const { contextBridge, ipcRenderer } = require('electron');

const invokeAllowList = new Set([
  'web2electron:get-opened-file',
  'save-current-file',
  'save-file-as',
  'get-initial-file',
]);

const onAllowList = new Set([
  'open-external-file',
]);

contextBridge.exposeInMainWorld('web2electron', {
  getOpenedFilePath: () => ipcRenderer.invoke('web2electron:get-opened-file'),
  openFolderDialog: () => ipcRenderer.invoke('web2electron:open-folder-dialog'),
  listFolder: (payload) => ipcRenderer.invoke('web2electron:list-folder', payload || {}),
  readWorkspaceFile: (payload) => ipcRenderer.invoke('web2electron:read-workspace-file', payload || {}),
  openReadonlyWorkspaceFile: (payload) => ipcRenderer.invoke('web2electron:open-readonly-workspace-file', payload || {}),
  openExternalFile: (payload) => ipcRenderer.invoke('web2electron:open-external-file', payload || {}),
  createDirectory: (payload) => ipcRenderer.invoke('web2electron:create-directory', payload || {}),
  createWorkspaceFile: (payload) => ipcRenderer.invoke('web2electron:create-file', payload || {}),
  onOpenedFile: (callback) => {
    ipcRenderer.on('web2electron:opened-file', (_e, filePath, text, payload) => {
      if (typeof callback === 'function') callback(filePath, text, payload);
    });
  },
  onOpenedFolder: (callback) => {
    ipcRenderer.on('web2electron:opened-folder', (_e, folder) => {
      if (typeof callback === 'function') callback(folder);
    });
  },
  onMenuCommand: (callback) => {
    ipcRenderer.on('web2electron:menu-command', (_e, command) => {
      if (typeof callback === 'function') callback(command);
    });
  },
});

contextBridge.exposeInMainWorld('electron', {
  ipcRenderer: {
    invoke: (channel, payload) => {
      if (!invokeAllowList.has(channel)) return Promise.reject(new Error('IPC channel is not allowed.'));
      return ipcRenderer.invoke(channel, payload);
    },
    on: (channel, callback) => {
      if (!onAllowList.has(channel) || typeof callback !== 'function') return;
      ipcRenderer.on(channel, (_event, data) => callback(_event, data));
    },
  },
});
