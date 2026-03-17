const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("kmDesktop", {
  getDefaultPdfFolder: () => ipcRenderer.invoke("pdf:get-default-folder"),
  getPdfIndexSummary: () => ipcRenderer.invoke("pdf:get-index-summary"),
  pickPdfFolder: () => ipcRenderer.invoke("pdf:pick-folder"),
  indexPdfFolder: (folderPath) => ipcRenderer.invoke("pdf:index-folder", folderPath),
  searchPdf: ({ query, limit, config }) => ipcRenderer.invoke("pdf:search", { query, limit, config }),
  summarizePdfFile: (payload) => ipcRenderer.invoke("pdf:summarize-file", payload),
  onPdfSummarizeProgress: (handler) => {
    if (typeof handler !== "function") return () => {};
    const listener = (_event, payload) => {
      try {
        handler(payload);
      } catch {
        // Ignore renderer callback errors to keep IPC stable.
      }
    };
    ipcRenderer.on("pdf:summarize-progress", listener);
    return () => ipcRenderer.removeListener("pdf:summarize-progress", listener);
  },
  openPath: (targetPath) => ipcRenderer.invoke("system:open-path", targetPath),
  openPathAtPage: (targetPath, page) =>
    ipcRenderer.invoke("system:open-path-at-page", { targetPath, page }),
  testLocalAi: (config) => ipcRenderer.invoke("ai:test-connection", config),
  listLocalAiModels: (config) => ipcRenderer.invoke("ai:list-models", config),
  generateLocalAiText: (payload) => ipcRenderer.invoke("ai:generate-text", payload),
});
