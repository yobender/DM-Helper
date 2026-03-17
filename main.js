const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require("electron");
const { spawn } = require("child_process");
const fs = require("fs/promises");
const path = require("path");
const { pathToFileURL } = require("url");
const pdfParse = require("pdf-parse");

const DEFAULT_PDF_FOLDER = "C:\\Users\\Chris Bender\\OneDrive\\Desktop";
const MAX_CHARS_PER_FILE = 300000;
const DEFAULT_AI_ENDPOINT = "http://127.0.0.1:11434";
const DEFAULT_AI_MODEL = "llama3.1:8b";
const AI_TIMEOUT_MS = 120000;
const AI_PDF_SNIPPET_LIMIT = 6;
const OLLAMA_BOOT_RETRY_COUNT = 12;
const OLLAMA_BOOT_RETRY_DELAY_MS = 1000;
const PDF_INDEX_CACHE_FILENAME = "pdf-index-cache.v1.json";
const PDF_RETRIEVAL_CHUNK_SIZE = 1400;
const PDF_RETRIEVAL_CHUNK_OVERLAP = 240;
const PDF_EMBED_BATCH_SIZE = 16;
const PDF_HYBRID_CANDIDATE_FILE_LIMIT = 6;
const PDF_HYBRID_MATCH_LIMIT = 60;
const PDF_EMBEDDING_MODEL_CANDIDATES = [
  "all-minilm:latest",
  "all-minilm",
  "embeddinggemma:latest",
  "embeddinggemma",
  "qwen3-embedding:latest",
  "qwen3-embedding",
  "nomic-embed-text:latest",
  "nomic-embed-text",
  "mxbai-embed-large:latest",
  "mxbai-embed-large",
];
const AI_MODEL_LABELS = Object.freeze({
  "lorebound-pf2e:latest": "LoreBound PF2e Deep (20B)",
  "lorebound-pf2e-fast:latest": "LoreBound PF2e Fast (20B)",
  "lorebound-pf2e-minimal:latest": "LoreBound PF2e Minimal (20B)",
  "lorebound-pf2e-ultra-fast:latest": "LoreBound PF2e Ultra-Fast (1.5B)",
  "lorebound-pf2e-qwen:latest": "LoreBound PF2e Qwen Deep (32B)",
  "lorebound-pf2e-cpu:latest": "LoreBound PF2e CPU (20B)",
  "lorebound-pf2e-cpu-minimal:latest": "LoreBound PF2e CPU Minimal (20B)",
  "lorebound-pf2e-v2:latest": "LoreBound PF2e V2",
  "lorebound-pf2e-clean:latest": "LoreBound PF2e Clean",
  "gpt-oss:20b": "GPT-OSS Base (20B)",
  "gpt-oss-20b-fast:latest": "GPT-OSS Fast (20B)",
  "gpt-oss-20b-optimized:latest": "GPT-OSS Optimized (20B)",
  "gpt-oss-20b-cpu:latest": "GPT-OSS CPU (20B)",
  "llama3.1:8b": "Llama 3.1 (8B)",
  "qwen2.5-coder:1.5b-base": "Qwen 2.5 Coder Base (1.5B)",
  "qwen2.5-coder:32b": "Qwen 2.5 Coder (32B)",
});
const KINGDOM_RULES_DATA = loadKingdomRulesData();

let mainWindow = null;
const pdfViewerWindows = new Set();
let pdfIndexCache = {
  folderPath: "",
  indexedAt: "",
  files: [],
};
let ollamaBootPromise = null;
let pdfEmbeddingModelCache = {
  endpoint: "",
  checkedAt: 0,
  model: "",
};

function loadKingdomRulesData() {
  try {
    return require("./kingdom-rules-data.json");
  } catch {
    return {
      latestProfileId: "fallback",
      profiles: [
        {
          id: "fallback",
          label: "Kingdom Rules Profile",
          shortLabel: "Kingdom",
          summary: "Fallback kingdom rules profile used because the shared rules file could not be loaded.",
          turnStructure: [],
          aiContextSummary: [],
        },
      ],
    };
  }
}

wireProcessStabilityGuards();

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1200,
    minHeight: 760,
    title: "DM Helper",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: true,
    },
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
  wireSpellcheckContextMenu(mainWindow);
}

app.whenReady().then(async () => {
  await loadPdfIndexCacheFromDisk();
  registerIpc();
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

function wireProcessStabilityGuards() {
  const ignoreEpipe = (err) => isBrokenPipeError(err);

  if (process.stdout?.on) {
    process.stdout.on("error", (err) => {
      ignoreEpipe(err);
    });
  }
  if (process.stderr?.on) {
    process.stderr.on("error", (err) => {
      ignoreEpipe(err);
    });
  }

  process.on("uncaughtException", (err) => {
    if (ignoreEpipe(err)) return;
    safeMainLog("uncaughtException", err);
  });

  process.on("unhandledRejection", (reason) => {
    if (ignoreEpipe(reason)) return;
    safeMainLog("unhandledRejection", reason);
  });
}

function isBrokenPipeError(err) {
  const code = String(err?.code || "");
  const message = String(err?.message || err || "");
  return code === "EPIPE" || /broken pipe/i.test(message);
}

function safeMainLog(scope, err) {
  const name = String(err?.name || "Error");
  const message = String(err?.message || err || "Unknown error");
  try {
    process.stderr.write(`[${scope}] ${name}: ${message}\n`);
  } catch {
    // Ignore stderr write errors (including EPIPE).
  }
}

function getAiModelDisplayName(model) {
  const raw = String(model || "").trim();
  if (!raw) return "";
  return AI_MODEL_LABELS[raw.toLowerCase()] || raw;
}

function getPdfIndexCachePath() {
  return path.join(app.getPath("userData"), PDF_INDEX_CACHE_FILENAME);
}

function sanitizeEmbeddingVector(rawVector) {
  if (!Array.isArray(rawVector) || !rawVector.length) return null;
  const clean = rawVector
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value))
    .slice(0, 4096);
  if (!clean.length) return null;
  return normalizeEmbeddingVector(clean);
}

function normalizeEmbeddingVector(vector) {
  const clean = Array.isArray(vector) ? vector.map((value) => Number(value)).filter((value) => Number.isFinite(value)) : [];
  if (!clean.length) return null;
  let magnitude = 0;
  for (const value of clean) {
    magnitude += value * value;
  }
  const norm = Math.sqrt(magnitude);
  if (!Number.isFinite(norm) || norm <= 0) return null;
  return clean.map((value) => Number((value / norm).toFixed(6)));
}

function serializeEmbeddingVector(vector) {
  return Array.isArray(vector) ? vector.map((value) => Number(Number(value).toFixed(6))).filter((value) => Number.isFinite(value)) : [];
}

function buildFileSemanticSource(fileName, summary, pages) {
  const summaryText = normalizePdfText(String(summary || ""));
  if (summaryText) {
    return normalizePdfText(`${fileName}\n${summaryText}`).slice(0, 2400);
  }
  const preview = (Array.isArray(pages) ? pages : [])
    .slice(0, 3)
    .map((page) => normalizePdfText(page?.text || ""))
    .filter(Boolean)
    .join(" ");
  return normalizePdfText(`${fileName}\n${preview}`).slice(0, 2400);
}

function buildRetrievalChunksForPages(pages, chunkSize = PDF_RETRIEVAL_CHUNK_SIZE, overlap = PDF_RETRIEVAL_CHUNK_OVERLAP) {
  const chunks = [];
  for (const page of Array.isArray(pages) ? pages : []) {
    const pageNum = Number.parseInt(String(page?.page || chunks.length + 1), 10) || chunks.length + 1;
    const text = normalizePdfText(String(page?.text || ""));
    if (!text) continue;
    let start = 0;
    let chunkIndex = 0;
    while (start < text.length) {
      let end = Math.min(text.length, start + Math.max(320, Number(chunkSize) || PDF_RETRIEVAL_CHUNK_SIZE));
      if (end < text.length) {
        const window = text.slice(start, Math.min(text.length, end + 180));
        const breakpoints = [". ", "? ", "! ", "; ", ", ", " "];
        let best = -1;
        for (const marker of breakpoints) {
          const hit = window.lastIndexOf(marker);
          if (hit > best) best = hit;
        }
        if (best >= 240) {
          end = start + best + 1;
        }
      }
      const chunkText = normalizePdfText(text.slice(start, end));
      if (chunkText.length >= 120) {
        chunks.push({
          id: `p${pageNum}-c${chunkIndex + 1}`,
          page: pageNum,
          pageEnd: pageNum,
          text: chunkText,
          textLower: chunkText.toLowerCase(),
          charCount: chunkText.length,
          embedding: null,
        });
        chunkIndex += 1;
      }
      if (end >= text.length) break;
      start = Math.max(start + 1, end - Math.max(60, Number(overlap) || PDF_RETRIEVAL_CHUNK_OVERLAP));
    }
  }
  return chunks;
}

function sanitizeRetrievalChunk(rawChunk, fallbackChunk = null) {
  const source = rawChunk && typeof rawChunk === "object" ? rawChunk : {};
  const baseText = normalizePdfText(String(source.text || fallbackChunk?.text || ""));
  if (!baseText) return null;
  const page = Number.parseInt(String(source.page ?? fallbackChunk?.page ?? 1), 10) || 1;
  const pageEnd = Number.parseInt(String(source.pageEnd ?? fallbackChunk?.pageEnd ?? page), 10) || page;
  const embedding = sanitizeEmbeddingVector(source.embedding || fallbackChunk?.embedding);
  return {
    id: String(source.id || fallbackChunk?.id || `p${page}-c1`).trim() || `p${page}-c1`,
    page,
    pageEnd: Math.max(page, pageEnd),
    text: baseText,
    textLower: baseText.toLowerCase(),
    charCount: baseText.length,
    embedding,
  };
}

function sanitizeFileRetrieval(rawRetrieval, pages, fileName, summary) {
  const source = rawRetrieval && typeof rawRetrieval === "object" ? rawRetrieval : {};
  const baseChunks = buildRetrievalChunksForPages(pages);
  const existingChunks = new Map();
  if (Array.isArray(source.chunks)) {
    for (const rawChunk of source.chunks) {
      const cleanChunk = sanitizeRetrievalChunk(rawChunk);
      if (!cleanChunk) continue;
      existingChunks.set(cleanChunk.id, cleanChunk);
    }
  }
  const chunks = baseChunks
    .map((baseChunk) => sanitizeRetrievalChunk(existingChunks.get(baseChunk.id), baseChunk))
    .filter(Boolean);
  return {
    chunkSize: Number.parseInt(String(source.chunkSize || PDF_RETRIEVAL_CHUNK_SIZE), 10) || PDF_RETRIEVAL_CHUNK_SIZE,
    chunkOverlap: Number.parseInt(String(source.chunkOverlap || PDF_RETRIEVAL_CHUNK_OVERLAP), 10) || PDF_RETRIEVAL_CHUNK_OVERLAP,
    embeddingModel: String(source.embeddingModel || "").trim(),
    fileEmbedding: sanitizeEmbeddingVector(source.fileEmbedding),
    fileEmbeddingUpdatedAt: String(source.fileEmbeddingUpdatedAt || "").trim(),
    fileSemanticSource: buildFileSemanticSource(fileName, summary, pages),
    chunks,
  };
}

function sanitizeIndexedPage(rawPage, fallbackPageNumber) {
  const pageNum = Number.parseInt(String(rawPage?.page ?? fallbackPageNumber), 10) || fallbackPageNumber;
  const text = normalizePdfText(String(rawPage?.text || ""));
  if (!text) return null;
  return {
    page: pageNum,
    text,
    textLower: text.toLowerCase(),
    charCount: text.length,
  };
}

function normalizeStoredPdfSummaryText(text) {
  const source = String(text || "").replace(/\r\n?/g, "\n");
  if (!source) return "";
  const lines = source.split("\n").map((line) => line.replace(/[ \t]+/g, " ").trimEnd());
  const cleaned = [];
  let lastBlank = false;
  for (const line of lines) {
    const normalized = line.trim() ? line : "";
    if (!normalized) {
      if (lastBlank) continue;
      cleaned.push("");
      lastBlank = true;
      continue;
    }
    lastBlank = false;
    cleaned.push(normalized);
  }
  return cleaned.join("\n").trim().slice(0, 24000);
}

function sanitizeIndexedFile(rawFile) {
  const filePath = String(rawFile?.path || "").trim();
  const fileName = String(rawFile?.fileName || path.basename(filePath || "")).trim();
  const pagesRaw = Array.isArray(rawFile?.pages) ? rawFile.pages : [];
  const pages = pagesRaw
    .map((rawPage, index) => sanitizeIndexedPage(rawPage, index + 1))
    .filter(Boolean);
  if (!filePath || !fileName || !pages.length) return null;
  const merged = normalizePdfText(pages.map((page) => page.text).join(" "));
  const summary = normalizeStoredPdfSummaryText(rawFile?.summary);
  const summaryUpdatedAt = String(rawFile?.summaryUpdatedAt || "").trim();
  const retrieval = sanitizeFileRetrieval(rawFile?.retrieval, pages, fileName, summary);
  return {
    path: filePath,
    fileName,
    pages,
    text: merged,
    textLower: merged.toLowerCase(),
    charCount: merged.length,
    summary,
    summaryUpdatedAt,
    retrieval,
  };
}

function sanitizePdfIndexCache(rawCache) {
  const filesRaw = Array.isArray(rawCache?.files) ? rawCache.files : [];
  const files = filesRaw.map((file) => sanitizeIndexedFile(file)).filter(Boolean);
  return {
    folderPath: String(rawCache?.folderPath || "").trim(),
    indexedAt: String(rawCache?.indexedAt || "").trim(),
    files,
  };
}

function buildPdfIndexSummary() {
  const files = pdfIndexCache.files
    .map((file) => ({
      fileName: String(file?.fileName || "").trim(),
      path: String(file?.path || "").trim(),
      pageCount: Array.isArray(file?.pages) ? file.pages.length : 0,
      charCount: Number.parseInt(String(file?.charCount || 0), 10) || 0,
      summary: normalizeStoredPdfSummaryText(file?.summary),
      summaryUpdatedAt: String(file?.summaryUpdatedAt || "").trim(),
      retrievalChunks: Array.isArray(file?.retrieval?.chunks) ? file.retrieval.chunks.length : 0,
      retrievalEmbeddingModel: String(file?.retrieval?.embeddingModel || "").trim(),
    }))
    .filter((file) => file.fileName && file.path)
    .sort((a, b) => a.fileName.localeCompare(b.fileName) || a.path.localeCompare(b.path));
  const fileNames = files.map((file) => file.fileName);
  return {
    folderPath: String(pdfIndexCache.folderPath || "").trim(),
    indexedAt: String(pdfIndexCache.indexedAt || "").trim(),
    count: pdfIndexCache.files.length,
    fileNames,
    files,
  };
}

async function loadPdfIndexCacheFromDisk() {
  try {
    const cachePath = getPdfIndexCachePath();
    const raw = await fs.readFile(cachePath, "utf8");
    const parsed = JSON.parse(raw);
    pdfIndexCache = sanitizePdfIndexCache(parsed);
  } catch {
    pdfIndexCache = {
      folderPath: "",
      indexedAt: "",
      files: [],
    };
  }
}

async function savePdfIndexCacheToDisk() {
  const cachePath = getPdfIndexCachePath();
  const serializable = {
    folderPath: String(pdfIndexCache.folderPath || "").trim(),
    indexedAt: String(pdfIndexCache.indexedAt || "").trim(),
    files: pdfIndexCache.files.map((file) => ({
      path: String(file?.path || "").trim(),
      fileName: String(file?.fileName || "").trim(),
      summary: normalizeStoredPdfSummaryText(file?.summary),
      summaryUpdatedAt: String(file?.summaryUpdatedAt || "").trim(),
      retrieval: {
        chunkSize: Number.parseInt(String(file?.retrieval?.chunkSize || PDF_RETRIEVAL_CHUNK_SIZE), 10) || PDF_RETRIEVAL_CHUNK_SIZE,
        chunkOverlap: Number.parseInt(String(file?.retrieval?.chunkOverlap || PDF_RETRIEVAL_CHUNK_OVERLAP), 10) || PDF_RETRIEVAL_CHUNK_OVERLAP,
        embeddingModel: String(file?.retrieval?.embeddingModel || "").trim(),
        fileEmbeddingUpdatedAt: String(file?.retrieval?.fileEmbeddingUpdatedAt || "").trim(),
        fileEmbedding: serializeEmbeddingVector(file?.retrieval?.fileEmbedding),
        chunks: (Array.isArray(file?.retrieval?.chunks) ? file.retrieval.chunks : [])
          .map((chunk) => sanitizeRetrievalChunk(chunk))
          .filter(Boolean)
          .map((chunk) => ({
            id: chunk.id,
            page: chunk.page,
            pageEnd: chunk.pageEnd,
            text: chunk.text,
            embedding: serializeEmbeddingVector(chunk.embedding),
          })),
      },
      pages: (Array.isArray(file?.pages) ? file.pages : [])
        .map((page, index) => sanitizeIndexedPage(page, index + 1))
        .filter(Boolean)
        .map((page) => ({
          page: page.page,
          text: page.text,
        })),
    })),
  };
  await fs.writeFile(cachePath, JSON.stringify(serializable), "utf8");
}

function emitPdfSummarizeProgress(sender, payload) {
  if (!sender) return;
  try {
    if (typeof sender.isDestroyed === "function" && sender.isDestroyed()) return;
    sender.send("pdf:summarize-progress", {
      at: new Date().toISOString(),
      ...payload,
    });
  } catch {
    // Ignore renderer IPC delivery issues.
  }
}

function registerIpc() {
  ipcMain.handle("pdf:get-default-folder", async () => {
    return DEFAULT_PDF_FOLDER;
  });

  ipcMain.handle("pdf:get-index-summary", async () => {
    return buildPdfIndexSummary();
  });

  ipcMain.handle("pdf:pick-folder", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: "Choose PDF Folder",
      properties: ["openDirectory"],
      defaultPath: DEFAULT_PDF_FOLDER,
    });
    if (result.canceled || !result.filePaths?.length) return "";
    return result.filePaths[0];
  });

  ipcMain.handle("pdf:index-folder", async (_event, folderPath) => {
    if (!folderPath) {
      throw new Error("Folder path is required.");
    }

    const pdfPaths = await collectPdfPaths(folderPath);
    const existingByPath = new Map(
      (Array.isArray(pdfIndexCache.files) ? pdfIndexCache.files : [])
        .map((file) => [String(file?.path || ""), file])
        .filter((entry) => entry[0])
    );
    const indexedFiles = [];
    let failed = 0;

    for (const pdfPath of pdfPaths) {
      try {
        const buffer = await fs.readFile(pdfPath);
        const pages = await extractPdfPages(buffer, MAX_CHARS_PER_FILE);
        const existing = existingByPath.get(pdfPath);
        const merged = normalizePdfText(pages.map((page) => page.text).join(" "));
        const canReuseRetrieval =
          existing &&
          normalizePdfText(String(existing?.text || existing?.pages?.map((page) => page.text || "").join(" ") || "")) === merged;
        indexedFiles.push(sanitizeIndexedFile({
          path: pdfPath,
          fileName: path.basename(pdfPath),
          pages,
          summary: normalizeStoredPdfSummaryText(existing?.summary),
          summaryUpdatedAt: String(existing?.summaryUpdatedAt || "").trim(),
          retrieval: canReuseRetrieval ? existing?.retrieval : null,
        }));
      } catch {
        failed += 1;
      }
    }

    pdfIndexCache = {
      folderPath,
      indexedAt: new Date().toISOString(),
      files: indexedFiles,
    };

    await savePdfIndexCacheToDisk();

    return {
      ...buildPdfIndexSummary(),
      failed,
    };
  });

  ipcMain.handle("pdf:read-file-data", async (_event, targetPath) => {
    const resolvedPath = String(targetPath || "").trim();
    if (!resolvedPath) {
      throw new Error("PDF path is required.");
    }
    return fs.readFile(resolvedPath);
  });

  ipcMain.handle("pdf:search", async (_event, payload) => {
    const query = String(payload?.query || "").trim();
    const limitRaw = Number.parseInt(String(payload?.limit || "20"), 10);
    const limit = Number.isNaN(limitRaw) ? 20 : Math.max(1, Math.min(limitRaw, 100));
    const config = normalizeAiConfig(payload?.config);

    if (!query) return { results: [] };
    if (!pdfIndexCache.files.length) {
      throw new Error("No PDFs indexed yet. Run Index PDFs first.");
    }

    return searchIndexedPdfHybrid(query, limit, { config });
  });

  ipcMain.handle("pdf:summarize-file", async (event, payload) => {
    if (!pdfIndexCache.files.length) {
      throw new Error("No PDFs indexed yet. Run Index PDFs first.");
    }

    const target = resolveIndexedPdfFile(payload);
    if (!target) {
      throw new Error("Could not find that indexed PDF. Re-index and try again.");
    }

    const force = payload?.force === true;
    if (!force && normalizePdfText(String(target.summary || ""))) {
      emitPdfSummarizeProgress(event?.sender, {
        stage: "done",
        fileName: target.fileName,
        path: target.path,
        current: 1,
        total: 1,
        message: `Loaded saved summary for ${target.fileName}.`,
      });
      return {
        ...buildPdfSummaryResponse(target),
        reused: true,
        chunks: 0,
      };
    }

    const config = normalizeAiConfig(payload?.config);
    const chunkSummaryConfig = buildPdfSummaryConfig(config, "chunk");
    const finalSummaryConfig = buildPdfSummaryConfig(config, "final");
    const text = normalizePdfText(String(target.text || target.pages?.map((page) => page.text || "").join(" ") || ""));
    if (!text) {
      throw new Error("Indexed PDF text is empty for this file.");
    }

    const chunks = splitTextIntoChunks(text, 7600, 6);
    const progressTotal = Math.max(2, chunks.length + 1);
    emitPdfSummarizeProgress(event?.sender, {
      stage: "start",
      fileName: target.fileName,
      path: target.path,
      current: 0,
      total: progressTotal,
      message: `Starting summary for ${target.fileName}...`,
    });
    const chunkSummaries = [];
    for (let i = 0; i < chunks.length; i += 1) {
      const chunkText = chunks[i];
      emitPdfSummarizeProgress(event?.sender, {
        stage: "chunk",
        fileName: target.fileName,
        path: target.path,
        current: i + 1,
        total: progressTotal,
        message: `Summarizing chunk ${i + 1}/${chunks.length}...`,
      });
      const prompt = buildPdfChunkSummaryPrompt({
        fileName: target.fileName,
        chunkText,
        index: i + 1,
        total: chunks.length,
      });
      try {
        const raw = await generateWithOllama(chunkSummaryConfig, prompt);
        const cleaned = sanitizeAiTextOutput(raw);
        chunkSummaries.push(cleaned || extractiveChunkSummary(chunkText));
      } catch {
        chunkSummaries.push(extractiveChunkSummary(chunkText));
      }
    }

    const combinedPrompt = buildPdfFinalSummaryPrompt({
      fileName: target.fileName,
      chunkSummaries,
    });
    emitPdfSummarizeProgress(event?.sender, {
      stage: "combine",
      fileName: target.fileName,
      path: target.path,
      current: Math.max(1, chunks.length),
      total: progressTotal,
      message: `Combining chunk summaries for ${target.fileName}...`,
    });
    let finalSummary = "";
    try {
      const rawFinal = await generateWithOllama(finalSummaryConfig, combinedPrompt);
      finalSummary = sanitizeAiTextOutput(rawFinal);
    } catch {
      finalSummary = "";
    }
    if (!isUsefulPdfSummary(finalSummary)) {
      emitPdfSummarizeProgress(event?.sender, {
        stage: "combine",
        fileName: target.fileName,
        path: target.path,
        current: Math.max(1, chunks.length),
        total: progressTotal,
        message: `Refining summary structure for ${target.fileName}...`,
      });
      try {
        const retryPrompt = buildPdfFinalSummaryRetryPrompt({
          fileName: target.fileName,
          chunkSummaries,
        });
        const retryConfig = {
          ...finalSummaryConfig,
          maxOutputTokens: Math.max(Number(finalSummaryConfig.maxOutputTokens || 0) || 0, 1600),
          timeoutSec: Math.max(Number(finalSummaryConfig.timeoutSec || 0) || 0, 480),
        };
        retryConfig.timeoutMs = Math.max(15000, Number(retryConfig.timeoutSec || 0) * 1000);
        const rawRetry = await generateWithOllama(retryConfig, retryPrompt);
        const cleanedRetry = sanitizeAiTextOutput(rawRetry);
        if (isUsefulPdfSummary(cleanedRetry)) {
          finalSummary = cleanedRetry;
        }
      } catch {
        // Fall back below.
      }
    }
    if (!isUsefulPdfSummary(finalSummary)) {
      finalSummary = combineChunkSummariesFallback(target.fileName, chunkSummaries);
    }

    target.summary = normalizeStoredPdfSummaryText(finalSummary);
    target.summaryUpdatedAt = new Date().toISOString();
    await savePdfIndexCacheToDisk();
    emitPdfSummarizeProgress(event?.sender, {
      stage: "done",
      fileName: target.fileName,
      path: target.path,
      current: progressTotal,
      total: progressTotal,
      message: `Summary complete for ${target.fileName}.`,
    });

    return {
      ...buildPdfSummaryResponse(target),
      reused: false,
      chunks: chunks.length,
    };
  });

  ipcMain.handle("system:open-path", async (_event, targetPath) => {
    if (!targetPath) return false;
    await shell.openPath(targetPath);
    return true;
  });

  ipcMain.handle("system:open-path-at-page", async (_event, payload) => {
    const targetPath = String(payload?.targetPath || "").trim();
    if (!targetPath) return false;
    const pageRaw = Number.parseInt(String(payload?.page || "0"), 10);
    const page = Number.isNaN(pageRaw) ? 0 : Math.max(0, pageRaw);
    if (!page) {
      await shell.openPath(targetPath);
      return true;
    }

    try {
      await openPdfAtPageInApp(targetPath, page);
      return true;
    } catch {
      await shell.openPath(targetPath);
      return true;
    }
  });

  ipcMain.handle("ai:test-connection", async (_event, rawConfig) => {
    const config = normalizeAiConfig(rawConfig);
    return testOllamaConnection(config);
  });

  ipcMain.handle("ai:list-models", async (_event, rawConfig) => {
    const config = normalizeAiConfig(rawConfig);
    const models = await listOllamaModelsWithRecovery(config.endpoint, Math.min(20000, Number(config?.timeoutMs) || 10000));
    return {
      endpoint: config.endpoint,
      models,
    };
  });

  ipcMain.handle("ai:generate-text", async (_event, payload) => {
    const input = String(payload?.input || "").trim();
    if (!input) {
      throw new Error("Draft text is required.");
    }

    const mode = String(payload?.mode || "session");
    const activeTab = String(payload?.context?.activeTab || "").trim();
    const selectedPdfFile = String(payload?.context?.selectedPdfFile || "").trim();
    const context =
      payload?.context && typeof payload.context === "object" && !Array.isArray(payload.context)
        ? payload.context
        : {};
    const config = normalizeAiConfig(payload?.config);
    const enrichedContext = {
      ...context,
      selectedPdfFile,
      selectedPdfPreview: selectedPdfFile ? buildSelectedPdfPreview(selectedPdfFile, 1200) : "",
      pdfContextEnabled: false,
      pdfSnippets: [],
      pdfIndexedFiles: getIndexedPdfFileNames(config.compactContext ? 20 : 40),
      pdfIndexedFileCount: pdfIndexCache.files.length,
    };
    if (config.usePdfContext && pdfIndexCache.files.length) {
      const query = [
        input,
        String(context?.tabContext || ""),
        String(context?.latestSession?.summary || ""),
        String(context?.latestSession?.nextPrep || ""),
      ]
        .join(" ")
        .trim();
      const snippetLimit = config.compactContext ? Math.min(3, AI_PDF_SNIPPET_LIMIT) : AI_PDF_SNIPPET_LIMIT;
      enrichedContext.pdfSnippets = await collectPdfContextForAi(query, snippetLimit, {
        config,
        preferredFileName: selectedPdfFile,
      });
      if (!enrichedContext.pdfSnippets.length && selectedPdfFile && (activeTab === "pdf" || isPdfGroundedQuestion(input))) {
        enrichedContext.pdfSnippets = collectSelectedPdfFallbackContext(selectedPdfFile, query, snippetLimit);
      }
      enrichedContext.pdfContextEnabled = enrichedContext.pdfSnippets.length > 0;
    }

    const userPrompt = buildAiUserPrompt({
      mode,
      input,
      context: enrichedContext,
      compactContext: config.compactContext,
    });
    const text = await generateWithOllama(config, userPrompt);
    const finalized = finalizeAiOutput({
      rawText: text,
      mode,
      input,
      tabId: String(enrichedContext?.activeTab || "").trim(),
    });
    const scopedSourceReply = maybeBuildSourceScopeReply(mode, input, enrichedContext);
    if (scopedSourceReply) {
      return {
        text: scopedSourceReply,
        model: config.model,
        endpoint: config.endpoint,
        usedFallback: false,
        filtered: true,
        fallbackReason: "",
      };
    }
    return {
      text: finalized.text,
      model: config.model,
      endpoint: config.endpoint,
      usedFallback: finalized.usedFallback,
      filtered: finalized.filtered,
      fallbackReason: finalized.fallbackReason || "",
    };
  });
}

async function collectPdfPaths(rootDir) {
  const out = [];
  const stack = [rootDir];

  while (stack.length) {
    const current = stack.pop();
    const entries = await fs.readdir(current, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        stack.push(fullPath);
      } else if (entry.isFile() && entry.name.toLowerCase().endsWith(".pdf")) {
        out.push(fullPath);
      }
    }
  }

  return out;
}

async function extractPdfPages(buffer, maxCharsPerFile = MAX_CHARS_PER_FILE) {
  const pages = [];
  let remainingChars = Math.max(1, Number.parseInt(String(maxCharsPerFile || MAX_CHARS_PER_FILE), 10));

  await pdfParse(buffer, {
    pagerender: async (pageData) => {
      if (remainingChars <= 0) {
        pages.push({
          page: pages.length + 1,
          text: "",
          textLower: "",
          charCount: 0,
        });
        return "";
      }

      const textContent = await pageData.getTextContent({
        normalizeWhitespace: true,
        disableCombineTextItems: false,
      });
      const rawText = textContent.items.map((item) => String(item.str || "")).join(" ");
      const cleaned = normalizePdfText(rawText);
      const clipped = cleaned.slice(0, remainingChars);
      remainingChars -= clipped.length;

      pages.push({
        page: pages.length + 1,
        text: clipped,
        textLower: clipped.toLowerCase(),
        charCount: clipped.length,
      });

      return clipped;
    },
  });

  return pages.filter((page) => page.charCount > 0);
}

function normalizePdfText(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function buildSearchParts(query) {
  const phrase = normalizePdfText(query).toLowerCase();
  const words = phrase
    .split(/\s+/)
    .map((word) => word.replace(/[^a-z0-9]/g, ""))
    .filter(Boolean);
  return { phrase, words };
}

function getIndexedPages(file) {
  if (Array.isArray(file?.pages) && file.pages.length) {
    return file.pages
      .map((page, index) => {
        const text = String(page?.text || "");
        return {
          page: Number.parseInt(String(page?.page || index + 1), 10) || index + 1,
          text,
          textLower: String(page?.textLower || text.toLowerCase()),
        };
      })
      .filter((page) => page.text.length > 0);
  }

  const text = String(file?.text || "");
  return text
    ? [
        {
          page: 1,
          text,
          textLower: String(file?.textLower || text.toLowerCase()),
        },
      ]
    : [];
}

function ensureFileRetrievalState(file) {
  if (!file || typeof file !== "object") return { chunks: [], embeddingModel: "", fileEmbedding: null, fileSemanticSource: "" };
  if (!file.retrieval || typeof file.retrieval !== "object") {
    file.retrieval = sanitizeFileRetrieval(null, getIndexedPages(file), String(file?.fileName || ""), String(file?.summary || ""));
    return file.retrieval;
  }
  const expectedSource = buildFileSemanticSource(String(file?.fileName || ""), String(file?.summary || ""), getIndexedPages(file));
  if (String(file.retrieval.fileSemanticSource || "") !== expectedSource) {
    file.retrieval.fileSemanticSource = expectedSource;
    file.retrieval.fileEmbedding = null;
    file.retrieval.fileEmbeddingUpdatedAt = "";
  }
  if (!Array.isArray(file.retrieval.chunks) || !file.retrieval.chunks.length) {
    file.retrieval = sanitizeFileRetrieval(file.retrieval, getIndexedPages(file), String(file?.fileName || ""), String(file?.summary || ""));
  }
  return file.retrieval;
}

function dotProduct(vectorA, vectorB) {
  if (!Array.isArray(vectorA) || !Array.isArray(vectorB) || !vectorA.length || vectorA.length !== vectorB.length) return 0;
  let total = 0;
  for (let i = 0; i < vectorA.length; i += 1) {
    total += Number(vectorA[i] || 0) * Number(vectorB[i] || 0);
  }
  return total;
}

function getChunkKey(file, chunk) {
  const filePath = String(file?.path || "").trim();
  const chunkId = String(chunk?.id || "").trim();
  return `${filePath}::${chunkId}`;
}

function buildChunkSearchResult(file, chunk, score, firstHit = -1) {
  return {
    key: getChunkKey(file, chunk),
    fileName: String(file?.fileName || "").trim(),
    path: String(file?.path || "").trim(),
    page: Number.parseInt(String(chunk?.page || 1), 10) || 1,
    pageEnd: Number.parseInt(String(chunk?.pageEnd || chunk?.page || 1), 10) || (Number.parseInt(String(chunk?.page || 1), 10) || 1),
    score,
    snippet: makeSnippet(String(chunk?.text || ""), firstHit),
    chunkId: String(chunk?.id || "").trim(),
  };
}

function collectLexicalChunkMatches(files, query, limit = PDF_HYBRID_MATCH_LIMIT, preferredFileName = "") {
  const searchParts = buildSearchParts(query);
  if (!searchParts.words.length && !searchParts.phrase) return [];
  const preferred = String(preferredFileName || "").trim().toLowerCase();
  const ranked = [];
  for (const file of files) {
    const retrieval = ensureFileRetrievalState(file);
    for (const chunk of retrieval.chunks) {
      const match = scoreTextAgainstQuery(chunk.textLower, searchParts);
      if (!match.score) continue;
      const fileBoost = preferred && String(file?.fileName || "").trim().toLowerCase() === preferred ? 3 : 0;
      ranked.push({
        ...buildChunkSearchResult(file, chunk, match.score + fileBoost, match.firstHit),
        lexicalScore: match.score + fileBoost,
      });
    }
  }
  ranked.sort((a, b) => b.lexicalScore - a.lexicalScore || a.fileName.localeCompare(b.fileName) || a.page - b.page);
  return ranked.slice(0, Math.max(1, Math.min(Number(limit) || PDF_HYBRID_MATCH_LIMIT, 120)));
}

async function pickAvailableEmbeddingModel(endpoint, timeoutMs = 10000) {
  const safeEndpoint = String(endpoint || DEFAULT_AI_ENDPOINT).replace(/\/+$/g, "");
  const now = Date.now();
  if (
    pdfEmbeddingModelCache.endpoint === safeEndpoint &&
    now - Number(pdfEmbeddingModelCache.checkedAt || 0) < 120000
  ) {
    return pdfEmbeddingModelCache.model || "";
  }
  const models = await listOllamaModelsWithRecovery(safeEndpoint, timeoutMs);
  const match = PDF_EMBEDDING_MODEL_CANDIDATES.find((candidate) =>
    models.some((model) => String(model || "").trim().toLowerCase() === candidate.toLowerCase())
  );
  pdfEmbeddingModelCache = {
    endpoint: safeEndpoint,
    checkedAt: now,
    model: match || "",
  };
  return match || "";
}

async function requestOllamaEmbeddings(endpoint, model, inputs, timeoutMs = 45000) {
  const safeInputs = (Array.isArray(inputs) ? inputs : [inputs]).map((item) => normalizePdfText(String(item || ""))).filter(Boolean);
  if (!safeInputs.length) return [];

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), Math.max(10000, Number(timeoutMs) || 45000));
  try {
    try {
      const data = await requestOllamaJson(
        `${endpoint}/api/embed`,
        {
          model,
          input: safeInputs,
          truncate: true,
        },
        controller.signal
      );
      const vectors = Array.isArray(data?.embeddings)
        ? data.embeddings
        : Array.isArray(data?.embedding)
          ? [data.embedding]
          : [];
      return vectors.map((vector) => sanitizeEmbeddingVector(vector)).filter(Boolean);
    } catch (err) {
      const message = String(err?.message || err || "").toLowerCase();
      if (!message.includes("/api/embed")) {
        // requestOllamaJson uses generic errors, so retry the legacy endpoint unless we clearly timed out.
      }
      const vectors = [];
      for (const text of safeInputs) {
        const data = await requestOllamaJson(
          `${endpoint}/api/embeddings`,
          {
            model,
            prompt: text,
          },
          controller.signal
        );
        vectors.push(sanitizeEmbeddingVector(data?.embedding));
      }
      return vectors.filter(Boolean);
    }
  } catch (err) {
    if (err?.name === "AbortError") {
      throw new Error(`Embedding request timed out after ${Math.round((Number(timeoutMs) || 45000) / 1000)}s.`);
    }
    throw err;
  } finally {
    clearTimeout(timeout);
  }
}

async function ensureFileEmbeddings(files, endpoint, model, timeoutMs = 45000) {
  const pending = [];
  for (const file of files) {
    const retrieval = ensureFileRetrievalState(file);
    if (retrieval.embeddingModel && retrieval.embeddingModel !== model) {
      retrieval.fileEmbedding = null;
      retrieval.fileEmbeddingUpdatedAt = "";
      for (const chunk of retrieval.chunks) {
        chunk.embedding = null;
      }
      retrieval.embeddingModel = "";
    }
    if (Array.isArray(retrieval.fileEmbedding) && retrieval.fileEmbedding.length) {
      continue;
    }
    const semanticText = normalizePdfText(String(retrieval.fileSemanticSource || buildFileSemanticSource(file.fileName, file.summary, getIndexedPages(file))));
    if (!semanticText) continue;
    pending.push({ file, text: semanticText });
  }
  if (!pending.length) return false;

  let changed = false;
  for (let index = 0; index < pending.length; index += PDF_EMBED_BATCH_SIZE) {
    const batch = pending.slice(index, index + PDF_EMBED_BATCH_SIZE);
    const vectors = await requestOllamaEmbeddings(endpoint, model, batch.map((item) => item.text), timeoutMs);
    for (let offset = 0; offset < batch.length; offset += 1) {
      const vector = vectors[offset];
      if (!vector) continue;
      const retrieval = ensureFileRetrievalState(batch[offset].file);
      retrieval.fileEmbedding = vector;
      retrieval.fileEmbeddingUpdatedAt = new Date().toISOString();
      retrieval.embeddingModel = model;
      changed = true;
    }
  }
  return changed;
}

async function ensureChunkEmbeddingsForFiles(files, endpoint, model, timeoutMs = 45000) {
  const pending = [];
  for (const file of files) {
    const retrieval = ensureFileRetrievalState(file);
    if (retrieval.embeddingModel && retrieval.embeddingModel !== model) {
      retrieval.fileEmbedding = null;
      retrieval.fileEmbeddingUpdatedAt = "";
      for (const chunk of retrieval.chunks) {
        chunk.embedding = null;
      }
      retrieval.embeddingModel = "";
    }
    for (const chunk of retrieval.chunks) {
      if (Array.isArray(chunk.embedding) && chunk.embedding.length) continue;
      pending.push({ file, chunk });
    }
  }
  if (!pending.length) return false;

  let changed = false;
  for (let index = 0; index < pending.length; index += PDF_EMBED_BATCH_SIZE) {
    const batch = pending.slice(index, index + PDF_EMBED_BATCH_SIZE);
    const vectors = await requestOllamaEmbeddings(endpoint, model, batch.map((item) => item.chunk.text), timeoutMs);
    for (let offset = 0; offset < batch.length; offset += 1) {
      const vector = vectors[offset];
      if (!vector) continue;
      const { file, chunk } = batch[offset];
      const retrieval = ensureFileRetrievalState(file);
      const target = retrieval.chunks.find((entry) => entry.id === chunk.id);
      if (!target) continue;
      target.embedding = vector;
      retrieval.embeddingModel = model;
      changed = true;
    }
  }
  return changed;
}

function collectSemanticFileCandidates(files, queryEmbedding, preferredFileName = "", limit = PDF_HYBRID_CANDIDATE_FILE_LIMIT) {
  const preferred = String(preferredFileName || "").trim().toLowerCase();
  const ranked = [];
  for (const file of files) {
    const retrieval = ensureFileRetrievalState(file);
    if (!Array.isArray(retrieval.fileEmbedding) || !retrieval.fileEmbedding.length) continue;
    let score = dotProduct(queryEmbedding, retrieval.fileEmbedding);
    if (preferred && String(file?.fileName || "").trim().toLowerCase() === preferred) {
      score += 0.035;
    }
    ranked.push({ file, score });
  }
  ranked.sort((a, b) => b.score - a.score || String(a.file?.fileName || "").localeCompare(String(b.file?.fileName || "")));
  return ranked.slice(0, Math.max(1, Math.min(Number(limit) || PDF_HYBRID_CANDIDATE_FILE_LIMIT, 20)));
}

function collectSemanticChunkMatches(files, queryEmbedding, limit = PDF_HYBRID_MATCH_LIMIT, preferredFileName = "") {
  const preferred = String(preferredFileName || "").trim().toLowerCase();
  const ranked = [];
  for (const file of files) {
    const retrieval = ensureFileRetrievalState(file);
    for (const chunk of retrieval.chunks) {
      if (!Array.isArray(chunk.embedding) || !chunk.embedding.length) continue;
      let score = dotProduct(queryEmbedding, chunk.embedding);
      if (preferred && String(file?.fileName || "").trim().toLowerCase() === preferred) {
        score += 0.02;
      }
      if (!Number.isFinite(score) || score <= 0) continue;
      ranked.push({
        ...buildChunkSearchResult(file, chunk, score, -1),
        semanticScore: score,
      });
    }
  }
  ranked.sort((a, b) => b.semanticScore - a.semanticScore || a.fileName.localeCompare(b.fileName) || a.page - b.page);
  return ranked.slice(0, Math.max(1, Math.min(Number(limit) || PDF_HYBRID_MATCH_LIMIT, 120)));
}

function fuseHybridMatches(lexicalMatches, semanticMatches, limit = 20) {
  const fused = new Map();
  const rankConstant = 60;
  lexicalMatches.forEach((match, index) => {
    const current = fused.get(match.key) || { ...match, lexicalRank: 0, semanticRank: 0, lexicalScore: 0, semanticScore: 0, fusedScore: 0 };
    current.lexicalRank = index + 1;
    current.lexicalScore = Math.max(Number(current.lexicalScore || 0), Number(match.lexicalScore || match.score || 0));
    current.fusedScore += 1 / (rankConstant + index + 1);
    fused.set(match.key, current);
  });
  semanticMatches.forEach((match, index) => {
    const current = fused.get(match.key) || { ...match, lexicalRank: 0, semanticRank: 0, lexicalScore: 0, semanticScore: 0, fusedScore: 0 };
    current.semanticRank = index + 1;
    current.semanticScore = Math.max(Number(current.semanticScore || 0), Number(match.semanticScore || match.score || 0));
    current.fusedScore += 1 / (rankConstant + index + 1);
    if (!current.snippet && match.snippet) current.snippet = match.snippet;
    fused.set(match.key, current);
  });

  const entries = [...fused.values()].map((entry) => {
    const searchMode = entry.lexicalRank && entry.semanticRank ? "hybrid" : entry.semanticRank ? "semantic" : "lexical";
    return {
      ...entry,
      searchMode,
      score: Math.round(Number(entry.fusedScore || 0) * 100000),
    };
  });
  entries.sort(
    (a, b) =>
      b.score - a.score ||
      Number(b.semanticScore || 0) - Number(a.semanticScore || 0) ||
      Number(b.lexicalScore || 0) - Number(a.lexicalScore || 0) ||
      a.fileName.localeCompare(b.fileName) ||
      a.page - b.page
  );
  return entries.slice(0, Math.max(1, Math.min(Number(limit) || 20, 100)));
}

async function searchIndexedPdfHybrid(query, limit = 20, options = {}) {
  const normalizedQuery = normalizePdfText(String(query || ""));
  if (!normalizedQuery) {
    return {
      results: [],
      retrieval: {
        mode: "none",
        embeddingModel: "",
        note: "Empty query.",
      },
    };
  }

  const preferredFileName = String(options?.preferredFileName || "").trim();
  const restrictFileName = String(options?.restrictFileName || "").trim().toLowerCase();
  const config = normalizeAiConfig(options?.config || {});
  const allFiles = pdfIndexCache.files.filter((file) => {
    if (!restrictFileName) return true;
    return String(file?.fileName || "").trim().toLowerCase() === restrictFileName;
  });
  const lexicalMatches = collectLexicalChunkMatches(allFiles, normalizedQuery, Math.max(limit * 4, PDF_HYBRID_MATCH_LIMIT), preferredFileName);

  let embeddingModel = "";
  try {
    embeddingModel = await pickAvailableEmbeddingModel(config.endpoint, Math.min(15000, Number(config.timeoutMs || 15000)));
  } catch {
    embeddingModel = "";
  }
  if (!embeddingModel) {
    return {
      results: lexicalMatches.slice(0, limit).map((entry) => ({ ...entry, searchMode: "lexical" })),
      retrieval: {
        mode: "lexical",
        embeddingModel: "",
        note: "No local embedding model is installed. Search used keyword ranking only.",
      },
    };
  }

  let changed = false;
  try {
    changed = (await ensureFileEmbeddings(allFiles, config.endpoint, embeddingModel, Math.max(25000, Number(config.timeoutMs || 0)))) || changed;
    const queryEmbedding = (await requestOllamaEmbeddings(config.endpoint, embeddingModel, [normalizedQuery], Math.max(25000, Number(config.timeoutMs || 0))))[0];
    if (!queryEmbedding) {
      return {
        results: lexicalMatches.slice(0, limit).map((entry) => ({ ...entry, searchMode: "lexical" })),
        retrieval: {
          mode: "lexical",
          embeddingModel,
          note: `Embedding model "${embeddingModel}" did not return a query vector. Search used keyword ranking only.`,
        },
      };
    }

    const lexicalCandidateFiles = new Map();
    for (const match of lexicalMatches.slice(0, PDF_HYBRID_CANDIDATE_FILE_LIMIT * 3)) {
      if (!lexicalCandidateFiles.has(match.path)) {
        const file = allFiles.find((entry) => String(entry?.path || "").trim() === match.path);
        if (file) lexicalCandidateFiles.set(match.path, file);
      }
      if (lexicalCandidateFiles.size >= PDF_HYBRID_CANDIDATE_FILE_LIMIT) break;
    }
    const semanticFileCandidates = collectSemanticFileCandidates(allFiles, queryEmbedding, preferredFileName, PDF_HYBRID_CANDIDATE_FILE_LIMIT);
    const candidateFiles = new Map(semanticFileCandidates.map((entry) => [String(entry.file?.path || "").trim(), entry.file]));
    for (const [filePath, file] of lexicalCandidateFiles.entries()) {
      candidateFiles.set(filePath, file);
    }
    const candidateList = [...candidateFiles.values()];
    changed = (await ensureChunkEmbeddingsForFiles(candidateList, config.endpoint, embeddingModel, Math.max(30000, Number(config.timeoutMs || 0)))) || changed;
    if (changed) {
      await savePdfIndexCacheToDisk();
    }

    const semanticMatches = collectSemanticChunkMatches(candidateList, queryEmbedding, Math.max(limit * 4, PDF_HYBRID_MATCH_LIMIT), preferredFileName);
    const fused = fuseHybridMatches(lexicalMatches, semanticMatches, limit);
    return {
      results: fused,
      retrieval: {
        mode: semanticMatches.length && lexicalMatches.length ? "hybrid" : semanticMatches.length ? "semantic" : "lexical",
        embeddingModel,
        note:
          semanticMatches.length && lexicalMatches.length
            ? `Hybrid search combined keyword and semantic retrieval using ${embeddingModel}.`
            : semanticMatches.length
              ? `Semantic retrieval used ${embeddingModel}.`
              : `Keyword ranking remained stronger than semantic retrieval for this query.`,
      },
    };
  } catch (err) {
    if (changed) {
      await savePdfIndexCacheToDisk();
    }
    return {
      results: lexicalMatches.slice(0, limit).map((entry) => ({ ...entry, searchMode: "lexical" })),
      retrieval: {
        mode: "lexical",
        embeddingModel,
        note: `Semantic retrieval failed, so search fell back to keyword ranking only. ${String(err?.message || err || "")}`.trim(),
      },
    };
  }
}

function scoreTextAgainstQuery(textLower, searchParts) {
  const haystack = String(textLower || "");
  if (!haystack) return { score: 0, firstHit: -1 };

  let score = 0;
  let firstHit = -1;
  let matchedWords = 0;
  const phrase = String(searchParts?.phrase || "");
  const words = Array.isArray(searchParts?.words) ? searchParts.words : [];

  if (phrase) {
    const phraseHit = haystack.indexOf(phrase);
    if (phraseHit >= 0) {
      score += 16;
      firstHit = phraseHit;
    }
  }

  for (const word of words) {
    const hit = haystack.indexOf(word);
    if (hit < 0) continue;
    matchedWords += 1;
    if (firstHit < 0 || hit < firstHit) firstHit = hit;
    score += 4;
    score += Math.min(6, countOccurrences(haystack, word, 6));
  }

  if (matchedWords && words.length) {
    score += Math.round((matchedWords / words.length) * 8);
  }

  return {
    score,
    firstHit,
  };
}

function countOccurrences(haystack, needle, cap = 6) {
  if (!haystack || !needle) return 0;
  let count = 0;
  let cursor = 0;
  while (count < cap) {
    const hit = haystack.indexOf(needle, cursor);
    if (hit < 0) break;
    count += 1;
    cursor = hit + needle.length;
  }
  return count;
}

function makeSnippet(text, firstHit) {
  const clean = normalizePdfText(text);
  if (!clean) return "";

  const hit = Number.isFinite(firstHit) ? firstHit : -1;
  const start = Math.max(0, hit >= 0 ? hit - 120 : 0);
  const end = Math.min(clean.length, hit >= 0 ? hit + 220 : 260);
  const prefix = start > 0 ? "..." : "";
  const suffix = end < clean.length ? "..." : "";
  return `${prefix}${clean.slice(start, end)}${suffix}`;
}

function resolveIndexedPdfFile(payload) {
  const matchPath = String(payload?.path || "").trim();
  const matchName = String(payload?.fileName || "").trim().toLowerCase();
  if (matchPath) {
    const byPath = pdfIndexCache.files.find((file) => String(file?.path || "").trim() === matchPath);
    if (byPath) return byPath;
  }
  if (matchName) {
    const byName = pdfIndexCache.files.find((file) => String(file?.fileName || "").trim().toLowerCase() === matchName);
    if (byName) return byName;
  }
  return pdfIndexCache.files[0] || null;
}

function buildPdfSummaryResponse(file) {
  return {
    fileName: String(file?.fileName || "").trim(),
    path: String(file?.path || "").trim(),
    summary: normalizeStoredPdfSummaryText(file?.summary),
    summaryUpdatedAt: String(file?.summaryUpdatedAt || "").trim(),
  };
}

function buildPdfSummaryConfig(config, stage = "chunk") {
  const isFinalStage = String(stage || "").toLowerCase() === "final";
  const next = {
    ...config,
    temperature: Math.max(0, Math.min(Number(config?.temperature ?? 0.2), isFinalStage ? 0.25 : 0.35)),
    maxOutputTokens: Math.max(Number(config?.maxOutputTokens || 0), isFinalStage ? 1400 : 520),
  };
  let timeoutSec = Number.parseInt(String(config?.timeoutSec || "120"), 10);
  if (!Number.isFinite(timeoutSec)) timeoutSec = 120;
  timeoutSec = Math.max(timeoutSec, isFinalStage ? 420 : 240);
  if (/20b/i.test(String(next.model || ""))) {
    timeoutSec = Math.max(timeoutSec, isFinalStage ? 480 : 360);
  }
  next.timeoutSec = timeoutSec;
  next.timeoutMs = timeoutSec * 1000;
  return next;
}

function splitTextIntoChunks(text, chunkSize = 7600, maxChunks = 6) {
  const clean = normalizePdfText(text);
  if (!clean) return [];
  const out = [];
  let cursor = 0;
  while (cursor < clean.length && out.length < Math.max(1, maxChunks)) {
    let end = Math.min(clean.length, cursor + Math.max(1200, chunkSize));
    if (end < clean.length) {
      const pivot = clean.lastIndexOf(" ", end);
      if (pivot > cursor + 900) end = pivot;
    }
    const piece = clean.slice(cursor, end).trim();
    if (piece) out.push(piece);
    if (end <= cursor) break;
    cursor = end;
  }
  return out.length ? out : [clean.slice(0, Math.max(1200, chunkSize))];
}

function buildPdfChunkSummaryPrompt({ fileName, chunkText, index, total }) {
  return [
    `You are summarizing indexed PDF content for DM Helper, a GM prep tool.`,
    `Book: ${fileName}`,
    `Chunk ${index} of ${total}.`,
    `Task: extract the most useful GM-facing facts from this chunk only.`,
    `Return these headings exactly:`,
    `Adventure Beats:`,
    `- bullet`,
    `Key People / Factions:`,
    `- bullet`,
    `Key Places / Scenes:`,
    `- bullet`,
    `Threats / Obstacles / Clues:`,
    `- bullet`,
    `Rules / Mechanics Worth Prep:`,
    `- bullet`,
    `GM Use:`,
    `- bullet`,
    `Keep it factual and grounded only in the provided chunk.`,
    `No markdown tables. No bold. No numbering. Prefer short bullets over paragraphs.`,
    `If a section has nothing useful, write "- None noted in this chunk."`,
    ``,
    `Chunk text:`,
    chunkText,
  ].join("\n");
}

function buildPdfFinalSummaryPrompt({ fileName, chunkSummaries }) {
  return [
    `You are combining chunk summaries into one persistent GM-ready book brief for DM Helper.`,
    `Book: ${fileName}`,
    `Return these headings exactly:`,
    `Adventure Premise:`,
    `- 2 to 4 bullets`,
    `Main Threats / Stakes:`,
    `- 3 to 6 bullets`,
    `Key People / Factions:`,
    `- 3 to 8 bullets`,
    `Key Places / Scenes:`,
    `- 3 to 8 bullets`,
    `Likely Flow / Structure:`,
    `- 3 to 6 bullets`,
    `What To Prep First:`,
    `- 5 to 8 bullets`,
    `Fast Table Reference:`,
    `- 4 to 8 bullets`,
    `Keep it concise, factual, and immediately useful at the table.`,
    `No markdown tables. No bold. No numbering. No incomplete sentences. Prefer bullets over paragraphs.`,
    ``,
    `Chunk summaries:`,
    ...chunkSummaries.map((summary, i) => `Chunk ${i + 1}:\n${summary}`),
  ].join("\n");
}

function buildPdfFinalSummaryRetryPrompt({ fileName, chunkSummaries }) {
  return [
    `You are fixing a weak or incomplete summary for DM Helper.`,
    `Book: ${fileName}`,
    `Return only a clean GM summary in this exact structure:`,
    `Adventure Premise:`,
    `- bullet`,
    `Main Threats / Stakes:`,
    `- bullet`,
    `Key People / Factions:`,
    `- bullet`,
    `Key Places / Scenes:`,
    `- bullet`,
    `Likely Flow / Structure:`,
    `- bullet`,
    `What To Prep First:`,
    `- bullet`,
    `Fast Table Reference:`,
    `- bullet`,
    `Requirements:`,
    `- use bullets only`,
    `- no markdown tables`,
    `- no bold`,
    `- no numbered lists`,
    `- no section may be empty`,
    `- if a detail is unclear, keep it general rather than inventing specifics`,
    ``,
    `Chunk summaries:`,
    ...chunkSummaries.map((summary, i) => `Chunk ${i + 1}:\n${summary}`),
  ].join("\n");
}

function extractiveChunkSummary(chunkText) {
  const clean = normalizePdfText(chunkText);
  if (!clean) return "";
  const sentences = clean.match(/[^.!?]+[.!?]?/g) || [];
  const picks = sentences.map((s) => s.trim()).filter(Boolean).slice(0, 5);
  const lines = picks.map((s) => `- ${s.endsWith(".") ? s : `${s}.`}`);
  return lines.length ? lines.join("\n") : `- ${clean.slice(0, 280)}...`;
}

function collectBulletsFromPdfSummarySection(text, label, fallbackLabels = []) {
  const labels = [label, ...fallbackLabels];
  const picked = [];
  const seen = new Set();
  for (const item of labels) {
    const block = extractLabeledBlock(text, item);
    if (!block) continue;
    for (const line of splitAiLines(block)) {
      const clean = String(line || "").replace(/^[-*]\s*/, "").trim();
      if (!clean || /none noted in this chunk/i.test(clean)) continue;
      const key = clean.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      picked.push(clean.endsWith(".") ? clean : `${clean}.`);
    }
  }
  return picked;
}

function takePdfSummaryBullets(lines, minCount, maxCount, fallbackLines = []) {
  const unique = [];
  const seen = new Set();
  for (const line of [...(lines || []), ...(fallbackLines || [])]) {
    const clean = normalizeSentenceText(line);
    if (!clean) continue;
    const key = clean.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(clean.endsWith(".") ? clean : `${clean}.`);
  }
  if (!unique.length) return [];
  const min = Math.max(0, Number(minCount) || 0);
  const max = Math.max(min || 1, Number(maxCount) || min || 1);
  if (unique.length >= min) return unique.slice(0, max);
  return unique.slice(0, max);
}

function isUsefulPdfSummary(text) {
  const raw = String(text || "").trim();
  if (!raw) return false;
  if (isClearlyTruncatedOutput(raw)) return false;
  if (/\|[^\n]{0,220}\|/.test(raw)) return false;
  const required = [
    "Adventure Premise",
    "Main Threats / Stakes",
    "Key People / Factions",
    "Key Places / Scenes",
    "Likely Flow / Structure",
    "What To Prep First",
    "Fast Table Reference",
  ];
  const headingCount = required.filter((label) => new RegExp(`^${escapeRegex(label)}\\s*:`, "im").test(raw)).length;
  if (headingCount < 5) return false;
  if (countBulletLikeLines(raw) < 10) return false;
  return raw.length >= 320;
}

function combineChunkSummariesFallback(fileName, chunkSummaries) {
  const summaries = chunkSummaries.map((item) => String(item || "")).filter(Boolean);
  if (!summaries.length) {
    return [
      `Adventure Premise:`,
      `- ${fileName} was indexed, but no usable summary could be generated from the saved chunk notes.`,
      `Main Threats / Stakes:`,
      `- Re-run summary or use PDF search for focused book details.`,
      `Key People / Factions:`,
      `- Use PDF search to identify named NPCs and factions.`,
      `Key Places / Scenes:`,
      `- Use PDF search to locate the opening area, main sites, and likely set pieces.`,
      `Likely Flow / Structure:`,
      `- Review the saved PDF search results for chapter or encounter order.`,
      `What To Prep First:`,
      `- Summarize the selected PDF again after confirming the indexed text looks clean.`,
      `Fast Table Reference:`,
      `- Keep PDF Intel open for exact page lookups.`,
    ].join("\n");
  }

  const adventureBeats = takePdfSummaryBullets(
    summaries.flatMap((summary) => collectBulletsFromPdfSummarySection(summary, "Adventure Beats")),
    2,
    6
  );
  const people = takePdfSummaryBullets(
    summaries.flatMap((summary) => collectBulletsFromPdfSummarySection(summary, "Key People / Factions")),
    3,
    8
  );
  const places = takePdfSummaryBullets(
    summaries.flatMap((summary) => collectBulletsFromPdfSummarySection(summary, "Key Places / Scenes")),
    3,
    8
  );
  const threats = takePdfSummaryBullets(
    summaries.flatMap((summary) => collectBulletsFromPdfSummarySection(summary, "Threats / Obstacles / Clues")),
    3,
    8
  );
  const rules = takePdfSummaryBullets(
    summaries.flatMap((summary) => collectBulletsFromPdfSummarySection(summary, "Rules / Mechanics Worth Prep")),
    2,
    6
  );
  const gmUse = takePdfSummaryBullets(
    summaries.flatMap((summary) => collectBulletsFromPdfSummarySection(summary, "GM Use")),
    3,
    8
  );

  return [
    "Adventure Premise:",
    ...takePdfSummaryBullets(adventureBeats, 1, 4, threats).map((line) => `- ${line}`),
    "Main Threats / Stakes:",
    ...takePdfSummaryBullets(threats, 1, 6, adventureBeats).map((line) => `- ${line}`),
    "Key People / Factions:",
    ...takePdfSummaryBullets(people, 1, 8, adventureBeats).map((line) => `- ${line}`),
    "Key Places / Scenes:",
    ...takePdfSummaryBullets(places, 1, 8, adventureBeats).map((line) => `- ${line}`),
    "Likely Flow / Structure:",
    ...takePdfSummaryBullets(adventureBeats, 1, 6, places).map((line) => `- ${line}`),
    "What To Prep First:",
    ...takePdfSummaryBullets(gmUse, 1, 8, threats.concat(rules)).map((line) => `- ${line}`),
    "Fast Table Reference:",
    ...takePdfSummaryBullets(rules.concat(people.slice(0, 2), places.slice(0, 2)), 1, 8, gmUse).map((line) => `- ${line}`),
  ].join("\n");
}

async function openPdfAtPageInApp(targetPath, page) {
  const safePage = Math.max(1, page);
  const fileUrl = pathToFileURL(targetPath).toString();
  const viewerWin = new BrowserWindow({
    width: 1200,
    height: 900,
    minWidth: 960,
    minHeight: 680,
    autoHideMenuBar: true,
    title: `${path.basename(targetPath)} - Page ${Math.max(1, page)}`,
    webPreferences: {
      preload: path.join(__dirname, "viewer-preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: false,
    },
  });

  pdfViewerWindows.add(viewerWin);
  viewerWin.on("closed", () => {
    pdfViewerWindows.delete(viewerWin);
  });

  await viewerWin.loadFile(path.join(__dirname, "pdf-viewer.html"), {
    query: {
      targetPath: targetPath,
      fileUrl,
      page: String(safePage),
    },
  });
}

function wireSpellcheckContextMenu(win) {
  win.webContents.on("context-menu", (_event, params) => {
    const menu = new Menu();

    if (params.misspelledWord && params.dictionarySuggestions.length) {
      for (const suggestion of params.dictionarySuggestions.slice(0, 6)) {
        menu.append({
          label: suggestion,
          click: () => win.webContents.replaceMisspelling(suggestion),
        });
      }
      menu.append({ type: "separator" });
      menu.append({
        label: `Add "${params.misspelledWord}" to dictionary`,
        click: () => win.webContents.session.addWordToSpellCheckerDictionary(params.misspelledWord),
      });
      menu.append({ type: "separator" });
    }

    menu.append({ role: "undo" });
    menu.append({ role: "redo" });
    menu.append({ type: "separator" });
    menu.append({ role: "cut" });
    menu.append({ role: "copy" });
    menu.append({ role: "paste" });
    menu.append({ role: "selectAll" });

    menu.popup();
  });
}

function normalizeAiConfig(rawConfig) {
  const endpointRaw = String(rawConfig?.endpoint || DEFAULT_AI_ENDPOINT).trim();
  const endpoint = endpointRaw.replace(/\/+$/g, "");
  const model = String(rawConfig?.model || DEFAULT_AI_MODEL).trim() || DEFAULT_AI_MODEL;
  const tempRaw = Number.parseFloat(String(rawConfig?.temperature ?? "0.2"));
  const temperature = Number.isFinite(tempRaw) ? Math.max(0, Math.min(tempRaw, 2)) : 0.2;
  const usePdfContext = rawConfig?.usePdfContext === false ? false : true;
  const compactContext = rawConfig?.compactContext === false ? false : true;
  const maxOutputTokensRaw = Number.parseInt(String(rawConfig?.maxOutputTokens ?? "320"), 10);
  const maxOutputTokens = Number.isFinite(maxOutputTokensRaw)
    ? Math.max(64, Math.min(maxOutputTokensRaw, 2048))
    : 320;
  const timeoutSecRaw = Number.parseInt(String(rawConfig?.timeoutSec ?? "120"), 10);
  let timeoutSec = Number.isFinite(timeoutSecRaw) ? Math.max(15, Math.min(timeoutSecRaw, 1200)) : 120;
  if (/20b/i.test(model) && timeoutSec < 300) {
    timeoutSec = 300;
  }
  return {
    endpoint: endpoint || DEFAULT_AI_ENDPOINT,
    model,
    temperature,
    usePdfContext,
    compactContext,
    maxOutputTokens,
    timeoutSec,
    timeoutMs: timeoutSec * 1000 || AI_TIMEOUT_MS,
  };
}

async function testOllamaConnection(config) {
  const modelNames = await listOllamaModelsWithRecovery(config.endpoint, Math.min(15000, Number(config?.timeoutMs) || 15000));
  const hasModel = modelNames.includes(config.model);
  const modelLabel = getAiModelDisplayName(config.model);
  return {
    ok: true,
    modelFound: hasModel,
    models: modelNames.slice(0, 80),
    message: hasModel
      ? `Connected. Model "${modelLabel}" is available.`
      : `Connected. Model "${modelLabel}" not found locally yet.`,
  };
}

function isLocalEndpoint(endpoint) {
  try {
    const parsed = new URL(String(endpoint || ""));
    const host = String(parsed.hostname || "").toLowerCase();
    return host === "127.0.0.1" || host === "localhost" || host === "::1";
  } catch {
    return false;
  }
}

function isLikelyOllamaConnectionError(err) {
  const msg = String(err?.message || err || "").toLowerCase();
  return (
    msg.includes("could not reach local ai endpoint") ||
    msg.includes("could not connect to local ai") ||
    msg.includes("econnrefused") ||
    msg.includes("fetch failed")
  );
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function launchOllamaServe() {
  const child = spawn("ollama", ["serve"], {
    detached: true,
    stdio: "ignore",
    windowsHide: true,
    shell: false,
  });
  child.unref();
}

async function ensureOllamaAvailable(endpoint, timeoutMs = 10000) {
  if (ollamaBootPromise) {
    await ollamaBootPromise;
    return;
  }

  ollamaBootPromise = (async () => {
    try {
      launchOllamaServe();
    } catch {
      // Ignore launch errors; readiness probe below will still validate status.
    }

    const probeTimeout = Math.max(1500, Math.min(5000, Number(timeoutMs) || 2500));
    for (let i = 0; i < OLLAMA_BOOT_RETRY_COUNT; i += 1) {
      try {
        await listOllamaModels(endpoint, probeTimeout);
        return;
      } catch (probeErr) {
        if (!isLikelyOllamaConnectionError(probeErr)) throw probeErr;
        await sleep(OLLAMA_BOOT_RETRY_DELAY_MS);
      }
    }

    throw new Error(
      "Could not reach local AI endpoint after restart attempt. Confirm Ollama is installed and running."
    );
  })();

  try {
    await ollamaBootPromise;
  } finally {
    ollamaBootPromise = null;
  }
}

async function listOllamaModelsWithRecovery(endpoint, timeoutMs = 10000) {
  try {
    return await listOllamaModels(endpoint, timeoutMs);
  } catch (err) {
    if (!isLocalEndpoint(endpoint) || !isLikelyOllamaConnectionError(err)) throw err;
    await ensureOllamaAvailable(endpoint, timeoutMs);
    return listOllamaModels(endpoint, Math.max(10000, Number(timeoutMs) || 10000));
  }
}

async function listOllamaModels(endpoint, timeoutMs = 10000) {
  const safeTimeoutMs = Number(timeoutMs) > 0 ? Number(timeoutMs) : 10000;
  const url = `${endpoint}/api/tags`;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), safeTimeoutMs);
  let response = null;
  try {
    response = await fetch(url, { signal: controller.signal });
  } catch (err) {
    if (err?.name === "AbortError") {
      throw new Error(`Could not reach local AI endpoint (timeout after ${Math.round(safeTimeoutMs / 1000)}s).`);
    }
    throw err;
  } finally {
    clearTimeout(timeout);
  }
  if (!response.ok) {
    throw new Error(`Could not reach local AI endpoint (${response.status}).`);
  }
  const data = await response.json();
  const modelNames = Array.isArray(data?.models)
    ? data.models.map((model) => String(model?.name || "").trim()).filter(Boolean)
    : [];
  return modelNames.slice(0, 160);
}

async function generateWithOllama(config, userPrompt, recovered = false, timeoutRetried = false) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), Number(config?.timeoutMs) || AI_TIMEOUT_MS);
  try {
    const chatData = await requestOllamaJson(
      `${config.endpoint}/api/chat`,
      {
        model: config.model,
        stream: false,
        think: false,
        options: {
          temperature: config.temperature,
          num_predict: config.maxOutputTokens,
        },
        messages: [
          {
            role: "system",
            content:
              "You are a tabletop GM writing assistant. Return practical, complete GM-facing text only. Match requested structure exactly and never omit requested fields.",
          },
          { role: "user", content: userPrompt },
        ],
      },
      controller.signal
    );

    let text = extractOllamaText(chatData);
    if (!text) {
      const generateData = await requestOllamaJson(
        `${config.endpoint}/api/generate`,
        {
          model: config.model,
          stream: false,
          options: {
            temperature: config.temperature,
            num_predict: config.maxOutputTokens,
          },
          prompt: [
            "You are a tabletop GM writing assistant.",
            "Return practical, complete GM-facing text only.",
            "Match requested structure exactly and never omit requested fields.",
            "",
            userPrompt,
          ].join("\n"),
        },
        controller.signal
      );
      text = extractOllamaText(generateData);
    }

    // Some local models occasionally return no final content from both endpoints.
    return text;
  } catch (err) {
    if (err?.name === "AbortError") {
      const timeoutMs = Number(config?.timeoutMs) || AI_TIMEOUT_MS;
      const modelName = String(config?.model || "");
      const canRetryWithLongerTimeout = !timeoutRetried && /20b/i.test(modelName) && timeoutMs < 600000;
      if (canRetryWithLongerTimeout) {
        const retriedTimeoutMs = Math.min(600000, Math.max(timeoutMs * 2, 420000));
        const retriedConfig = {
          ...config,
          timeoutMs: retriedTimeoutMs,
          timeoutSec: Math.max(Number(config?.timeoutSec) || 0, Math.round(retriedTimeoutMs / 1000)),
          maxOutputTokens: Math.max(192, Math.min(Number(config?.maxOutputTokens || 320), 1024)),
        };
        return generateWithOllama(retriedConfig, userPrompt, recovered, true);
      }
      throw new Error(
        `Local AI request timed out after ${Math.round(timeoutMs / 1000)}s. Try compact context, fewer output tokens, a faster model, or a longer timeout.`
      );
    }
    if (!recovered && isLocalEndpoint(config.endpoint) && isLikelyOllamaConnectionError(err)) {
      await ensureOllamaAvailable(config.endpoint, Number(config?.timeoutMs) || AI_TIMEOUT_MS);
      return generateWithOllama(config, userPrompt, true, timeoutRetried);
    }
    const message = String(err?.message || "");
    if (/fetch failed/i.test(message) || /ECONNREFUSED/i.test(message)) {
      throw new Error(`Could not connect to local AI at ${config.endpoint}. Confirm Ollama is running and endpoint is correct.`);
    }
    throw err;
  } finally {
    clearTimeout(timeout);
  }
}

async function requestOllamaJson(url, payload, signal) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    signal,
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const details = (await response.text()).slice(0, 260);
    throw new Error(`Local AI generation failed (${response.status}): ${details}`);
  }
  const data = await response.json();
  if (typeof data?.error === "string" && data.error.trim()) {
    throw new Error(`Local AI generation failed: ${data.error.trim()}`);
  }
  return data;
}

function extractOllamaText(data) {
  const direct = String(data?.message?.content || "").trim();
  if (direct) return direct;

  const contentParts = Array.isArray(data?.message?.content) ? data.message.content : [];
  if (contentParts.length) {
    const joined = contentParts
      .map((part) => {
        if (typeof part === "string") return part;
        if (part && typeof part.text === "string") return part.text;
        return "";
      })
      .join("\n")
      .trim();
    if (joined) return joined;
  }

  const alternatives = [data?.response, data?.output_text, data?.completion, data?.text];
  for (const candidate of alternatives) {
    const clean = String(candidate || "").trim();
    if (clean) return clean;
  }

  return "";
}

function buildAiUserPrompt({ mode, input, context, compactContext = true }) {
  const latestSession = context?.latestSession || {};
  const recentSessions = Array.isArray(context?.recentSessions) ? context.recentSessions : [];
  const openQuests = Array.isArray(context?.openQuests) ? context.openQuests : [];
  const quests = Array.isArray(context?.quests) ? context.quests : [];
  const npcs = Array.isArray(context?.npcs) ? context.npcs : [];
  const locations = Array.isArray(context?.locations) ? context.locations : [];
  const kingdom = context?.kingdom || null;
  const selectedPdfFile = summarizeForPrompt(String(context?.selectedPdfFile || ""), 120);
  const selectedPdfSummary = summarizeForPrompt(String(context?.selectedPdfSummary || ""), compactContext ? 720 : 1100);
  const selectedPdfPreview = summarizeForPrompt(String(context?.selectedPdfPreview || ""), compactContext ? 720 : 1100);
  const indexedPdfFiles = Array.isArray(context?.pdfIndexedFiles) ? context.pdfIndexedFiles : [];
  const pdfSummaryBriefs = Array.isArray(context?.pdfSummaryBriefs) ? context.pdfSummaryBriefs : [];
  const indexedPdfCount = Number.parseInt(String(context?.pdfIndexedFileCount || indexedPdfFiles.length || 0), 10) || 0;
  const aiHistory = Array.isArray(context?.aiHistory) ? context.aiHistory : [];
  const activeTab = summarizeForPrompt(String(context?.activeTab || ""), 40);
  const tabLabel = summarizeForPrompt(String(context?.tabLabel || ""), 80);
  const limits = compactContext
    ? { draft: 1500, tab: 900, latest: 220, snippet: 180 }
    : { draft: 2400, tab: 1800, latest: 360, snippet: 280 };
  const tabContext = summarizeForPrompt(String(context?.tabContext || ""), limits.tab);
  const pdfSnippets = Array.isArray(context?.pdfSnippets) ? context.pdfSnippets : [];
  const pdfEnabled = context?.pdfContextEnabled === true;
  const appRoleLines = getDmHelperAppRoleLines(activeTab);
  const recentSessionLines = summarizeRecentSessionsForPrompt(recentSessions, compactContext ? 3 : 5);
  const trackedNpcLines = summarizeTrackedNpcsForPrompt(npcs, compactContext ? 5 : 8);
  const trackedQuestLines = summarizeTrackedQuestsForPrompt(quests, compactContext ? 5 : 8);
  const trackedLocationLines = summarizeTrackedLocationsForPrompt(locations, compactContext ? 5 : 8);
  const kingdomLines = summarizeKingdomForPrompt(kingdom, compactContext);
  const historyLimit = compactContext ? 6 : 10;
  const historyTurns = aiHistory
    .slice(-historyLimit)
    .map((turn) => {
      const role = String(turn?.role || "").toLowerCase() === "assistant" ? "AI" : "GM";
      const tab = summarizeForPrompt(String(turn?.tabId || ""), 18);
      const text = summarizeForPrompt(String(turn?.text || ""), compactContext ? 160 : 240);
      return `${role}${tab ? ` [${tab}]` : ""}: ${text}`;
    })
    .filter((line) => line.endsWith(": ") === false);

  const modeGuide = {
    assistant: "Answer the GM's question directly with practical, table-ready guidance.",
    session:
      "Produce structured, table-ready session notes. Follow requested section labels exactly and provide substantive detail (not just one short paragraph).",
    recap: "Rewrite as a short read-aloud recap for players (3-6 sentences).",
    npc: "Produce one table-ready NPC using the requested labels exactly. Give the NPC a clear motive, pressure, and immediately playable detail.",
    quest: "Rewrite as a quest objective with stakes and next actionable beat.",
    location: "Rewrite as a location briefing with atmosphere and immediate tension.",
    prep: "Rewrite as next-session prep bullet points.",
  };
  const modeSpecificLines = getModeSpecificPromptLines(mode, input, context);

  const lines = [
    `Mode: ${mode}`,
    `Goal: ${modeGuide[mode] || modeGuide.session}`,
    "",
    ...appRoleLines,
    ...(modeSpecificLines.length ? ["", ...modeSpecificLines] : []),
    "",
    "Draft input:",
    summarizeForPrompt(input, limits.draft),
    "",
    "Campaign context:",
    `Active tab: ${activeTab || "unknown"} (${tabLabel || "unknown"})`,
    `Tab context: ${tabContext || "None provided."}`,
    `Latest session: ${summarizeForPrompt(String(latestSession?.title || ""), 120)} | ${summarizeForPrompt(
      String(latestSession?.summary || ""),
      limits.latest
    )}`,
    ...(recentSessionLines.length ? ["Recent sessions in app:", ...recentSessionLines] : []),
    `Open quests: ${openQuests.map((q) => summarizeForPrompt(String(q?.title || ""), 80)).join("; ") || "None listed."}`,
    ...(trackedQuestLines.length ? ["Tracked quest records:", ...trackedQuestLines] : ["Tracked quest records: None listed."]),
    ...(trackedNpcLines.length ? ["Tracked NPC records:", ...trackedNpcLines] : ["Tracked NPC records: None listed."]),
    ...(trackedLocationLines.length ? ["Tracked location records:", ...trackedLocationLines] : ["Tracked location records: None listed."]),
    ...(kingdomLines.length ? ["Kingdom records:", ...kingdomLines] : []),
    `PDF context enabled: ${pdfEnabled ? "yes" : "no"}`,
    `Indexed PDF files (${indexedPdfCount}): ${
      indexedPdfFiles.length
        ? indexedPdfFiles.map((name) => summarizeForPrompt(String(name || ""), 70)).join("; ")
        : "None indexed."
    }`,
    `Selected PDF focus: ${selectedPdfFile || "None selected."}`,
    ...(selectedPdfSummary ? [`Selected PDF summary: ${selectedPdfSummary}`] : []),
    ...(selectedPdfPreview ? [`Selected PDF preview: ${selectedPdfPreview}`] : []),
    ...(pdfSummaryBriefs.length
      ? [
          "Saved PDF memory briefs:",
          ...pdfSummaryBriefs.map(
            (entry, idx) =>
              `${idx + 1}. ${summarizeForPrompt(String(entry?.fileName || "PDF"), 80)} - ${summarizeForPrompt(
                String(entry?.summary || ""),
                limits.snippet * 2
              )}`
          ),
        ]
      : ["Saved PDF memory briefs: None yet."]),
    "",
    ...(historyTurns.length
      ? [
          "Recent AI conversation:",
          ...historyTurns.map((line, index) => `${index + 1}. ${line}`),
          "",
        ]
      : []),
    ...(pdfSnippets.length
      ? [
          "Relevant PDF excerpts:",
          ...pdfSnippets.map(
            (snippet, index) =>
              `${index + 1}. [${summarizeForPrompt(String(snippet?.fileName || "PDF"), 80)}${
                snippet?.page ? ` p.${snippet.page}` : ""
              }] ${summarizeForPrompt(String(snippet?.snippet || ""), limits.snippet)}`
          ),
          "",
        ]
      : []),
    "Return only final GM-facing content.",
    "No instruction lists, no policy text, and no meta-rules in the answer.",
    "Keep facts grounded in provided context and avoid invented lore.",
    "Source scope rule: you only have access to campaign context, the active app-bundled kingdom rules profile if one is provided above, and indexed PDF files listed above.",
    "If asked what books/sources/rules/PDFs you can access, answer using only the campaign data above, the active kingdom rules profile if present, and the indexed PDF file names.",
    "Never claim access to external books, websites, or rules not listed in indexed PDF files or the active kingdom rules profile.",
  ];

  return lines.join("\n");
}

function getDmHelperAppRoleLines(activeTab) {
  const tabId = String(activeTab || "").toLowerCase();
  const workflowByTab = {
    dashboard: "You are planning the GM's next moves. Output should be ready to attach to the latest session prep in the app.",
    sessions: "You are helping maintain the session record. Output should be ready to apply into the latest session summary or next-prep fields.",
    capture: "You are cleaning raw table notes into structured records the app can keep as live capture or session notes.",
    writing: "You are drafting clean GM-facing text that the writing helper can store or paste into campaign notes.",
    kingdom: "You are helping run a PF2e Kingmaker kingdom inside DM Helper. Output should be ready to save into kingdom notes, turn logs, settlement plans, or leader assignments.",
    npcs: "You are creating or enriching structured NPC records for DM Helper. Output may be imported directly into NPC entries.",
    quests: "You are creating or refining structured quest records for DM Helper. Output may be imported directly into quest entries.",
    locations: "You are creating or refining structured location records for DM Helper. Output may be imported directly into location entries.",
    pdf: "You are answering from indexed PDF context or producing a PDF query the app can run immediately.",
    foundry: "You are preparing handoff/export notes that the GM can use in Foundry and store in session prep.",
  };
  return [
    "App role: You are Loremaster inside the DM Helper app, helping the GM build and maintain structured campaign records and usable session prep.",
    `Current app workflow: ${workflowByTab[tabId] || "You are helping the GM create content that can be saved back into DM Helper without cleanup."}`,
    "Write content that fits the app workflow for the active tab and is ready to save or apply inside DM Helper.",
  ];
}

function summarizeRecentSessionsForPrompt(sessions, limit = 4) {
  return (sessions || [])
    .slice(0, Math.max(0, Number(limit) || 0))
    .map((session, index) => {
      const title = summarizeForPrompt(String(session?.title || ""), 70) || `Session ${index + 1}`;
      const date = summarizeForPrompt(String(session?.date || ""), 24);
      const arc = summarizeForPrompt(String(session?.arc || ""), 50);
      const summary = summarizeForPrompt(String(session?.summary || ""), 120);
      return `- ${title}${date ? ` (${date})` : ""}${arc ? ` | arc: ${arc}` : ""} | ${summary || "No summary yet."}`;
    })
    .filter(Boolean);
}

function summarizeTrackedNpcsForPrompt(npcs, limit = 6) {
  return (npcs || [])
    .slice(0, Math.max(0, Number(limit) || 0))
    .map((npc) => {
      const name = summarizeForPrompt(String(npc?.name || ""), 60);
      if (!name) return "";
      const parts = [
        summarizeForPrompt(String(npc?.role || ""), 40),
        summarizeForPrompt(String(npc?.agenda || ""), 70),
        summarizeForPrompt(String(npc?.disposition || ""), 32),
      ].filter(Boolean);
      const notes = summarizeForPrompt(String(npc?.notes || ""), 120);
      return `- ${name}${parts.length ? ` | ${parts.join(" | ")}` : ""}${notes ? ` | notes: ${notes}` : ""}`;
    })
    .filter(Boolean);
}

function summarizeTrackedQuestsForPrompt(quests, limit = 6) {
  return (quests || [])
    .slice(0, Math.max(0, Number(limit) || 0))
    .map((quest) => {
      const title = summarizeForPrompt(String(quest?.title || ""), 70);
      if (!title) return "";
      const status = summarizeForPrompt(String(quest?.status || ""), 24);
      const objective = summarizeForPrompt(String(quest?.objective || ""), 90);
      const stakes = summarizeForPrompt(String(quest?.stakes || ""), 110);
      const giver = summarizeForPrompt(String(quest?.giver || ""), 48);
      return `- ${title}${status ? ` (${status})` : ""}${giver ? ` | giver: ${giver}` : ""}${objective ? ` | objective: ${objective}` : ""}${stakes ? ` | stakes: ${stakes}` : ""}`;
    })
    .filter(Boolean);
}

function summarizeTrackedLocationsForPrompt(locations, limit = 6) {
  return (locations || [])
    .slice(0, Math.max(0, Number(limit) || 0))
    .map((location) => {
      const name = summarizeForPrompt(String(location?.name || ""), 70);
      if (!name) return "";
      const hex = summarizeForPrompt(String(location?.hex || ""), 24);
      const whatChanged = summarizeForPrompt(String(location?.whatChanged || ""), 90);
      const notes = summarizeForPrompt(String(location?.notes || ""), 110);
      return `- ${name}${hex ? ` [${hex}]` : ""}${whatChanged ? ` | changed: ${whatChanged}` : ""}${notes ? ` | notes: ${notes}` : ""}`;
    })
    .filter(Boolean);
}

function summarizeKingdomForPrompt(kingdom, compactContext = true) {
  const data = kingdom && typeof kingdom === "object" ? kingdom : null;
  if (!data) return [];
  const profile = data?.rulesProfile || getDefaultKingdomRulesProfile();
  const leaders = Array.isArray(data?.leaders) ? data.leaders : [];
  const settlements = Array.isArray(data?.settlements) ? data.settlements : [];
  const regions = Array.isArray(data?.regions) ? data.regions : [];
  const recentTurns = Array.isArray(data?.recentTurns) ? data.recentTurns : [];
  const commodities = data?.commodities || {};
  const ruin = data?.ruin || {};
  const lines = [
    `- Sheet: ${summarizeForPrompt(String(data?.name || "Unnamed kingdom"), 90)} | turn ${summarizeForPrompt(String(data?.currentTurnLabel || "not set"), 36)} | level ${Number.parseInt(String(data?.level || "1"), 10) || 1} | size ${Number.parseInt(String(data?.size || "1"), 10) || 1} | Control DC ${Number.parseInt(String(data?.controlDC || "14"), 10) || 14}`,
    `- Economy: RP ${Number.parseInt(String(data?.resourcePoints || "0"), 10) || 0} | resource die ${summarizeForPrompt(String(data?.resourceDie || "d4"), 8)} | consumption ${Math.max(0, Number.parseInt(String(data?.consumption || "0"), 10) || 0)} | commodities F:${Number.parseInt(String(commodities.food || "0"), 10) || 0} L:${Number.parseInt(String(commodities.lumber || "0"), 10) || 0} Lux:${Number.parseInt(String(commodities.luxuries || "0"), 10) || 0} O:${Number.parseInt(String(commodities.ore || "0"), 10) || 0} S:${Number.parseInt(String(commodities.stone || "0"), 10) || 0}`,
    `- Pressure: unrest ${Math.max(0, Number.parseInt(String(data?.unrest || "0"), 10) || 0)} | renown ${Math.max(0, Number.parseInt(String(data?.renown || "0"), 10) || 0)} | fame ${Math.max(0, Number.parseInt(String(data?.fame || "0"), 10) || 0)} | infamy ${Math.max(0, Number.parseInt(String(data?.infamy || "0"), 10) || 0)} | ruin C:${Math.max(0, Number.parseInt(String(ruin.corruption || "0"), 10) || 0)} Cr:${Math.max(0, Number.parseInt(String(ruin.crime || "0"), 10) || 0)} D:${Math.max(0, Number.parseInt(String(ruin.decay || "0"), 10) || 0)} S:${Math.max(0, Number.parseInt(String(ruin.strife || "0"), 10) || 0)} / threshold ${Math.max(1, Number.parseInt(String(ruin.threshold || "5"), 10) || 5)}`,
    `- Rules profile: ${summarizeForPrompt(String(profile?.label || "Kingdom profile"), 90)} | ${summarizeForPrompt(String(profile?.summary || ""), compactContext ? 180 : 260)}`,
  ];
  const turnStructure = Array.isArray(profile?.turnStructure) ? profile.turnStructure : [];
  if (turnStructure.length) {
    lines.push(
      `- Turn structure: ${turnStructure
        .slice(0, compactContext ? 4 : 5)
        .map((entry) => `${summarizeForPrompt(String(entry || ""), compactContext ? 60 : 90)}`)
        .join(" | ")}`
    );
  }
  const aiSummary = Array.isArray(profile?.aiSummary)
    ? profile.aiSummary
    : Array.isArray(profile?.aiContextSummary)
      ? profile.aiContextSummary
      : [];
  if (aiSummary.length) {
    lines.push(
      `- Kingdom AI guide: ${aiSummary
        .slice(0, compactContext ? 4 : 6)
        .map((entry) => summarizeForPrompt(String(entry || ""), compactContext ? 80 : 120))
        .join(" | ")}`
    );
  }
  const leaderLines = leaders
    .slice(0, compactContext ? 5 : 8)
    .map((leader) => {
      const name = summarizeForPrompt(String(leader?.name || ""), 40);
      const role = summarizeForPrompt(String(leader?.role || ""), 28);
      const type = summarizeForPrompt(String(leader?.type || ""), 8);
      const skills = summarizeForPrompt(String(leader?.specializedSkills || ""), compactContext ? 60 : 90);
      return `- Leader: ${role || "Role"} = ${name || "Unassigned"}${type ? ` (${type})` : ""}${skills ? ` | specialized: ${skills}` : ""}`;
    })
    .filter(Boolean);
  const settlementLines = settlements
    .slice(0, compactContext ? 4 : 6)
    .map((settlement) => {
      const name = summarizeForPrompt(String(settlement?.name || ""), 40);
      const size = summarizeForPrompt(String(settlement?.size || ""), 20);
      const structure = summarizeForPrompt(String(settlement?.civicStructure || ""), 28);
      return `- Settlement: ${name || "Unnamed"}${size ? ` (${size})` : ""}${structure ? ` | civic: ${structure}` : ""} | influence ${Math.max(0, Number.parseInt(String(settlement?.influence || "0"), 10) || 0)} | dice ${Math.max(0, Number.parseInt(String(settlement?.resourceDice || "0"), 10) || 0)}`;
    })
    .filter(Boolean);
  const regionLines = regions
    .slice(0, compactContext ? 5 : 8)
    .map((region) => {
      const hex = summarizeForPrompt(String(region?.hex || ""), 24);
      const terrain = summarizeForPrompt(String(region?.terrain || ""), 24);
      const workSite = summarizeForPrompt(String(region?.workSite || ""), 30);
      return `- Region: ${hex || "Unknown"} | ${summarizeForPrompt(String(region?.status || ""), 24) || "status unknown"}${terrain ? ` | terrain: ${terrain}` : ""}${workSite ? ` | work site: ${workSite}` : ""}`;
    })
    .filter(Boolean);
  const turnLines = recentTurns
    .slice(0, compactContext ? 3 : 5)
    .map((turn) => {
      const title = summarizeForPrompt(String(turn?.title || ""), 36);
      const summary = summarizeForPrompt(String(turn?.summary || ""), compactContext ? 90 : 140);
      return `- Recent turn: ${title || "Turn"}${summary ? ` | ${summary}` : ""}`;
    })
    .filter(Boolean);
  const projectLines = (Array.isArray(data?.pendingProjects) ? data.pendingProjects : [])
    .slice(0, compactContext ? 4 : 6)
    .map((entry) => `- Pending: ${summarizeForPrompt(String(entry || ""), compactContext ? 100 : 140)}`)
    .filter(Boolean);
  const notes = summarizeForPrompt(String(data?.notes || ""), compactContext ? 180 : 280);
  if (leaderLines.length) lines.push(...leaderLines);
  if (settlementLines.length) lines.push(...settlementLines);
  if (regionLines.length) lines.push(...regionLines);
  if (turnLines.length) lines.push(...turnLines);
  if (projectLines.length) lines.push(...projectLines);
  if (notes) lines.push(`- Kingdom notes: ${notes}`);
  return lines;
}

function getDefaultKingdomRulesProfile() {
  const profiles = Array.isArray(KINGDOM_RULES_DATA?.profiles) ? KINGDOM_RULES_DATA.profiles : [];
  const wanted = String(KINGDOM_RULES_DATA?.latestProfileId || "").trim();
  if (wanted) {
    const match = profiles.find((profile) => String(profile?.id || "").trim() === wanted);
    if (match) return match;
  }
  return profiles[0] || { label: "Kingdom profile", summary: "", turnStructure: [], aiContextSummary: [] };
}

function getModeSpecificPromptLines(mode, input, context = null) {
  const activeTab = String(context?.activeTab || "").toLowerCase();
  if (String(mode || "").toLowerCase() === "npc") {
    const lines = [
      "Format requirements:",
      "Return exactly these top-level labels: Name, Role, Agenda, Disposition, Notes.",
      "Under Notes include 6 to 8 short bullets covering core want, leverage, current pressure or fear, voice and mannerisms, first impression or look, hidden truth or complication, and best way to use them in the next session.",
    ];
    const lowerInput = String(input || "").toLowerCase();
    if (/\b(few|several|multiple|some|2|3|4)\b/.test(lowerInput) && /\bnpcs?\b/.test(lowerInput)) {
      lines.push("If the GM asks for multiple NPCs, return 2 to 4 NPC blocks separated by a line containing only ---.");
    }
    if (String(context?.selectedPdfFile || "").trim() && isPdfGroundedQuestion(lowerInput)) {
      lines.push(
        "Book-grounding requirement: base the NPCs on the selected PDF context. Prefer named or clearly implied figures from that book, and mark any inferred role as inferred instead of inventing unsupported lore."
      );
    }
    return lines;
  }
  if (activeTab === "kingdom") {
    return [
      "Kingdom workflow requirements:",
      "Base advice on the current kingdom sheet and the active V&K rules profile before inventing new plans.",
      "When you suggest changes, make the consequences and tradeoffs explicit so the GM can decide whether to update the records.",
      "Prefer outputs that help the GM record the turn cleanly: action order, what changed, risks, and what should be saved into DM Helper.",
    ];
  }
  return [];
}

function finalizeAiOutput({ rawText, mode, input, tabId }) {
  const raw = String(rawText || "").trim();
  const cleaned = sanitizeAiTextOutput(raw);
  const candidate = cleaned || raw;

  if (candidate && !isLikelyInstructionEcho(candidate) && !isLikelyWeakAiOutput(candidate, mode, input, tabId)) {
    return {
      text: candidate,
      usedFallback: false,
      filtered: cleaned !== raw,
      fallbackReason: "",
    };
  }

  const fallbackReason = !candidate ? "empty" : isLikelyInstructionEcho(candidate) ? "instruction" : "weak";

  return {
    text: generateFallbackAiOutput(mode, input, tabId),
    usedFallback: true,
    filtered: true,
    fallbackReason,
  };
}

function sanitizeAiTextOutput(rawText) {
  const lines = splitAiLines(rawText);
  if (!lines.length) return "";
  const cleaned = lines
    .filter((line) => !isConstraintInstructionLine(line))
    .filter((line) => !isLikelyDuplicateLine(line, lines));
  return cleaned.join("\n").trim();
}

function escapeRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractLabeledBlock(text, label) {
  const source = String(text || "");
  if (!source) return "";
  const regex = new RegExp(`${escapeRegex(label)}\\s*:\\s*([\\s\\S]*?)(?=\\n[A-Za-z][A-Za-z ]{1,28}:|$)`, "i");
  const match = source.match(regex);
  return match ? String(match[1] || "").trim() : "";
}

function collectNpcSupplementalDetailLines(text) {
  const fieldMap = [
    ["Core Want", "Core want"],
    ["Want", "Core want"],
    ["Goal", "Core want"],
    ["Leverage", "Leverage"],
    ["Pressure", "Current pressure"],
    ["Fear", "Current pressure"],
    ["Voice", "Voice and mannerisms"],
    ["Mannerisms", "Voice and mannerisms"],
    ["First Impression", "First impression or look"],
    ["Appearance", "First impression or look"],
    ["Look", "First impression or look"],
    ["Secret", "Hidden truth or complication"],
    ["Hidden Truth", "Hidden truth or complication"],
    ["Complication", "Hidden truth or complication"],
    ["Hook", "Best way to use them in the next session"],
    ["Use At Table", "Best way to use them in the next session"],
    ["Use in Next Session", "Best way to use them in the next session"],
  ];
  const seenTitles = new Set();
  const lines = [];
  for (const [label, title] of fieldMap) {
    const value = extractLabeledBlock(text, label).replace(/\s*\n+\s*/g, " ").trim();
    if (!value || seenTitles.has(title.toLowerCase())) continue;
    seenTitles.add(title.toLowerCase());
    lines.push(`- ${title}: ${value}`);
  }
  return lines;
}

function buildNpcNotesFromAi(text) {
  const baseNotes = extractLabeledBlock(text, "Notes");
  const extraLines = collectNpcSupplementalDetailLines(text);
  if (!baseNotes && !extraLines.length) return "";
  if (!baseNotes) return extraLines.join("\n");

  const existing = baseNotes.toLowerCase();
  const freshLines = extraLines.filter((line) => !existing.includes(line.replace(/^- /, "").split(":")[0].toLowerCase()));
  return [baseNotes, freshLines.join("\n")].filter(Boolean).join("\n");
}

function isWeakNpcOutput(text) {
  const name = extractLabeledBlock(text, "Name");
  const role = extractLabeledBlock(text, "Role");
  const agenda = extractLabeledBlock(text, "Agenda");
  const disposition = extractLabeledBlock(text, "Disposition");
  const notes = buildNpcNotesFromAi(text);
  if (!name || !role || !agenda || !disposition || !notes) return true;

  const noteLines = splitAiLines(notes);
  const bulletCount = noteLines.filter((line) => /^[-*]\s+/.test(line)).length;
  const noteChars = notes.replace(/\s+/g, " ").trim().length;
  if (noteChars < 110) return true;
  if (bulletCount > 0 && bulletCount < 4) return true;
  return false;
}

function countBulletLikeLines(text) {
  return splitAiLines(text).filter((line) => /^[-*]\s+/.test(line) || /^\d+[.)]\s+/.test(line)).length;
}

function isClearlyTruncatedOutput(text) {
  const clean = String(text || "").trim();
  if (!clean) return true;
  const lines = splitAiLines(clean);
  const lastLine = lines[lines.length - 1] || clean;
  if (/[:\-]\s*$/.test(lastLine)) return true;
  if (/[.!?]$/.test(lastLine)) return false;
  const lastWord = String(lastLine || "").toLowerCase().split(/\s+/).filter(Boolean).pop() || "";
  if (
    [
      "a",
      "an",
      "the",
      "and",
      "or",
      "to",
      "of",
      "in",
      "on",
      "with",
      "for",
      "from",
      "is",
      "are",
      "was",
      "were",
      "that",
      "this",
      "these",
      "those",
      "as",
      "at",
      "by",
    ].includes(lastWord)
  ) {
    return true;
  }
  return lines.length <= 2 && clean.length < 160;
}

function isLikelyWeakAiOutput(text, mode, input, tabId) {
  const clean = String(text || "").trim();
  if (!clean) return true;
  if (/^\*{1,2}[^*\n]{1,60}$/.test(clean)) return true;
  if (/^[A-Za-z][A-Za-z ]{1,30}:?$/.test(clean) && clean.length < 40) return true;

  const lines = splitAiLines(clean);
  if (lines.length === 1 && clean.length < 60 && !/[.!?]$/.test(clean)) return true;

  const lowerInput = String(input || "").toLowerCase();
  if (
    tabId === "dashboard" &&
    /\b(opening scene|objective|obstacle|consequence|hook|hooks)\b/.test(lowerInput) &&
    clean.length < 120
  ) {
    return true;
  }

  if (
    tabId === "sessions" &&
    /\b(idea|ideas|hook|hooks|scene|scenes|encounter|encounters|village|town|quest|npc|prep|session|run|opening)\b/.test(
      lowerInput
    ) &&
    clean.length < 160
  ) {
    return true;
  }

  if ((String(mode || "").toLowerCase() === "npc" || tabId === "npcs") && isWeakNpcOutput(clean)) {
    return true;
  }

  if (isClearlyTruncatedOutput(clean)) {
    return true;
  }

  if (isPdfGroundedQuestion(lowerInput)) {
    if (clean.length < 180) return true;
    if (/\b(give me 5 ways|five ways|5 ways to run|ways to run)\b/.test(lowerInput) && countBulletLikeLines(clean) < 3) {
      return true;
    }
  }

  if (String(mode || "").toLowerCase() !== "assistant" && clean.length < 24) return true;
  return false;
}

function splitAiLines(text) {
  return String(text || "")
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function isConstraintInstructionLine(line) {
  const text = String(line || "").trim().toLowerCase();
  if (!text) return false;
  if (/^(output rules|rules|constraints)\s*:/.test(text)) return true;
  if (/^\d+\)\s*/.test(text) && /(no|keep|return|do not|must)/.test(text)) return true;
  if (/(do not|don[’']t)\s+generate\b/.test(text) && /\b(text|content|anything|output)\b/.test(text)) return true;
  if (/\boutside of\b/.test(text) && /(do not|don[’']t|only|must|keep|limit|avoid)/.test(text)) return true;
  if (/^(avoid|do not|don[’']t|must not|never)\b.*\b(answer|response|output)\b/.test(text)) return true;
  if (/^avoid\s+using\b/.test(text) && /\b(in the answer|in your answer|in the response|in output)\b/.test(text)) return true;
  if (/\bin the answer\b/.test(text) && /^(avoid|do not|don[’']t|must|only|keep|limit)/.test(text)) return true;
  if (/^only\s+(return|respond|output)\b/.test(text)) return true;
  if (/^(return|respond|output)\s+only\b/.test(text)) return true;
  if (/^(keep|limit)\s+(the\s+)?(answer|response|output)\b/.test(text)) return true;
  if (/^(keep|limit)\b.*\b(to|under|below|within)\s+\d+\b/.test(text)) return true;
  if (/^(answer|response|output)\s*(length|limit)\b/.test(text)) return true;
  if (/(^|\s)no markdown(,|\.|\s|$)/.test(text)) return true;
  if (/(^|\s)no code(,|\.|\s|$)/.test(text)) return true;
  if (/(^|\s)no emojis?(,|\.|\s|$)/.test(text)) return true;
  if (/no more than\s+\d+/.test(text)) return true;
  if (/keep\s+(the\s+)?output/.test(text)) return true;
  if (/single response/.test(text)) return true;
  if (/output length/.test(text)) return true;
  if (/\b\d+\s*(characters?|words?|tokens?)\b/.test(text) && /(keep|limit|no more than)/.test(text)) return true;
  if (/return plain text only/.test(text)) return true;
  if (/do not repeat instructions/.test(text)) return true;
  return false;
}
function isLikelyDuplicateLine(line, allLines) {
  const candidate = normalizeEchoLine(line);
  if (!candidate) return false;
  let count = 0;
  for (const item of allLines) {
    if (normalizeEchoLine(item) === candidate) count += 1;
    if (count >= 2) return true;
  }
  return false;
}

function isLikelyInstructionEcho(text) {
  const lines = splitAiLines(text);
  if (!lines.length) return true;
  const hits = lines.filter((line) => isConstraintInstructionLine(line)).length;
  if (hits >= 2 || hits / lines.length >= 0.5) return true;
  if (hasRepeatedNearIdenticalLines(lines, 3)) return true;
  const lower = String(text || "").toLowerCase();
  const signalCount = [
    "no markdown",
    "no code",
    "no emojis",
    "no more than",
    "avoid using",
    "in the answer",
    "in your answer",
    "in the response",
    "dont generate any text",
    "don't generate any text",
    "outside of",
    "return only",
    "respond only",
    "output only",
    "keep the answer",
    "keep answer",
    "limit the answer",
    "answer length",
    "keep output",
    "single response",
    "output length",
    "output rules",
    "return plain text only",
  ].filter((token) => lower.includes(token)).length;
  return signalCount >= 2;
}
function hasRepeatedNearIdenticalLines(lines, threshold = 3) {
  const counts = new Map();
  for (const line of lines) {
    const key = normalizeEchoLine(line);
    if (!key) continue;
    const next = (counts.get(key) || 0) + 1;
    counts.set(key, next);
    if (next >= threshold) return true;
  }
  return false;
}

function normalizeEchoLine(line) {
  return String(line || "")
    .toLowerCase()
    .replace(/^\d+\)\s*/, "")
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function generateFallbackAiOutput(mode, input, tabId) {
  const normalizedMode = String(mode || "").toLowerCase();
  const cleanInput = normalizeSentenceText(String(input || "").trim());
  if (normalizedMode === "assistant") {
    return generateAssistantFallbackAnswer(cleanInput);
  }
  if (tabId) {
    const byTab = generateCopilotFallbackByTab(tabId, cleanInput);
    if (byTab) return byTab;
  }
  if (normalizedMode === "prep") {
    return toPrepBullets(cleanInput || "Prepare a concise opening, one obstacle, and one reveal.");
  }
  if (normalizedMode === "recap") {
    return buildRecapFallback(cleanInput || "The party advanced their goals and uncovered a new threat.");
  }
  if (normalizedMode === "npc") {
    return [
      "Name: Frontier Broker",
      "Role: Local fixer with contested loyalties",
      "Agenda: Stay useful to the party while concealing one damaging alliance",
      "Disposition: Helpful, but always measuring the cost",
      "Notes:",
      "- Core want: Survive the current power struggle without losing access or credibility.",
      "- Leverage over the party or locals: Knows the fastest route to the next lead and who is lying about it.",
      "- Current pressure or fear: A stronger faction is about to call in a favor they cannot safely refuse.",
      "- Voice and mannerisms: Speaks quietly, answers sideways first, and watches who reacts to every name.",
      "- First impression or look: Well-kept gear, tired eyes, and the posture of someone who expects betrayal.",
      "- Hidden truth or complication: Already helped the wrong people once and is trying to keep that buried.",
      "- Best way to use them in the next session: Make them the quickest path to progress, then reveal their complication when the party commits.",
    ].join("\n");
  }
  if (normalizedMode === "quest") {
    return `Objective: ${ensureSentence(cleanInput || "Advance the active quest with one clear obstacle and one consequence for delay")}`;
  }
  if (normalizedMode === "location") {
    return `Location Note: ${ensureSentence(cleanInput || "Describe atmosphere, immediate tension, and one clue tied to current events")}`;
  }
  return ensureSentence(cleanInput || "Session notes organized into practical GM prep text.");
}

function generateCopilotFallbackByTab(tabId, input) {
  if (tabId === "dashboard") {
    const lower = String(input || "").toLowerCase();
    if (
      /\b(opening scene|objective|obstacle|consequence)\b/.test(lower) ||
      /\bscene\b/.test(lower)
    ) {
      return [
        "Opening Scene:",
        "Objective: Get the party to commit to helping the border village before supplies run short.",
        "Obstacle: A shaken witness gives conflicting details, while rival locals push different priorities.",
        "Consequence: Delay lets the threat escalate, costing trust and creating a harder first encounter.",
        "Read-Aloud (4-6 sentences):",
        "A cold wind pushes dust across the village square as the party arrives to find shutters barred before dusk. A cart stands overturned near the well, its cargo half-looted and scattered. An exhausted runner grabs the nearest hero and points toward the road, warning that scouts never returned. At the same moment, two villagers begin arguing over whether to fortify the gate or send a rescue team now. Every voice turns toward the party, waiting to see what they do first.",
      ].join("\n");
    }
    if (/\b(hook|hooks)\b/.test(lower)) {
      return [
        "Three Fast Hooks:",
        "- Debt Hook: a caravan master offers payment and future discounts if the party secures the route tonight.",
        "- Duty Hook: a local leader names one missing family member and begs the party to bring them back before dark.",
        "- Rival Hook: another adventuring crew is already taking the job, and failure means losing influence in the region.",
      ].join("\n");
    }
    return [
      "Top Priorities:",
      "- Confirm the opening scene objective and one consequence.",
      "- Pick one active quest to advance this session.",
      "- Prepare one NPC reaction and one location change.",
      "",
      "60-Minute Prep Queue:",
      "- 15m: review last session consequences",
      "- 20m: prep one obstacle with stakes",
      "- 15m: prep clues and reveals",
      "- 10m: prep fallback scene",
    ].join("\n");
  }
  if (tabId === "sessions") {
    return [
      "Summary:",
      ensureSentence(input || "Session notes captured and next objectives clarified."),
      "",
      "Next Prep:",
      "- Open with urgency tied to one active quest.",
      "- Prepare one social beat and one challenge beat.",
      "- End with a clear hook for next session.",
    ].join("\n");
  }
  if (tabId === "capture") {
    return [
      "Summary:",
      ensureSentence(input || "Captured notes should be grouped by scene and consequence."),
      "",
      "Follow-up Tasks:",
      "- Group notes by scene.",
      "- Mark unresolved hooks.",
      "- Push key entries into the latest session log.",
    ].join("\n");
  }
  if (tabId === "kingdom") {
    return [
      "Kingdom Turn Focus:",
      "- Confirm Control DC, unrest, ruin, and consumption before spending actions.",
      "- Assign specialized leader actions first, then use flexible actions to cover gaps.",
      "- Check whether any civic structure, construction project, or event needs to resolve this turn.",
      "",
      "Recommended Action Order:",
      "1. Resolve Upkeep changes, including leadership gaps and automatic kingdom effects.",
      "2. Spend leader and settlement actions on the safest high-value activities for this turn.",
      "3. Record RP, commodities, unrest, ruin, renown, fame, infamy, and pending construction changes.",
      "",
      "Risks To Watch:",
      "- Rising unrest or ruin near the threshold can make the next event spiral fast.",
      "- Consumption and local settlement limits can quietly punish overexpansion.",
      "",
      "What To Record In DM Helper:",
      "- Which leaders acted, what changed, and which projects are still pending.",
      "- Any rulings or reminders you need before the next kingdom turn.",
    ].join("\n");
  }
  if (tabId === "npcs") {
    return [
      "Name: Frontier Contact",
      "Role: Information broker",
      "Agenda: Gain leverage over local factions",
      "Disposition: Cautiously allied",
      "Notes:",
      "- Core want: Stay indispensable to every side without becoming owned by any one faction.",
      "- Leverage over the party or locals: Holds a name, route, or hidden meeting place the party needs.",
      "- Current pressure or fear: One local faction suspects they are selling information twice.",
      "- Voice and mannerisms: Speaks in clipped sentences and never answers the exact question first.",
      "- First impression or look: Polished boots, travel cloak, and the calm posture of someone who expects trouble.",
      "- Hidden truth or complication: Their best source is a person the party would not trust on sight.",
      "- Best way to use them in the next session: Introduce them as the fastest path to a lead, then make the price for help social rather than monetary.",
    ].join("\n");
  }
  if (tabId === "quests") {
    return [
      "Title: Secure the Main Route",
      "Status: open",
      "Objective: Clear threats blocking movement between key settlements.",
      "Giver: Local council envoy",
      "Stakes: Trade and trust collapse if route remains unsafe.",
    ].join("\n");
  }
  if (tabId === "locations") {
    return [
      "Name: Old Waystation",
      "Hex: Frontier Route",
      "What Changed: Signs of hostile activity appeared nearby.",
      "Notes: Use fog, damaged supplies, and witness rumors as scene cues.",
    ].join("\n");
  }
  if (tabId === "pdf") {
    return [
      "Book Context Status: No PDF-grounded answer was generated here. This is a built-in fallback.",
      "Query: adventure summary opening chapter",
      "Backup Queries:",
      "- main threat final chapter",
      "- important NPCs clues chapter one",
      "Why: Summarize the book or search for the specific section you want before asking again.",
    ].join("\n");
  }
  if (tabId === "foundry") {
    return [
      "- Export NPC and quest updates from this session.",
      "- Verify names and titles before import.",
      "- Import JSON and spot-check journal links.",
    ].join("\n");
  }
  return "";
}

function generateAssistantFallbackAnswer(input) {
  const prompt = String(input || "").trim();
  const lower = prompt.toLowerCase();
  if (!prompt) return "Ask one clear GM question and I will generate table-ready options.";
  if (/^(hi|hello|hey|yo)\b/.test(lower) || lower.includes("how are you")) {
    return [
      "Hey. I am ready.",
      "Tell me what you want right now:",
      "- prep plan",
      "- encounter idea",
      "- kingdom turn help",
      "- NPC or quest help",
      "- cleanup of rough notes",
    ].join("\n");
  }
  if (/\b(who|what)\s+are\s+you\b/.test(lower)) {
    return [
      "I am your DM Helper Loremaster running on your local AI setup.",
      "I can help with hooks, session prep, kingdom turns, encounters, NPCs, quests, and note cleanup.",
      "Ask me for one specific thing and I will draft it in table-ready format.",
    ].join("\n");
  }
  if (/\bwhat can you do\b/.test(lower) || /^help\b/.test(lower)) {
    return [
      "I can help right now with:",
      "- Session hook ideas",
      "- Kingdom turn planning and record updates",
      "- Encounter setup (objective, obstacle, consequence)",
      "- NPC or quest drafts",
      "- Cleanup of rough notes into clean prep text",
    ].join("\n");
  }

  if (
    lower.includes("hook") ||
    lower.includes("where to start") ||
    lower.includes("start this") ||
    lower.includes("start the game") ||
    lower.includes("not sure where to start")
  ) {
    return [
      "Try one of these opening hooks:",
      "- Distress Hook: A messenger arrives injured and begs the party to act before nightfall.",
      "- Contract Hook: A local patron offers pay, supplies, and legal authority for one urgent job.",
      "- Personal Hook: The problem directly threatens one PC contact, home, or oath.",
      "",
      "Quick first scene plan:",
      "- Objective: Reach the threatened site and confirm what is happening.",
      "- Obstacle: A rival group or hazard blocks the fastest route.",
      "- Consequence: If delayed, the enemy secures leverage before the party arrives.",
    ].join("\n");
  }

  if (lower.includes("npc")) {
    return [
      "Quick NPC frame:",
      "- Goal: what they want right now.",
      "- Leverage: what they can offer or withhold.",
      "- Pressure: what happens if ignored.",
      "- Voice: one memorable speaking trait.",
    ].join("\n");
  }

  if (/\b(monster|monsters|enemy|enemies|creature|creatures|villain|threat)\b/.test(lower)) {
    return [
      "Monster prep frame:",
      "- Use 1 signature threat tied to the adventure's main problem.",
      "- Use 2 recurring lower-rank enemies so the region feels consistent.",
      "- Add 1 hazard or weird support creature to change the fight rhythm.",
      "",
      "Good categories to choose from:",
      "- Humanoid pressure: bandits, scouts, cultists, rival hunters.",
      "- Supernatural pressure: spirits, undead, cursed beasts, corrupted guardians.",
      "- Environment pressure: traps, haunted ground, fog, unstable bridges, shrine effects.",
      "",
      "Pick monsters that answer:",
      "- What does this region fear most?",
      "- What protects the main villain or secret?",
      "- What shows the consequences before the party reaches the main conflict?",
    ].join("\n");
  }

  if (/\b(run|start with|starting point|first step|adventure|scenario|module|book|players)\b/.test(lower)) {
    return [
      "Start here:",
      "- Read the adventure hook, final threat, and first 2 encounter areas before anything else.",
      "- Pick one clear opening scene that gets the party moving in the first 10 minutes.",
      "- Write 3 names: the first ally, the first problem NPC, and the first place the party reaches.",
      "",
      "First prep pass:",
      "- What do the players need to care about immediately?",
      "- What blocks them in scene one?",
      "- What gets worse if they wait?",
    ].join("\n");
  }

  const hasGmSignal = /\b(hook|quest|npc|session|encounter|player|players|party|location|monster|prep|campaign|story|adventure|scenario|module|book|run)\b/.test(
    lower
  );
  if (!hasGmSignal) {
    return [
      "I can help with your tabletop session prep.",
      "Try asking:",
      "- Give me 3 hooks for level 1 players.",
      "- Build one opening scene with objective, obstacle, consequence.",
      "- Turn these rough notes into session prep bullets.",
    ].join("\n");
  }

  return [
    "Quick answer:",
    ensureSentence(prompt),
    "Turn this into one immediate scene objective, one obstacle, and one consequence.",
  ].join("\n");
}

function toPrepBullets(text) {
  const lines = splitAiLines(text);
  const source = lines.length ? lines : [text];
  return source
    .map((line) => line.replace(/^[-*]\s*/, "").trim())
    .filter(Boolean)
    .map((line) => `- ${ensureSentence(line)}`)
    .join("\n");
}

function buildRecapFallback(text) {
  const clean = normalizeSentenceText(text);
  const parts = clean.split(/[.!?]+/).map((part) => part.trim()).filter(Boolean);
  const first = parts[0] || "the party pushed the story forward";
  const second = parts[1] || "they uncovered new pressure tied to their current objective";
  return `Last session, ${lowercaseFirst(first)}. ${capitalizeFirst(second)}. Now the next chapter begins.`;
}

function normalizeSentenceText(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .trim();
}

function ensureSentence(text) {
  const clean = normalizeSentenceText(text);
  if (!clean) return "";
  return /[.!?]$/.test(clean) ? clean : `${clean}.`;
}

function lowercaseFirst(text) {
  const clean = normalizeSentenceText(text);
  if (!clean) return "";
  return clean.charAt(0).toLowerCase() + clean.slice(1);
}

function capitalizeFirst(text) {
  const clean = normalizeSentenceText(text);
  if (!clean) return "";
  return clean.charAt(0).toUpperCase() + clean.slice(1);
}

function summarizeForPrompt(text, limit) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  if (clean.length <= limit) return clean;
  return `${clean.slice(0, limit)}...`;
}

function getIndexedPdfFileNames(limit = 30) {
  const max = Math.max(1, Math.min(Number.parseInt(String(limit || "30"), 10) || 30, 120));
  const names = pdfIndexCache.files
    .map((file) => String(file?.fileName || "").trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  return names.slice(0, max);
}

function getIndexedPdfFileNamesFromContext(context, limit = 30) {
  const max = Math.max(1, Math.min(Number.parseInt(String(limit || "30"), 10) || 30, 120));
  const names = Array.isArray(context?.pdfIndexedFiles)
    ? context.pdfIndexedFiles.map((name) => String(name || "").trim()).filter(Boolean)
    : [];
  names.sort((a, b) => a.localeCompare(b));
  return names.slice(0, max);
}

function isPdfGroundedQuestion(inputText) {
  const lower = String(inputText || "").toLowerCase().trim();
  if (!lower) return false;
  return /\b(selected pdf|this book|the book|book|pdf|adventure|module|chapter|section|main threat|run chapter|run it|run this)\b/.test(
    lower
  );
}

function isSourceScopeQuestion(text) {
  const lower = String(text || "").toLowerCase().trim();
  if (!lower) return false;
  const directScopePatterns = [
    /\bwhat do you have access to\b/,
    /\bwhat can you access\b/,
    /\bwhat books do you have\b/,
    /\bwhat pdfs?\s+do you have\b/,
    /\bwhich books\b.*\b(can|do)\s+you\b/,
    /\bwhat files are indexed\b/,
    /\bwhich files are indexed\b/,
    /\blist\b.*\b(indexed|loaded|available)\s+(books?|pdfs?|sources?|files?)\b/,
    /\bshow (me )?(the )?(indexed|loaded|available)\s+(books?|pdfs?|sources?|files?)\b/,
    /\bdo you have access to\b.*\b(books?|pdfs?|sources?|rulebooks?)\b/,
    /\bcan you access\b.*\b(books?|pdfs?|sources?|rulebooks?)\b/,
  ];
  return directScopePatterns.some((pattern) => pattern.test(lower));
}

function maybeBuildSourceScopeReply(mode, input, context = null) {
  if (String(mode || "").toLowerCase() !== "assistant") return "";
  if (!isSourceScopeQuestion(input)) return "";

  const liveFiles = getIndexedPdfFileNames(30);
  const contextFiles = getIndexedPdfFileNamesFromContext(context, 30);
  const files = liveFiles.length ? liveFiles : contextFiles;
  const contextCount = Number.parseInt(String(context?.pdfIndexedFileCount || files.length || 0), 10) || 0;
  const liveCount = pdfIndexCache.files.length;
  const total = Math.max(liveCount, contextCount, files.length);

  if (!total) {
    return [
      "I only use files indexed in this app.",
      "No PDFs are indexed yet.",
      "Open PDF Intel and run Index PDFs, then ask again.",
    ].join("\n");
  }

  const lines = [
    "I only have access to PDFs indexed in this app.",
    `Indexed files (${files.length}${total > files.length ? ` of ${total}` : ""}):`,
    ...files.map((name) => `- ${name}`),
  ];
  if (total > files.length) {
    lines.push(`- ...and ${total - files.length} more indexed files.`);
  }
  if (!liveCount && contextCount > 0) {
    lines.push("Note: these come from saved campaign index metadata. Re-index in PDF Intel to load full live text context.");
  }
  lines.push("If a book is not in this indexed list, I do not have access to it.");
  return lines.join("\n");
}

function findIndexedPdfFileByName(fileName) {
  const target = String(fileName || "").trim().toLowerCase();
  if (!target) return null;
  return pdfIndexCache.files.find((file) => String(file?.fileName || "").trim().toLowerCase() === target) || null;
}

function buildSelectedPdfPreview(fileName, maxChars = 1200) {
  const file = findIndexedPdfFileByName(fileName);
  if (!file) return "";
  const pages = getIndexedPages(file).slice(0, 3);
  const combined = pages.map((page) => normalizePdfText(page.text)).filter(Boolean).join(" ");
  return summarizeForPrompt(combined, Math.max(240, Number(maxChars) || 1200));
}

function collectSelectedPdfFallbackContext(fileName, query, limit = 6) {
  const file = findIndexedPdfFileByName(fileName);
  if (!file) return [];
  const pages = getIndexedPages(file);
  if (!pages.length) return [];

  const max = Math.max(1, Math.min(Number(limit) || 6, 12));
  const lowerQuery = String(query || "").toLowerCase();
  const wantsChapterOne = /\b(chapter\s*1|chapter one|opening|first chapter|start of the book)\b/.test(lowerQuery);
  const wantsThreat = /\b(threat|villain|antagonist|enemy|danger|main problem|main conflict)\b/.test(lowerQuery);
  const seen = new Set();
  const ranked = [];

  function addPage(page, score, pattern = null) {
    if (!page?.text || seen.has(page.page)) return;
    seen.add(page.page);
    const textLower = String(page.textLower || page.text.toLowerCase());
    const hit = pattern ? textLower.search(pattern) : -1;
    ranked.push({
      fileName: file.fileName,
      page: page.page,
      score,
      snippet: makeSnippet(page.text, hit),
    });
  }

  const chapterPattern = /\b(chapter\s*1|chapter one|introduction|overview|adventure summary|background|part 1)\b/i;
  const threatPattern = /\b(threat|villain|antagonist|enemy|danger|cult|mastermind|plot|menace)\b/i;

  for (const page of pages.slice(0, 2)) {
    addPage(page, 40);
  }
  if (wantsChapterOne) {
    for (const page of pages) {
      if (chapterPattern.test(page.text)) addPage(page, 90, chapterPattern);
      if (ranked.length >= max) break;
    }
  }
  if (wantsThreat) {
    for (const page of pages) {
      if (threatPattern.test(page.text)) addPage(page, 85, threatPattern);
      if (ranked.length >= max) break;
    }
  }
  for (const page of pages.slice(2)) {
    addPage(page, 20);
    if (ranked.length >= max) break;
  }

  ranked.sort((a, b) => b.score - a.score || a.page - b.page);
  return ranked.slice(0, max);
}

async function collectPdfContextForAi(query, limit = 6, options = {}) {
  const searchParts = buildSearchParts(query);
  const longWords = searchParts.words.filter((word) => word.length >= 4);
  if (!longWords.length && !searchParts.phrase) return [];
  const result = await searchIndexedPdfHybrid(query, Math.max(1, Math.min(Number(limit) || 6, 12)), {
    config: options?.config,
    preferredFileName: options?.preferredFileName,
  });
  return Array.isArray(result?.results) ? result.results.slice(0, Math.max(1, Math.min(Number(limit) || 6, 12))) : [];
}
