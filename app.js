const STORAGE_KEY = "km_gm_studio_v2";
const AI_HISTORY_LIMIT = 180;
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

const tabs = [
  { id: "dashboard", label: "Dashboard", group: "Run Game" },
  { id: "sessions", label: "Session Runner", group: "Run Game" },
  { id: "capture", label: "Live Capture HUD", group: "Run Game" },
  { id: "writing", label: "Writing Helper", group: "Run Game" },
  { id: "kingdom", label: "Kingdom", group: "World" },
  { id: "npcs", label: "NPCs", group: "World" },
  { id: "quests", label: "Quests", group: "World" },
  { id: "locations", label: "Locations", group: "World" },
  { id: "pdf", label: "PDF Intel", group: "Tools" },
  { id: "foundry", label: "Foundry Export", group: "Tools" },
];

const tabGroups = ["Run Game", "World", "Tools"];

const desktopApi = window.kmDesktop || null;
const kingdomRulesData = await loadKingdomRulesData();

let activeTab = "dashboard";
let state = loadState();
const ui = {
  pdfBusy: false,
  pdfMessage: "",
  pdfSearchResults: [],
  pdfSearchQuery: "",
  pdfSummaryBusy: false,
  pdfSummaryFile: "",
  pdfSummaryOutput: "",
  pdfSummaryProgressCurrent: 0,
  pdfSummaryProgressTotal: 0,
  pdfSummaryProgressLabel: "",
  sessionMessage: "",
  kingdomMessage: "",
  customChecklistDraft: "",
  checklistAiBusy: false,
  aiBusy: false,
  aiMessage: "",
  aiLastError: "",
  aiLastErrorAt: "",
  aiTestStatus: "Not run yet.",
  aiTestAt: "",
  aiModels: [],
  copilotBusy: false,
  copilotRequestSeq: 0,
  copilotActiveRequestId: 0,
  copilotMessage: "",
  copilotOpen: false,
  copilotShowOutput: false,
  copilotPendingFallbackMemory: null,
  copilotDraft: {
    input: "",
    output: "",
  },
  worldSelection: {
    npcs: "",
    quests: "",
    locations: "",
  },
  worldMessages: {
    npcs: "",
    quests: "",
    locations: "",
  },
  worldNewFolder: {
    npcs: "",
    quests: "",
    locations: "",
  },
  worldFolderDraft: {
    npcs: "",
    quests: "",
    locations: "",
  },
  wizardOpen: false,
  wizardDraft: {
    sessionId: "",
    highlights: "",
    cliffhanger: "",
    playerIntent: "",
  },
  captureMessage: "",
  captureDraft: {
    sessionId: "",
    kind: "Hook",
    note: "",
  },
  writingDraft: {
    mode: "session",
    input: "",
    output: "",
    autoLink: true,
  },
};

const tabsEl = document.getElementById("tabs");
const appEl = document.getElementById("app");
const seedBtn = document.getElementById("seed-btn");
const exportBtn = document.getElementById("export-btn");
const importBtn = document.getElementById("import-btn");
const importFile = document.getElementById("import-file");

render();
wireGlobalEvents();
void initDesktopDefaults();
// Avoid startup lockups: only auto-run after an explicit tab change.

async function initDesktopDefaults() {
  if (!desktopApi) return;
  try {
    const defaultFolder = await desktopApi.getDefaultPdfFolder();
    if (!state.meta.pdfFolder && defaultFolder) {
      state.meta.pdfFolder = defaultFolder;
    }
    if (desktopApi.getPdfIndexSummary) {
      const summary = await desktopApi.getPdfIndexSummary();
      const count = Number.parseInt(String(summary?.count || "0"), 10) || 0;
      if (count > 0) {
        state.meta.pdfFolder = str(summary?.folderPath) || state.meta.pdfFolder;
        state.meta.pdfIndexedAt = str(summary?.indexedAt) || state.meta.pdfIndexedAt || "";
        state.meta.pdfIndexedCount = count;
        state.meta.pdfIndexedFiles = Array.isArray(summary?.fileNames)
          ? summary.fileNames.map((name) => str(name)).filter(Boolean)
          : state.meta.pdfIndexedFiles || [];
        const files = Array.isArray(summary?.files) ? summary.files : [];
        if (files.length) {
          const summaries = getPdfSummaryMap();
          for (const file of files) {
            const fileName = str(file?.fileName);
            const filePath = str(file?.path);
            const key = filePath || fileName;
            if (!key) continue;
            const text = str(file?.summary);
            if (!text) continue;
            summaries[key] = {
              fileName: fileName || key,
              path: filePath,
              summary: text.slice(0, 24000),
              updatedAt: str(file?.summaryUpdatedAt) || str(summary?.indexedAt) || "",
            };
          }
          state.meta.pdfSummaries = summaries;
        }
      }
    }
    syncPdfSummarySelection();
    saveState();
    render();
    await refreshLocalAiModels(true);
  } catch {
    // Ignore startup desktop API failures.
  }
}

function wireGlobalEvents() {
  if (desktopApi?.onPdfSummarizeProgress) {
    desktopApi.onPdfSummarizeProgress((payload) => {
      applyPdfSummarizeProgress(payload);
    });
  }

  tabsEl.addEventListener("click", (event) => {
    const button = event.target.closest("[data-tab]");
    if (!button) return;
    const nextTab = button.dataset.tab;
    const changed = nextTab !== activeTab;
    activeTab = nextTab;
    render();
    if (changed) {
      void maybeAutoRunCopilotOnTabChange("tab-switch");
    }
  });

  seedBtn.addEventListener("click", () => {
    if (!confirm("Replace current in-app data with starter campaign data?")) return;
    state = createStarterState();
    saveState();
    ui.pdfMessage = "";
    ui.pdfSearchResults = [];
    ui.pdfSearchQuery = "";
    ui.pdfSummaryBusy = false;
    ui.pdfSummaryFile = "";
    ui.pdfSummaryOutput = "";
    resetPdfSummaryProgress();
    ui.sessionMessage = "";
    ui.customChecklistDraft = "";
    ui.checklistAiBusy = false;
    ui.wizardOpen = false;
    ui.wizardDraft = {
      sessionId: "",
      highlights: "",
      cliffhanger: "",
      playerIntent: "",
    };
    ui.captureMessage = "";
    ui.captureDraft = {
      sessionId: "",
      kind: "Hook",
      note: "",
    };
    ui.writingDraft = {
      mode: "session",
      input: "",
      output: "",
      autoLink: true,
    };
    ui.aiMessage = "";
    ui.aiLastError = "";
    ui.aiLastErrorAt = "";
    ui.aiTestStatus = "Not run yet.";
    ui.aiTestAt = "";
    ui.copilotBusy = false;
    ui.copilotRequestSeq = 0;
    ui.copilotActiveRequestId = 0;
    ui.copilotMessage = "";
    ui.copilotOpen = false;
    ui.copilotShowOutput = false;
    ui.copilotPendingFallbackMemory = null;
    ui.copilotDraft = {
      input: "",
      output: "",
    };
    ui.worldSelection = {
      npcs: "",
      quests: "",
      locations: "",
    };
    ui.worldMessages = {
      npcs: "",
      quests: "",
      locations: "",
    };
    ui.worldNewFolder = {
      npcs: "",
      quests: "",
      locations: "",
    };
    ui.worldFolderDraft = {
      npcs: "",
      quests: "",
      locations: "",
    };
    render();
  });

  exportBtn.addEventListener("click", () => {
    downloadJson(state, `dm-helper-campaign-${dateStamp()}.json`);
  });

  importBtn.addEventListener("click", () => importFile.click());
  importFile.addEventListener("change", async () => {
    const file = importFile.files?.[0];
    if (!file) return;
    try {
      const raw = await file.text();
      const parsed = JSON.parse(raw);
      state = normalizeState(parsed);
      saveState();
      ui.pdfSummaryBusy = false;
      ui.pdfSummaryFile = "";
      ui.pdfSummaryOutput = "";
      resetPdfSummaryProgress();
      ui.wizardOpen = false;
      ui.customChecklistDraft = "";
      ui.checklistAiBusy = false;
      ui.captureMessage = "";
      ui.writingDraft = {
        mode: "session",
        input: "",
        output: "",
        autoLink: true,
      };
      ui.aiMessage = "";
      ui.aiLastError = "";
      ui.aiLastErrorAt = "";
      ui.aiTestStatus = "Not run yet.";
      ui.aiTestAt = "";
      ui.copilotBusy = false;
      ui.copilotRequestSeq = 0;
      ui.copilotActiveRequestId = 0;
      ui.copilotMessage = "";
      ui.copilotOpen = false;
      ui.copilotShowOutput = false;
      ui.copilotPendingFallbackMemory = null;
      ui.copilotDraft = {
        input: "",
        output: "",
      };
      ui.worldSelection = {
        npcs: "",
        quests: "",
        locations: "",
      };
      ui.worldMessages = {
        npcs: "",
        quests: "",
        locations: "",
      };
      ui.worldNewFolder = {
        npcs: "",
        quests: "",
        locations: "",
      };
      ui.worldFolderDraft = {
        npcs: "",
        quests: "",
        locations: "",
      };
      render();
    } catch (err) {
      alert(`Import failed: ${String(err)}`);
    } finally {
      importFile.value = "";
    }
  });

  appEl.addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.target;
    if (!(form instanceof HTMLFormElement)) return;
    const type = form.dataset.form;
    if (!type) return;
    void handleFormSubmit(type, form);
  });

  appEl.addEventListener("click", (event) => {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const action = button.dataset.action;
    const collection = button.dataset.collection;
    const id = button.dataset.id;

    if (action === "delete" && collection && id) {
      deleteEntity(collection, id);
      return;
    }

    if (action === "world-select" && collection && id) {
      setWorldSelection(collection, id);
      return;
    }

    if (action === "world-add-folder" && collection) {
      addWorldFolderFromDraft(collection);
      return;
    }

    if (action === "session-wrapup-latest") {
      generateWrapUpForLatestSession();
      return;
    }

    if (action === "session-export-packet-latest") {
      exportSessionPacketForLatest();
      return;
    }

    if (action === "session-export-packet-one" && id) {
      exportSessionPacketForSession(id);
      return;
    }

    if (action === "prep-queue-mode") {
      const mode = Number.parseInt(String(button.dataset.mode || "60"), 10);
      setPrepQueueMode(mode);
      return;
    }

    if (action === "prep-queue-reset") {
      state.meta.prepQueueChecks = {};
      saveState();
      ui.sessionMessage = "Prep queue checks reset.";
      render();
      return;
    }

    if (action === "session-wizard-open-latest") {
      openSessionCloseWizard();
      return;
    }

    if (action === "session-wizard-open-one" && id) {
      openSessionCloseWizard(id);
      return;
    }

    if (action === "session-wizard-cancel") {
      closeSessionCloseWizard();
      return;
    }

    if (action === "session-wrapup-one" && id) {
      generateWrapUpForSession(id);
      return;
    }

    if (action === "session-reset-checklist") {
      state.meta.checklistChecks = {};
      saveState();
      ui.sessionMessage = "Checklist checks reset.";
      render();
      return;
    }

    if (action === "checklist-custom-add") {
      const draftInput = appEl.querySelector("[data-custom-check-draft]");
      const draftValue = draftInput instanceof HTMLInputElement ? draftInput.value : ui.customChecklistDraft;
      const label = normalizeChecklistLabel(draftValue);
      if (!label) {
        ui.sessionMessage = "Type a custom checklist item first.";
        render();
        return;
      }
      const existing = ensureCustomChecklistItems();
      const duplicate = existing.some((item) => item.label.toLowerCase() === label.toLowerCase());
      if (duplicate) {
        ui.sessionMessage = "That custom checklist item already exists.";
        render();
        return;
      }
      state.meta.customChecklistItems = [...existing, { id: `custom-check-${uid()}`, label }];
      ui.customChecklistDraft = "";
      saveState();
      ui.sessionMessage = "Custom checklist item added.";
      render();
      return;
    }

    if (action === "checklist-custom-delete" && id) {
      state.meta.customChecklistItems = ensureCustomChecklistItems().filter((item) => item.id !== id);
      const checks = ensureChecklistChecks();
      delete checks[id];
      state.meta.checklistChecks = checks;
      const overrides = ensureChecklistOverrides();
      delete overrides[id];
      state.meta.checklistOverrides = overrides;
      const archived = ensureChecklistArchived();
      delete archived[id];
      state.meta.checklistArchived = archived;
      saveState();
      ui.sessionMessage = "Custom checklist item removed.";
      render();
      return;
    }

    if (action === "checklist-archive-completed") {
      archiveCompletedChecklistItems();
      return;
    }

    if (action === "checklist-unarchive-all") {
      state.meta.checklistArchived = {};
      saveState();
      ui.sessionMessage = "Archived checklist items restored.";
      render();
      return;
    }

    if (action === "checklist-unarchive-one" && id) {
      const archived = ensureChecklistArchived();
      if (archived[id]) {
        delete archived[id];
        state.meta.checklistArchived = archived;
        saveState();
        ui.sessionMessage = "Checklist item restored.";
      }
      render();
      return;
    }

    if (action === "checklist-remove-old-custom") {
      if (!confirm("Remove all custom checklist items?")) return;
      const customIds = new Set(ensureCustomChecklistItems().map((item) => item.id));
      state.meta.customChecklistItems = [];
      const checks = ensureChecklistChecks();
      for (const id of customIds) delete checks[id];
      state.meta.checklistChecks = checks;
      const overrides = ensureChecklistOverrides();
      for (const id of customIds) delete overrides[id];
      state.meta.checklistOverrides = overrides;
      const archived = ensureChecklistArchived();
      for (const id of customIds) delete archived[id];
      state.meta.checklistArchived = archived;
      saveState();
      ui.sessionMessage = "Old custom checklist items removed.";
      render();
      return;
    }

    if (action === "checklist-ai-generate") {
      void generateChecklistWithAi();
      return;
    }

    if (action === "capture-quick") {
      const kind = str(button.dataset.kind || "Note");
      createCaptureEntry(kind, ui.captureDraft.note, getResolvedCaptureSessionId());
      return;
    }

    if (action === "capture-clear") {
      if (!confirm("Clear all live capture entries?")) return;
      state.liveCapture = [];
      saveState();
      ui.captureMessage = "Live capture log cleared.";
      render();
      return;
    }

    if (action === "capture-append-session") {
      appendCaptureToSession();
      return;
    }

    if (action === "writing-generate") {
      runWritingHelper();
      return;
    }

    if (action === "writing-generate-ai") {
      void runWritingHelperWithLocalAi();
      return;
    }

    if (action === "writing-test-ai") {
      ui.aiMessage = "Testing local AI connection...";
      ui.copilotMessage = "Testing local AI connection...";
      render();
      void testLocalAiConnection();
      return;
    }

    if (action === "writing-copy-output") {
      copyWritingOutput();
      return;
    }

    if (action === "writing-clear") {
      ui.writingDraft = {
        mode: "session",
        input: "",
        output: "",
        autoLink: true,
      };
      render();
      return;
    }

    if (action === "writing-auto-connect-latest") {
      autoConnectWritingOutputToLatestSession();
      return;
    }

    if (action === "writing-apply-latest-session-summary") {
      applyWritingOutputToLatestSession("summary");
      return;
    }

    if (action === "writing-apply-latest-session-prep") {
      applyWritingOutputToLatestSession("nextPrep");
      return;
    }

    if (action === "ai-copilot-generate") {
      void runGlobalAiCopilot();
      return;
    }

    if (action === "ai-copilot-toggle") {
      ui.copilotOpen = !ui.copilotOpen;
      render();
      return;
    }

    if (action === "ai-copilot-output-toggle") {
      ui.copilotShowOutput = !ui.copilotShowOutput;
      render();
      return;
    }

    if (action === "ai-copilot-apply") {
      void applyGlobalAiOutput();
      return;
    }

    if (action === "ai-copilot-copy") {
      void copyGlobalAiOutput();
      return;
    }

    if (action === "ai-copilot-unlock") {
      ui.copilotBusy = false;
      ui.copilotActiveRequestId = 0;
      ui.aiBusy = false;
      ui.copilotMessage = "AI controls unlocked.";
      render();
      return;
    }

    if (action === "ai-copilot-clear") {
      ui.copilotDraft = {
        input: "",
        output: "",
      };
      ui.copilotMessage = "";
      ui.aiLastError = "";
      ui.aiLastErrorAt = "";
      ui.copilotShowOutput = false;
      ui.copilotPendingFallbackMemory = null;
      render();
      return;
    }

    if (action === "ai-history-clear") {
      if (!confirm("Clear saved AI conversation memory?")) return;
      state.meta.aiHistory = [];
      saveState();
      ui.copilotMessage = "Conversation memory cleared.";
      render();
      return;
    }

    if (action === "ai-fallback-save-memory") {
      const pending = ui.copilotPendingFallbackMemory;
      if (!pending || !str(pending.text)) {
        ui.copilotMessage = "No fallback reply available to save.";
        render();
        return;
      }
      addAiHistoryTurn({
        tabId: pending.tabId || activeTab,
        role: "assistant",
        mode: pending.mode || "assistant",
        text: pending.text,
      });
      ui.copilotPendingFallbackMemory = null;
      ui.copilotMessage = "Fallback reply saved to conversation memory.";
      render();
      return;
    }

    if (action === "ai-history-use-input") {
      const turn = getAiHistoryEntryById(target.dataset.historyId);
      if (!turn) {
        ui.copilotMessage = "Saved conversation entry not found.";
        render();
        return;
      }
      ui.copilotDraft.input = str(turn.text);
      ui.copilotMessage = `Loaded ${turn.role === "assistant" ? "AI" : "your"} message into the prompt.`;
      render();
      return;
    }

    if (action === "ai-history-load-output") {
      const turn = getAiHistoryEntryById(target.dataset.historyId);
      if (!turn || turn.role !== "assistant") {
        ui.copilotMessage = "Only saved AI replies can be loaded into output.";
        render();
        return;
      }
      ui.copilotDraft.output = str(turn.text);
      ui.copilotShowOutput = true;
      ui.copilotMessage = "Loaded saved AI reply into the output panel.";
      render();
      return;
    }

    if (action === "ai-history-copy") {
      const turn = getAiHistoryEntryById(target.dataset.historyId);
      if (!turn) {
        ui.copilotMessage = "Saved conversation entry not found.";
        render();
        return;
      }
      void navigator.clipboard.writeText(str(turn.text)).then(
        () => {
          ui.copilotMessage = "Saved conversation entry copied.";
          render();
        },
        () => {
          ui.copilotMessage = "Copy failed. Select the saved entry manually and copy.";
          render();
        }
      );
      return;
    }

    if (action === "ai-copilot-seed") {
      ui.copilotDraft.input = buildGlobalCopilotSeedPrompt(activeTab);
      ui.copilotMessage = "Loaded a tab-specific prompt template.";
      render();
      return;
    }

    if (action === "ai-copilot-test") {
      ui.aiMessage = "Testing local AI connection...";
      ui.copilotMessage = "Testing local AI connection...";
      render();
      void testLocalAiConnection();
      return;
    }

    if (action === "ai-model-refresh") {
      void refreshLocalAiModels();
      return;
    }

    if (action === "ai-profile-fast") {
      applyAiProfile("fast");
      return;
    }

    if (action === "ai-profile-deep") {
      applyAiProfile("deep");
      return;
    }

    if (action === "export-foundry") {
      const kind = button.dataset.kind;
      exportFoundry(kind);
      return;
    }

    if (action === "pdf-choose-folder") {
      void choosePdfFolder();
      return;
    }

    if (action === "pdf-index") {
      void indexPdfLibrary();
      return;
    }

    if (action === "pdf-summarize-selected") {
      void summarizeSelectedPdf(false);
      return;
    }

    if (action === "pdf-summarize-refresh") {
      void summarizeSelectedPdf(true);
      return;
    }

    if (action === "pdf-open-path-page") {
      const encoded = button.dataset.path;
      const pageRaw = Number.parseInt(String(button.dataset.page || "0"), 10);
      if (encoded && desktopApi) {
        const filePath = decodeURIComponent(encoded);
        const page = Number.isNaN(pageRaw) ? 0 : Math.max(0, pageRaw);
        if (page > 0 && typeof desktopApi.openPathAtPage === "function") {
          void desktopApi.openPathAtPage(filePath, page);
        } else {
          void desktopApi.openPath(filePath);
        }
      }
      return;
    }
  });

  appEl.addEventListener("change", (event) => {
    const input = event.target;
    if (
      !(
        input instanceof HTMLInputElement ||
        input instanceof HTMLTextAreaElement ||
        input instanceof HTMLSelectElement
      )
    )
      return;

    if (input.id === "pdf-folder-input") {
      state.meta.pdfFolder = input.value.trim();
      saveState();
      return;
    }

    if (input.dataset.pdfSummaryFile !== undefined) {
      ui.pdfSummaryFile = str(input.value);
      ui.pdfSummaryOutput = "";
      render();
      return;
    }

    const checkId = input.dataset.checkId;
    if (checkId) {
      const checks = ensureChecklistChecks();
      if (input instanceof HTMLInputElement && input.checked) {
        checks[checkId] = true;
      } else {
        delete checks[checkId];
      }
      state.meta.checklistChecks = checks;
      saveState();
      return;
    }

    if (input.dataset.customCheckDraft !== undefined) {
      ui.customChecklistDraft = input.value;
      return;
    }

    const checkEditId = input.dataset.checkEditId;
    if (checkEditId) {
      updateChecklistLabel(checkEditId, input.value);
      return;
    }

    const prepId = input.dataset.prepId;
    if (prepId) {
      const checks = ensurePrepQueueChecks();
      if (input instanceof HTMLInputElement && input.checked) {
        checks[prepId] = true;
      } else {
        delete checks[prepId];
      }
      state.meta.prepQueueChecks = checks;
      saveState();
      return;
    }

    const wizardField = input.dataset.wizardField;
    if (wizardField && ui.wizardDraft && wizardField in ui.wizardDraft) {
      ui.wizardDraft[wizardField] = input.value;
      return;
    }

    const captureField = input.dataset.captureField;
    if (captureField && ui.captureDraft && captureField in ui.captureDraft) {
      ui.captureDraft[captureField] = input.value;
      return;
    }

    const writingField = input.dataset.writingField;
    if (writingField && ui.writingDraft && writingField in ui.writingDraft) {
      if (input instanceof HTMLInputElement && input.type === "checkbox") {
        ui.writingDraft[writingField] = input.checked;
      } else {
        ui.writingDraft[writingField] = input.value;
      }
      return;
    }

    const copilotField = input.dataset.copilotField;
    if (copilotField && ui.copilotDraft && copilotField in ui.copilotDraft) {
      ui.copilotDraft[copilotField] = input.value;
      return;
    }

    const aiField = input.dataset.aiField;
    if (aiField) {
      const config = ensureAiConfig();
      if (input instanceof HTMLInputElement && input.type === "checkbox") {
        config[aiField] = input.checked;
      } else if (input instanceof HTMLInputElement && input.type === "number") {
        if (aiField === "temperature") {
          const parsed = Number.parseFloat(input.value);
          config[aiField] = Number.isFinite(parsed) ? Math.max(0, Math.min(parsed, 2)) : 0.2;
        } else if (aiField === "maxOutputTokens") {
          const parsed = Number.parseInt(input.value, 10);
          config[aiField] = Number.isFinite(parsed) ? Math.max(64, Math.min(parsed, 2048)) : 320;
        } else if (aiField === "timeoutSec") {
          const parsed = Number.parseInt(input.value, 10);
          config[aiField] = Number.isFinite(parsed) ? Math.max(15, Math.min(parsed, 1200)) : 120;
        } else {
          const parsed = Number.parseFloat(input.value);
          config[aiField] = Number.isFinite(parsed) ? parsed : 0;
        }
      } else {
        config[aiField] = input.value;
      }
      config.aiProfile = "custom";
      state.meta.aiConfig = config;
      saveState();
      return;
    }

    const aiModelPick = input.dataset.aiModelPick;
    if (aiModelPick === "model") {
      const value = str(input.value);
      if (!value) return;
      const config = ensureAiConfig();
      config.model = value;
      config.aiProfile = "custom";
      state.meta.aiConfig = config;
      saveState();
      ui.copilotMessage = `Model set to "${value}".`;
      render();
      return;
    }

    const worldNewFolderCollection = input.dataset.worldNewFolder;
    if (worldNewFolderCollection && ui.worldNewFolder && worldNewFolderCollection in ui.worldNewFolder) {
      ui.worldNewFolder[worldNewFolderCollection] = input.value;
      return;
    }

    const worldFolderDraftCollection = input.dataset.worldFolderDraft;
    if (worldFolderDraftCollection && ui.worldFolderDraft && worldFolderDraftCollection in ui.worldFolderDraft) {
      ui.worldFolderDraft[worldFolderDraftCollection] = input.value;
      return;
    }

    const collection = input.dataset.collection;
    const id = input.dataset.id;
    const field = input.dataset.field;
    if (!collection || !id || !field) return;
    if (field === "folder" && isWorldCollection(collection)) {
      const folderName = normalizeWorldFolderName(input.value);
      if (folderName) addWorldFolder(collection, folderName);
    }
    patchEntity(collection, id, { [field]: input.value });
  });
}

function render() {
  tabsEl.innerHTML = renderTabLinks();

  let content = "";
  if (activeTab === "dashboard") content = renderDashboard();
  if (activeTab === "sessions") content = renderSessions();
  if (activeTab === "capture") content = renderCaptureHUD();
  if (activeTab === "writing") content = renderWritingHelper();
  if (activeTab === "kingdom") content = renderKingdom();
  if (activeTab === "npcs") content = renderNpcs();
  if (activeTab === "quests") content = renderQuests();
  if (activeTab === "locations") content = renderLocations();
  if (activeTab === "pdf") content = renderPdfIntel();
  if (activeTab === "foundry") content = renderFoundry();

  appEl.innerHTML = `${content}${renderGlobalAiCopilot()}`;
}

function renderTabLinks() {
  return tabGroups
    .map((group) => {
      const groupTabs = tabs.filter((tab) => tab.group === group);
      if (!groupTabs.length) return "";
      return `
        <section class="tab-group">
          <h3 class="tab-group-title">${escapeHtml(group)}</h3>
          <div class="tab-links">
            ${groupTabs
              .map(
                (tab) => `
                  <button class="tab-link ${tab.id === activeTab ? "active" : ""}" data-tab="${tab.id}">
                    ${escapeHtml(tab.label)}
                  </button>
                `
              )
              .join("")}
          </div>
        </section>
      `;
    })
    .join("");
}

function renderGlobalAiCopilot() {
  const aiConfig = ensureAiConfig();
  const tabLabel = getTabLabel(activeTab);
  const memoryTurns = getRecentAiHistory(activeTab, 16);
  const chatTurns = buildVisibleCopilotChatTurns(memoryTurns, ui.copilotDraft.output, ui.copilotBusy);
  const canSaveFallbackToMemory = !!ui.copilotPendingFallbackMemory?.text;
  const hasOutput = str(ui.copilotDraft.output).length > 0;
  const copilotBusy = ui.copilotBusy;
  const aiBusy = ui.aiBusy;
  const modelOptions = buildAiModelOptions(aiConfig.model, ui.aiModels);
  const rawMessage = replaceAiModelLabelsInText(str(ui.copilotMessage || ui.aiMessage));
  const message = summarizeCopilotStatus(rawMessage);
  const messageTitleAttr = rawMessage && rawMessage !== message ? ` title="${escapeHtml(rawMessage)}"` : "";
  const outputToggleLabel = ui.copilotShowOutput ? "Hide Output" : "Show Output";
  const testLabel = aiBusy ? "Testing..." : "Test AI";
  const aiTestStatus = renderAiTestStatus();

  if (!ui.copilotOpen) {
    return `
      <div class="copilot-layer">
        <button class="copilot-launch" data-action="ai-copilot-toggle">Loremaster (${escapeHtml(tabLabel)})</button>
        ${message ? `<p class="small copilot-mini-status"${messageTitleAttr}>${escapeHtml(message)}</p>` : ""}
      </div>
    `;
  }

  return `
    <div class="copilot-layer">
      <section class="copilot-shell">
        <div class="copilot-head">
          <div>
            <strong>Loremaster</strong>
            <span class="small"> • ${escapeHtml(tabLabel)}</span>
          </div>
          <button class="btn btn-secondary" data-action="ai-copilot-toggle">Hide</button>
        </div>
        ${renderCopilotChatLog(chatTurns, copilotBusy)}
        <label>Prompt
          <textarea class="copilot-prompt" data-copilot-field="input" placeholder="${escapeHtml(getGlobalCopilotPlaceholder(activeTab))}">${escapeHtml(
            ui.copilotDraft.input || ""
          )}</textarea>
        </label>
        <div class="toolbar">
          <button class="btn btn-primary" data-action="ai-copilot-generate">Generate</button>
          <button class="btn btn-secondary" data-action="ai-copilot-seed">Smart Prompt</button>
          <button class="btn btn-secondary" data-action="ai-copilot-test" ${aiBusy ? "disabled" : ""}>${testLabel}</button>
          <button class="btn btn-secondary" data-action="ai-copilot-unlock">Unlock</button>
          <button class="btn btn-secondary" data-action="ai-copilot-copy" ${hasOutput ? "" : "disabled"}>Copy</button>
          <button class="btn btn-primary" data-action="ai-copilot-apply" ${hasOutput ? "" : "disabled"}>${escapeHtml(
            getGlobalCopilotApplyLabel(activeTab)
          )}</button>
          <button class="btn btn-secondary" data-action="ai-copilot-output-toggle" ${hasOutput ? "" : "disabled"}>${outputToggleLabel}</button>
        </div>
        <details class="copilot-settings">
          <summary>AI Settings</summary>
          ${renderAiProfileControls(aiConfig)}
          <div class="row" style="margin-top:8px;">
            <label>Endpoint
              <input data-ai-field="endpoint" value="${escapeHtml(aiConfig.endpoint || "")}" placeholder="http://127.0.0.1:11434" />
            </label>
            <label>Model
              <input data-ai-field="model" value="${escapeHtml(aiConfig.model || "")}" placeholder="llama3.1:8b" />
            </label>
          </div>
          <div class="row">
            <label>Installed Models
              <select data-ai-model-pick="model">
                ${modelOptions}
              </select>
            </label>
            <div style="display:flex;align-items:end;">
              <button class="btn btn-secondary" data-action="ai-model-refresh" ${aiBusy ? "disabled" : ""}>Refresh Models</button>
            </div>
          </div>
          ${renderAiSelectedModelHelp(aiConfig.model)}
          <div class="row">
            <label>Temperature
              <input data-ai-field="temperature" type="number" min="0" max="2" step="0.1" value="${escapeHtml(
                String(aiConfig.temperature ?? 0.2)
              )}" />
            </label>
            <label>Max Output Tokens
              <input data-ai-field="maxOutputTokens" type="number" min="64" max="2048" step="1" value="${escapeHtml(
                String(aiConfig.maxOutputTokens ?? 320)
              )}" />
            </label>
            <label>Timeout (seconds)
              <input data-ai-field="timeoutSec" type="number" min="15" max="1200" step="5" value="${escapeHtml(
                String(aiConfig.timeoutSec ?? 120)
              )}" />
            </label>
          </div>
          <label style="margin-top:8px;">
            <input type="checkbox" data-ai-field="compactContext" ${aiConfig.compactContext ? "checked" : ""} />
            Compact context mode (faster, smaller prompts)
          </label>
          <label style="margin-top:8px;">
            <input type="checkbox" data-ai-field="autoRunTabs" ${aiConfig.autoRunTabs ? "checked" : ""} />
            Auto-run Loremaster on tab switch
          </label>
          <label style="margin-top:8px;">
            <input type="checkbox" data-ai-field="usePdfContext" ${aiConfig.usePdfContext ? "checked" : ""} />
            Use indexed PDF context in AI responses
          </label>
        </details>
        <details class="copilot-settings">
          <summary>Conversation Memory (${memoryTurns.length})</summary>
          <div class="toolbar" style="margin-top:8px;">
            <button class="btn btn-secondary" data-action="ai-history-clear" ${
              memoryTurns.length ? "" : "disabled"
            }>Clear Memory</button>
            <button class="btn btn-secondary" data-action="ai-fallback-save-memory" ${
              canSaveFallbackToMemory ? "" : "disabled"
            }>Save Fallback To Memory</button>
          </div>
          ${
            memoryTurns.length
              ? `<div class="ai-history-list">${memoryTurns.map((turn) => renderAiHistoryTurn(turn)).join("")}</div>`
              : `<p class="small">No saved conversation yet.</p>`
          }
        </details>
        ${message ? `<p class="small copilot-status-line"${messageTitleAttr}>${escapeHtml(message)}</p>` : ""}
        ${aiTestStatus}
        ${renderAiTroubleshootingPanel()}
        ${ui.copilotShowOutput ? renderCopilotOutputPanel(ui.copilotDraft.output || "") : ""}
      </section>
    </div>
  `;
}

function buildVisibleCopilotChatTurns(memoryTurns, currentOutput, isBusy) {
  const visible = Array.isArray(memoryTurns) ? [...memoryTurns] : [];
  const latestAssistant = [...visible].reverse().find((turn) => turn.role === "assistant");
  const outputText = str(currentOutput);
  if (outputText && outputText !== str(latestAssistant?.text)) {
    visible.push({
      id: "copilot-preview",
      role: "assistant",
      text: outputText,
      tabId: activeTab,
      at: new Date().toISOString(),
      ephemeral: true,
    });
  }
  if (isBusy) {
    visible.push({
      id: "copilot-thinking",
      role: "assistant",
      text: "Thinking...",
      tabId: activeTab,
      at: new Date().toISOString(),
      ephemeral: true,
      pending: true,
    });
  }
  return visible.slice(-12);
}

function renderCopilotChatLog(turns, isBusy = false) {
  const chatTurns = Array.isArray(turns) ? turns : [];
  return `
    <section class="copilot-chatlog">
      <div class="copilot-chatlog-head">
        <strong>Chat</strong>
        <span class="small">Follow-up prompts automatically use recent chat memory.</span>
      </div>
      ${
        chatTurns.length
          ? chatTurns.map((turn) => renderCopilotChatTurn(turn)).join("")
          : `<p class="small">Ask a question here, then keep asking follow-ups like a normal chat.</p>`
      }
    </section>
  `;
}

function renderCopilotChatTurn(turn) {
  const role = turn.role === "assistant" ? "assistant" : "user";
  const roleLabel = role === "assistant" ? "Loremaster" : "You";
  const bubbleClass = turn.pending ? `${role} pending` : role;
  const body = role === "assistant"
    ? renderReadableContent(str(turn.text))
    : `<p>${escapeHtml(str(turn.text)).replace(/\n/g, "<br />")}</p>`;
  return `
    <article class="copilot-chatturn ${bubbleClass}">
      <div class="copilot-chatbubble ${bubbleClass}">
        <div class="copilot-chatmeta">
          <strong>${escapeHtml(roleLabel)}</strong>
          <span class="small mono">${escapeHtml(formatAiHistoryTimestamp(turn.at) || "Now")}</span>
        </div>
        <div class="copilot-chatbody">${body}</div>
      </div>
    </article>
  `;
}

function renderAiHistoryTurn(turn) {
  const roleLabel = turn.role === "assistant" ? "AI" : "You";
  const preview = compactLine(str(turn.text).replace(/\s+/g, " "), 150);
  const tabLabel = getTabLabel(str(turn.tabId) || "dashboard");
  const meta = [tabLabel, formatAiHistoryTimestamp(turn.at)].filter(Boolean).join(" • ");
  const canLoadOutput = turn.role === "assistant";
  return `
    <details class="panel" style="margin-top:8px;">
      <summary><strong>${escapeHtml(roleLabel)}</strong>: ${escapeHtml(preview)}${meta ? ` <span class="small">(${escapeHtml(meta)})</span>` : ""}</summary>
      <div class="toolbar" style="margin-top:8px;">
        <button class="btn btn-secondary" data-action="ai-history-use-input" data-history-id="${escapeHtml(turn.id)}">Use as Prompt</button>
        <button class="btn btn-secondary" data-action="ai-history-copy" data-history-id="${escapeHtml(turn.id)}">Copy</button>
        <button class="btn btn-secondary" data-action="ai-history-load-output" data-history-id="${escapeHtml(turn.id)}" ${
          canLoadOutput ? "" : "disabled"
        }>Load Output</button>
      </div>
      <textarea class="copilot-output" readonly style="margin-top:8px;min-height:140px;">${escapeHtml(str(turn.text))}</textarea>
    </details>
  `;
}

function formatAiHistoryTimestamp(value) {
  const time = safeDate(value);
  if (!Number.isFinite(time)) return "";
  return new Date(time).toLocaleString();
}

function renderPageIntro(title, description) {
  return `
    <section class="panel page-intro">
      <h2>${escapeHtml(title)}</h2>
      <p class="small">${escapeHtml(description)}</p>
    </section>
  `;
}

function renderDashboard() {
  const openQuests = state.quests.filter((q) => q.status !== "completed" && q.status !== "failed");
  const recentSessions = [...state.sessions]
    .sort((a, b) => safeDate(b.date) - safeDate(a.date))
    .slice(0, 4);

  return `
    <div class="page-stack">
      ${renderPageIntro("Dashboard", "Quick campaign snapshot and the two things you likely need next: open threads and prep focus.")}
      <section class="grid grid-3">
        <article class="panel stat">
          <span class="small">Sessions Logged</span>
          <span class="stat-value">${state.sessions.length}</span>
        </article>
        <article class="panel stat">
          <span class="small">Active Quests</span>
          <span class="stat-value">${openQuests.length}</span>
        </article>
        <article class="panel stat">
          <span class="small">Tracked NPCs</span>
          <span class="stat-value">${state.npcs.length}</span>
        </article>
      </section>

      <section class="grid grid-2">
        <article class="panel">
          <h2>Open Threads</h2>
          ${
            openQuests.length
              ? `<ul class="list">${openQuests
                  .map((q) => `<li><strong>${escapeHtml(q.title)}</strong> <span class="small">(${escapeHtml(q.status)})</span></li>`)
                  .join("")}</ul>`
              : `<p class="empty">No open quests tracked.</p>`
          }
        </article>

        <article class="panel">
          <h2>Latest Session Prep Notes</h2>
          ${
            recentSessions.length
              ? `<ul class="list">${recentSessions
                  .map(
                    (s) =>
                      `<li><strong>${escapeHtml(s.title)}</strong>: ${escapeHtml(
                        (s.nextPrep || "").slice(0, 110) || "No prep note yet"
                      )}</li>`
                  )
                  .join("")}</ul>`
              : `<p class="empty">No sessions yet.</p>`
          }
        </article>
      </section>
    </div>
  `;
}

function renderSessions() {
  const sessions = [...state.sessions].sort((a, b) => safeDate(b.date) - safeDate(a.date));
  const checklistItems = generateSmartChecklist();
  const customChecklistIds = new Set(ensureCustomChecklistItems().map((item) => item.id));
  const checklistChecks = ensureChecklistChecks();
  const checklistArchived = ensureChecklistArchived();
  const allChecklistItems = generateSmartChecklist({ includeArchived: true });
  const checkedVisibleItems = checklistItems.filter((item) => checklistChecks[item.id]);
  const archivedItems = getArchivedChecklistItems(allChecklistItems, checklistArchived);
  const archivedCount = Object.keys(checklistArchived).length;
  const completedChecklist = checklistItems.filter((item) => checklistChecks[item.id]).length;
  const checklistAiBusyAttr = ui.checklistAiBusy ? "disabled" : "";
  const prepMode = getPrepQueueMode();
  const prepQueue = generatePrepQueue(prepMode);
  const prepChecks = ensurePrepQueueChecks();
  const prepDone = prepQueue.filter((task) => prepChecks[task.id]).length;

  return `
    <div class="page-stack">
      ${renderPageIntro(
        "Session Runner",
        "Use this in order: Step 1 prep before game, Step 2 log what happened, Step 3 close and generate your next prep packet."
      )}

      <section class="panel flow-panel">
        <h2>Run Order</h2>
        <ol class="flow-list">
          <li><strong>Step 1 Prep:</strong> complete checklist and time-boxed prep queue.</li>
          <li><strong>Step 2 Run + Log:</strong> create/update session log during or right after play.</li>
          <li><strong>Step 3 Close:</strong> run wrap-up wizard, then export next-session packet.</li>
        </ol>
        ${ui.sessionMessage ? `<p class="small">${escapeHtml(ui.sessionMessage)}</p>` : ""}
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">1</span>
          <h2>Prep Before Session</h2>
        </div>
        <section class="step-grid">
          <article class="step-sub">
            <h3>Smart Checklist</h3>
            <p class="small mono">Progress: ${completedChecklist}/${checklistItems.length}</p>
            <ul class="checklist-list">
              ${
                checklistItems.length
                  ? checklistItems
                      .map(
                        (item) => `
                        <li>
                          <div class="check-row check-edit-row">
                            <input type="checkbox" data-check-id="${item.id}" ${checklistChecks[item.id] ? "checked" : ""} />
                            <input class="check-label-input" data-check-edit-id="${item.id}" value="${escapeHtml(item.label)}" />
                            ${
                              customChecklistIds.has(item.id)
                                ? `<button class="btn btn-danger check-row-delete" data-action="checklist-custom-delete" data-id="${item.id}">X</button>`
                                : ""
                            }
                          </div>
                        </li>
                      `
                      )
                      .join("")
                  : `<li class="empty">No checklist items yet.</li>`
              }
            </ul>
            <div class="toolbar">
              <button class="btn btn-secondary" data-action="session-reset-checklist">Reset Checks</button>
              <button class="btn btn-secondary" data-action="checklist-archive-completed">Archive Completed</button>
              <button class="btn btn-secondary" data-action="checklist-unarchive-all" ${archivedCount ? "" : "disabled"}>
                Unarchive (${archivedCount})
              </button>
              <button class="btn btn-secondary" data-action="checklist-remove-old-custom" ${
                customChecklistIds.size ? "" : "disabled"
              }>Remove Old Custom</button>
            </div>
            <div class="toolbar">
              <input data-custom-check-draft value="${escapeHtml(ui.customChecklistDraft || "")}" placeholder="Add custom checklist item..." />
              <button class="btn btn-secondary" data-action="checklist-custom-add">Add Custom Item</button>
              <button class="btn btn-primary" data-action="checklist-ai-generate" ${checklistAiBusyAttr}>
                ${ui.checklistAiBusy ? "AI Generating..." : "AI Create Checklist"}
              </button>
            </div>
            ${archivedCount ? `<p class="small">Archived items are hidden until you click Unarchive.</p>` : ""}
            <details class="world-create" style="margin-top:10px;">
              <summary>Completed / Archived Checklist</summary>
              <div style="margin-top:8px;">
                <p class="small"><strong>Checked (current):</strong> ${checkedVisibleItems.length}</p>
                ${
                  checkedVisibleItems.length
                    ? `<ul class="list">${checkedVisibleItems.map((item) => `<li>${escapeHtml(item.label)}</li>`).join("")}</ul>`
                    : `<p class="empty">No currently checked items.</p>`
                }
                <p class="small" style="margin-top:8px;"><strong>Archived (hidden):</strong> ${archivedItems.length}</p>
                ${
                  archivedItems.length
                    ? `<ul class="list">${archivedItems
                        .map(
                          (item) =>
                            `<li>${escapeHtml(item.label)} <button class="btn btn-secondary" data-action="checklist-unarchive-one" data-id="${item.id}">Restore</button></li>`
                        )
                        .join("")}</ul>`
                    : `<p class="empty">No archived checklist items.</p>`
                }
              </div>
            </details>
          </article>
          <article class="step-sub">
            <h3>Prep Queue (${prepMode}m)</h3>
            <p class="small mono">Progress: ${prepDone}/${prepQueue.length}</p>
            <div class="toolbar">
              <button class="btn ${prepMode === 30 ? "btn-primary" : "btn-secondary"}" data-action="prep-queue-mode" data-mode="30">30m</button>
              <button class="btn ${prepMode === 60 ? "btn-primary" : "btn-secondary"}" data-action="prep-queue-mode" data-mode="60">60m</button>
              <button class="btn ${prepMode === 90 ? "btn-primary" : "btn-secondary"}" data-action="prep-queue-mode" data-mode="90">90m</button>
              <button class="btn btn-secondary" data-action="prep-queue-reset">Reset Queue Checks</button>
            </div>
            <ul class="checklist-list" style="margin-top:10px;">
              ${
                prepQueue.length
                  ? prepQueue
                      .map(
                        (task) => `
                        <li>
                          <div class="check-row">
                            <input type="checkbox" data-prep-id="${task.id}" ${prepChecks[task.id] ? "checked" : ""} />
                            <span>${escapeHtml(task.label)} <span class="small mono">(${task.minutes}m)</span></span>
                          </div>
                        </li>
                      `
                      )
                      .join("")
                  : `<li class="empty">No prep queue items yet.</li>`
              }
            </ul>
          </article>
        </section>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">2</span>
          <h2>Run + Log Session</h2>
        </div>
        <section class="step-grid sessions-step-grid">
          <article class="step-sub">
            <h3>Create Session Log</h3>
            <form data-form="sessions">
              <div class="row">
                <label>Session Title
                  <input name="title" required placeholder="Session 07 - Echoes at Blackbridge" />
                </label>
                <label>Date
                  <input name="date" type="date" required />
                </label>
              </div>
              <div class="row">
                <label>Campaign Arc
                  <input name="arc" placeholder="Frontier Arc / Court Arc / Campaign Turn" />
                </label>
                <label>Campaign Turn
                  <input name="kingdomTurn" placeholder="Turn 3 (optional)" />
                </label>
              </div>
              <label>What Happened
                <textarea name="summary" placeholder="Fast bullets from table play: scenes, consequences, hooks..."></textarea>
              </label>
              <label>Next Session Prep
                <textarea name="nextPrep" placeholder="Cold open, likely encounters, NPCs to prep, PDFs to recheck..."></textarea>
              </label>
              <div class="toolbar">
                <button class="btn btn-primary" type="submit">Add Session</button>
              </div>
            </form>
          </article>
          <article class="step-sub">
            <h3>Session Logs</h3>
            <div class="card-list">
              ${
                sessions.length
                  ? sessions.map((s) => sessionEntry(s)).join("")
                  : `<p class="empty">No sessions tracked yet.</p>`
              }
            </div>
          </article>
        </section>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">3</span>
          <h2>Close Session + Export</h2>
        </div>
        <p class="small">When play ends, generate wrap-up bullets, scene openers, and a prep packet for next game.</p>
        <div class="toolbar">
          <button class="btn btn-primary" data-action="session-wrapup-latest">Smart Wrap-Up Latest Session</button>
          <button class="btn btn-primary" data-action="session-wizard-open-latest">Open Session Close Wizard</button>
          <button class="btn btn-secondary" data-action="session-export-packet-latest">Export Next Session Packet</button>
        </div>
        ${ui.wizardOpen ? renderSessionCloseWizard(sessions) : ""}
      </section>
    </div>
  `;
}

function renderSessionCloseWizard(sessions) {
  const fallbackId = ui.wizardDraft.sessionId || sessions[0]?.id || "";
  return `
    <section class="panel session-wizard-panel">
      <h2>Session Close Wizard (3-Step)</h2>
      <p class="small">Answer the three prompts, then the app auto-generates wrap-up bullets and 3 scene openers.</p>
      <form data-form="session-close-wizard">
        <label>Target Session
          <select name="sessionId" data-wizard-field="sessionId">
            ${sessions
              .map(
                (session) =>
                  `<option value="${session.id}" ${session.id === fallbackId ? "selected" : ""}>${escapeHtml(
                    session.title
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label>Step 1: Biggest moments tonight
          <textarea name="highlights" data-wizard-field="highlights" placeholder="What happened that must matter next session?">${escapeHtml(
            ui.wizardDraft.highlights || ""
          )}</textarea>
        </label>
        <label>Step 2: Cliffhanger or unresolved pressure
          <textarea name="cliffhanger" data-wizard-field="cliffhanger" placeholder="What tension is still hanging?">${escapeHtml(
            ui.wizardDraft.cliffhanger || ""
          )}</textarea>
        </label>
        <label>Step 3: What players want to do next
          <textarea name="playerIntent" data-wizard-field="playerIntent" placeholder="What did players say they want next?">${escapeHtml(
            ui.wizardDraft.playerIntent || ""
          )}</textarea>
        </label>
        <div class="toolbar">
          <button class="btn btn-primary" type="submit">Run Wizard + Generate Prep</button>
          <button class="btn btn-secondary" type="button" data-action="session-wizard-cancel">Cancel</button>
        </div>
      </form>
    </section>
  `;
}

function renderCaptureHUD() {
  const sessions = [...state.sessions].sort((a, b) => safeDate(b.date) - safeDate(a.date));
  const activeSessionId = ui.captureDraft.sessionId || sessions[0]?.id || "";
  const entries = [...(state.liveCapture || [])].sort((a, b) => safeDate(b.timestamp) - safeDate(a.timestamp));

  return `
    <div class="page-stack">
      ${renderPageIntro("Live Capture HUD", "Fast in-session notes with clear tags, then push them into your session log when ready.")}
      <section class="grid grid-2">
    <section class="panel">
      <h2>Live Capture HUD</h2>
      <p class="small">Use this while running the table. Add short timestamped notes fast.</p>
      <div class="row">
        <label>Attach Notes To Session
          <select data-capture-field="sessionId">
            <option value="">No session link</option>
            ${sessions
              .map(
                (session) =>
                  `<option value="${session.id}" ${session.id === activeSessionId ? "selected" : ""}>${escapeHtml(
                    session.title
                  )}</option>`
              )
              .join("")}
          </select>
        </label>
        <label>Default Capture Type
          <select data-capture-field="kind">
            ${["Hook", "NPC", "Rule", "Loot", "Retcon", "Scene", "Combat", "Note"]
              .map(
                (kind) =>
                  `<option value="${kind}" ${ui.captureDraft.kind === kind ? "selected" : ""}>${kind}</option>`
              )
              .join("")}
          </select>
        </label>
      </div>

      <label>Quick Note
        <textarea data-capture-field="note" placeholder="Short, table-speed note...">${escapeHtml(ui.captureDraft.note || "")}</textarea>
      </label>

      <div class="toolbar">
        <button class="btn btn-primary" data-action="capture-quick" data-kind="${escapeHtml(ui.captureDraft.kind || "Note")}">Capture (${escapeHtml(
          ui.captureDraft.kind || "Note"
        )})</button>
        <button class="btn btn-secondary" data-action="capture-quick" data-kind="NPC">NPC</button>
        <button class="btn btn-secondary" data-action="capture-quick" data-kind="Hook">Hook</button>
        <button class="btn btn-secondary" data-action="capture-quick" data-kind="Rule">Rule</button>
        <button class="btn btn-secondary" data-action="capture-quick" data-kind="Loot">Loot</button>
        <button class="btn btn-secondary" data-action="capture-quick" data-kind="Retcon">Retcon</button>
        <button class="btn btn-secondary" data-action="capture-append-session">Append to Session</button>
        <button class="btn btn-danger" data-action="capture-clear">Clear Log</button>
      </div>
      ${ui.captureMessage ? `<p class="small">${escapeHtml(ui.captureMessage)}</p>` : ""}
    </section>

    <section class="panel">
      <h2>Captured Entries</h2>
      <div class="card-list">
        ${
          entries.length
            ? entries.map((entry) => renderCaptureEntry(entry, sessions)).join("")
            : `<p class="empty">No live capture entries yet.</p>`
        }
      </div>
    </section>
      </section>
    </div>
  `;
}

function renderCaptureEntry(entry, sessions) {
  const linked = sessions.find((session) => session.id === entry.sessionId);
  const stamp = entry.timestamp ? new Date(entry.timestamp).toLocaleString() : "Unknown time";
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(entry.kind || "Note")}</span>
        <span class="entry-meta">${escapeHtml(stamp)}</span>
      </div>
      <p>${escapeHtml(entry.note || "")}</p>
      <p class="small">${linked ? `Linked Session: ${escapeHtml(linked.title)}` : "No session link"}</p>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="liveCapture" data-id="${entry.id}">Delete</button>
      </div>
    </article>
  `;
}

function renderWritingHelper() {
  const hasOutput = str(ui.writingDraft.output).length > 0;
  const aiConfig = ensureAiConfig();
  const aiBusy = ui.aiBusy ? "disabled" : "";
  const aiStatus = replaceAiModelLabelsInText(ui.aiMessage || "");
  const testLabel = ui.aiBusy ? "Testing..." : "Test Local AI";
  return `
    <div class="page-stack">
      ${renderPageIntro("Writing Helper", "Turn rough notes into clean GM text, then apply directly to your latest session fields.")}
      <section class="grid grid-2">
        <section class="panel">
          <h2>Draft + Actions</h2>
          <div class="row">
            <label>Mode
              <select data-writing-field="mode">
                ${[
                  ["assistant", "GM Assistant (Q&A)"],
                  ["session", "Session Summary"],
                  ["recap", "Read-Aloud Recap"],
                  ["npc", "NPC Blurb"],
                  ["quest", "Quest Objective"],
                  ["location", "Location Description"],
                  ["prep", "Next Session Prep Bullets"],
                ]
                  .map(
                    ([value, label]) =>
                      `<option value="${value}" ${ui.writingDraft.mode === value ? "selected" : ""}>${label}</option>`
                  )
                  .join("")}
              </select>
            </label>
          </div>

          <label>Draft Input
            <textarea data-writing-field="input" placeholder="Ask naturally or type rough notes here...">${escapeHtml(
              ui.writingDraft.input || ""
            )}</textarea>
          </label>
          <div class="toolbar">
            <button class="btn btn-primary" data-action="writing-generate">Generate Clean Text</button>
            <button class="btn btn-primary" data-action="writing-generate-ai" ${aiBusy}>Generate With Local AI</button>
            <button class="btn btn-secondary" data-action="writing-copy-output" ${hasOutput ? "" : "disabled"}>Copy Output</button>
            <button class="btn btn-secondary" data-action="writing-apply-latest-session-summary" ${hasOutput ? "" : "disabled"}>Use as Latest Summary</button>
            <button class="btn btn-secondary" data-action="writing-apply-latest-session-prep" ${hasOutput ? "" : "disabled"}>Use as Latest Prep</button>
            <button class="btn btn-secondary" data-action="writing-auto-connect-latest" ${hasOutput ? "" : "disabled"}>Auto-Connect to Latest Session</button>
            <button class="btn btn-danger" data-action="writing-clear">Clear</button>
          </div>
          <label style="margin-top:8px;">
            <input type="checkbox" data-writing-field="autoLink" ${ui.writingDraft.autoLink ? "checked" : ""} />
            Auto-connect entities after AI generate
          </label>
          ${ui.sessionMessage ? `<p class="small">${escapeHtml(ui.sessionMessage)}</p>` : ""}

          <details class="copilot-settings" style="margin-top:10px;">
            <summary>Local AI Setup</summary>
            ${renderAiProfileControls(aiConfig)}
            <div class="row" style="margin-top:8px;">
              <label>Endpoint
                <input data-ai-field="endpoint" value="${escapeHtml(aiConfig.endpoint || "")}" placeholder="http://127.0.0.1:11434" />
              </label>
              <label>Model
                <input data-ai-field="model" value="${escapeHtml(aiConfig.model || "")}" placeholder="llama3.1:8b" />
              </label>
            </div>
            ${renderAiSelectedModelHelp(aiConfig.model)}
            <div class="row">
              <label>Temperature
                <input data-ai-field="temperature" type="number" min="0" max="2" step="0.1" value="${escapeHtml(
                  String(aiConfig.temperature ?? 0.2)
                )}" />
              </label>
              <label>Max Output Tokens
                <input data-ai-field="maxOutputTokens" type="number" min="64" max="2048" step="1" value="${escapeHtml(
                  String(aiConfig.maxOutputTokens ?? 320)
                )}" />
              </label>
              <label>Timeout (seconds)
                <input data-ai-field="timeoutSec" type="number" min="15" max="1200" step="5" value="${escapeHtml(
                  String(aiConfig.timeoutSec ?? 120)
                )}" />
              </label>
            </div>
            <label style="margin-top:8px;">
              <input type="checkbox" data-ai-field="compactContext" ${aiConfig.compactContext ? "checked" : ""} />
              Compact context mode (faster, smaller prompts)
            </label>
            <label style="margin-top:8px;">
              <input type="checkbox" data-ai-field="usePdfContext" ${aiConfig.usePdfContext ? "checked" : ""} />
              Use indexed PDF context in AI responses
            </label>
            <div class="toolbar">
              <button class="btn btn-secondary" data-action="writing-test-ai">${testLabel}</button>
            </div>
            ${aiStatus ? `<p class="small">${escapeHtml(aiStatus)}</p>` : ""}
            ${renderAiTestStatus()}
            ${renderAiTroubleshootingPanel()}
          </details>
        </section>

        <section class="panel">
          <h2>Output</h2>
          <textarea readonly>${escapeHtml(ui.writingDraft.output || "")}</textarea>
          <p class="small">Tip: right-click in any text field to see spellcheck suggestions.</p>
        </section>
      </section>
    </div>
  `;
}

function renderNpcs() {
  const selected = getSelectedWorldEntry("npcs", state.npcs);
  const folderOptionsNew = renderWorldFolderOptions("npcs", ui.worldNewFolder.npcs || "", true);
  return `
    <div class="page-stack">
      ${renderPageIntro("NPCs", "Track voices, motives, and notes so recurring characters stay consistent at the table.")}
      <section class="world-layout">
        <section class="panel world-sidebar">
          <h2>NPC Links</h2>
          <div class="toolbar">
            <input data-world-folder-draft="npcs" value="${escapeHtml(ui.worldFolderDraft.npcs || "")}" placeholder="New folder (e.g., Rivergate)" />
            <button class="btn btn-secondary" data-action="world-add-folder" data-collection="npcs">New Folder</button>
          </div>
          ${ui.worldMessages.npcs ? `<p class="small">${escapeHtml(ui.worldMessages.npcs)}</p>` : ""}
          ${renderWorldLinkList("npcs", state.npcs, (npc) => ({
            title: npc.name || "Unnamed NPC",
            meta: `${getWorldFolderLabel(npc.folder)}${npc.role ? ` • ${npc.role}` : npc.disposition ? ` • ${npc.disposition}` : ""}`,
          }))}
          <details class="world-create" ${state.npcs.length ? "" : "open"}>
            <summary>New NPC</summary>
            <form data-form="npcs">
              <label>Folder
                <select name="folder" data-world-new-folder="npcs">
                  ${folderOptionsNew}
                </select>
              </label>
              <label>Name
                <input name="name" required placeholder="Lady Ardyn Vale" />
              </label>
              <label>Role
                <input name="role" placeholder="Swordlord patron" />
              </label>
              <label>Agenda
                <input name="agenda" placeholder="What they want right now" />
              </label>
              <label>Disposition
                <input name="disposition" placeholder="Allied / Neutral / Hostile" />
              </label>
              <label>Notes
                <textarea name="notes" placeholder="Voice cues, secrets, leverage, links to quests..."></textarea>
              </label>
              <button class="btn btn-primary" type="submit">Add NPC</button>
            </form>
          </details>
        </section>
        <section class="panel world-detail">
          <h2>NPC Details</h2>
          ${selected ? renderNpcDetails(selected) : `<p class="empty">No NPC selected.</p>`}
        </section>
      </section>
    </div>
  `;
}

function renderQuests() {
  const selected = getSelectedWorldEntry("quests", state.quests);
  const folderOptionsNew = renderWorldFolderOptions("quests", ui.worldNewFolder.quests || "", true);
  return `
    <div class="page-stack">
      ${renderPageIntro("Quests", "Keep objectives and stakes explicit so your players always have clear, actionable direction.")}
      <section class="world-layout">
        <section class="panel world-sidebar">
          <h2>Quest Links</h2>
          <div class="toolbar">
            <input data-world-folder-draft="quests" value="${escapeHtml(ui.worldFolderDraft.quests || "")}" placeholder="New folder (e.g., Main Campaign)" />
            <button class="btn btn-secondary" data-action="world-add-folder" data-collection="quests">New Folder</button>
          </div>
          ${ui.worldMessages.quests ? `<p class="small">${escapeHtml(ui.worldMessages.quests)}</p>` : ""}
          ${renderWorldLinkList("quests", state.quests, (quest) => ({
            title: quest.title || "Untitled Quest",
            meta: `${getWorldFolderLabel(quest.folder)} • ${quest.status || "open"}${quest.giver ? ` • ${quest.giver}` : ""}`,
          }))}
          <details class="world-create" ${state.quests.length ? "" : "open"}>
            <summary>New Quest</summary>
            <form data-form="quests">
              <label>Folder
                <select name="folder" data-world-new-folder="quests">
                  ${folderOptionsNew}
                </select>
              </label>
              <label>Title
                <input name="title" required placeholder="Bandit Pressure at the Trading Post" />
              </label>
              <label>Status
                <select name="status">
                  <option value="open">open</option>
                  <option value="in-progress">in-progress</option>
                  <option value="blocked">blocked</option>
                  <option value="completed">completed</option>
                  <option value="failed">failed</option>
                </select>
              </label>
              <label>Objective
                <textarea name="objective" placeholder="What must happen for this quest to move forward?"></textarea>
              </label>
              <label>Quest Giver
                <input name="giver" placeholder="Quartermaster Bren" />
              </label>
              <label>Stakes
                <input name="stakes" placeholder="If ignored, supply lines collapse..." />
              </label>
              <button class="btn btn-primary" type="submit">Add Quest</button>
            </form>
          </details>
        </section>
        <section class="panel world-detail">
          <h2>Quest Details</h2>
          ${selected ? renderQuestDetails(selected) : `<p class="empty">No quest selected.</p>`}
        </section>
      </section>
    </div>
  `;
}

function renderLocations() {
  const selected = getSelectedWorldEntry("locations", state.locations);
  const folderOptionsNew = renderWorldFolderOptions("locations", ui.worldNewFolder.locations || "", true);
  return `
    <div class="page-stack">
      ${renderPageIntro("Locations", "Record what changed in each place so your world state stays coherent between sessions.")}
      <section class="world-layout">
        <section class="panel world-sidebar">
          <h2>Location Links</h2>
          <div class="toolbar">
            <input data-world-folder-draft="locations" value="${escapeHtml(ui.worldFolderDraft.locations || "")}" placeholder="New folder (e.g., North March)" />
            <button class="btn btn-secondary" data-action="world-add-folder" data-collection="locations">New Folder</button>
          </div>
          ${ui.worldMessages.locations ? `<p class="small">${escapeHtml(ui.worldMessages.locations)}</p>` : ""}
          ${renderWorldLinkList("locations", state.locations, (location) => ({
            title: location.name || "Unnamed Location",
            meta: `${getWorldFolderLabel(location.folder)}${location.hex ? ` • ${location.hex}` : ""}`,
          }))}
          <details class="world-create" ${state.locations.length ? "" : "open"}>
            <summary>New Location / Hex</summary>
            <form data-form="locations">
              <label>Folder
                <select name="folder" data-world-new-folder="locations">
                  ${folderOptionsNew}
                </select>
              </label>
              <label>Name
                <input name="name" required placeholder="Blackbridge Waystation" />
              </label>
              <label>Hex / Region
                <input name="hex" placeholder="A2 / North March" />
              </label>
              <label>What Changed
                <textarea name="whatChanged" placeholder="Ownership shifts, threats cleared, new rumors, construction..."></textarea>
              </label>
              <label>Notes
                <textarea name="notes" placeholder="Scene hooks, map notes, hidden details..."></textarea>
              </label>
              <button class="btn btn-primary" type="submit">Add Location</button>
            </form>
          </details>
        </section>
        <section class="panel world-detail">
          <h2>Location Details</h2>
          ${selected ? renderLocationDetails(selected) : `<p class="empty">No location selected.</p>`}
        </section>
      </section>
    </div>
  `;
}

function renderWorldLinkList(collection, items, formatter) {
  if (!items.length) return `<p class="empty">No entries yet.</p>`;
  const selected = getSelectedWorldEntry(collection, items);
  const selectedId = selected?.id || "";
  const groups = buildWorldGroups(items);
  return `
    <div class="world-links">
      ${groups
        .map(
          (group) => `
          <section class="world-folder-group">
            <h3>${escapeHtml(group.label)} <span class="small">(${group.items.length})</span></h3>
            ${group.items
              .map((item) => {
                const view = formatter(item);
                return `
                  <button class="world-link ${item.id === selectedId ? "active" : ""}" data-action="world-select" data-collection="${collection}" data-id="${
                    item.id
                  }">
                    <span class="world-link-title">${escapeHtml(view.title)}</span>
                    <span class="world-link-meta">${escapeHtml(view.meta || "")}</span>
                  </button>
                `;
              })
              .join("")}
          </section>
        `
        )
        .join("")}
    </div>
  `;
}

function buildWorldGroups(items) {
  const map = new Map();
  for (const item of items) {
    const label = getWorldFolderLabel(item?.folder);
    if (!map.has(label)) map.set(label, []);
    map.get(label).push(item);
  }
  const labels = [...map.keys()].sort((a, b) => {
    if (a === "Unsorted") return -1;
    if (b === "Unsorted") return 1;
    return a.localeCompare(b);
  });
  return labels.map((label) => ({ label, items: map.get(label) || [] }));
}

function getWorldFolderLabel(value) {
  const clean = normalizeWorldFolderName(value);
  return clean || "Unsorted";
}

function normalizeWorldFolderName(value) {
  return str(value).replace(/\s+/g, " ");
}

function isWorldCollection(collection) {
  return collection === "npcs" || collection === "quests" || collection === "locations";
}

function ensureWorldFolders() {
  if (!state.meta.worldFolders || typeof state.meta.worldFolders !== "object" || Array.isArray(state.meta.worldFolders)) {
    state.meta.worldFolders = { npcs: [], quests: [], locations: [] };
  }
  for (const collection of ["npcs", "quests", "locations"]) {
    const current = Array.isArray(state.meta.worldFolders[collection]) ? state.meta.worldFolders[collection] : [];
    const seen = new Set();
    const cleaned = [];
    for (const raw of current) {
      const name = normalizeWorldFolderName(raw);
      if (!name) continue;
      const key = name.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      cleaned.push(name);
    }
    for (const item of state[collection] || []) {
      const entityFolder = normalizeWorldFolderName(item?.folder);
      if (!entityFolder) continue;
      const key = entityFolder.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      cleaned.push(entityFolder);
    }
    cleaned.sort((a, b) => a.localeCompare(b));
    state.meta.worldFolders[collection] = cleaned;
  }
  return state.meta.worldFolders;
}

function getWorldFolders(collection) {
  if (!isWorldCollection(collection)) return [];
  const folders = ensureWorldFolders();
  return folders[collection] || [];
}

function renderWorldFolderOptions(collection, selectedValue = "", includeUnsorted = true) {
  const folders = getWorldFolders(collection);
  const selected = normalizeWorldFolderName(selectedValue);
  const options = [];
  if (includeUnsorted) {
    options.push(`<option value="" ${selected ? "" : "selected"}>Unsorted</option>`);
  }
  for (const folder of folders) {
    options.push(`<option value="${escapeHtml(folder)}" ${folder.toLowerCase() === selected.toLowerCase() ? "selected" : ""}>${escapeHtml(folder)}</option>`);
  }
  return options.join("");
}

function addWorldFolder(collection, folderName) {
  if (!isWorldCollection(collection)) return { ok: false, message: "Unknown world collection." };
  const clean = normalizeWorldFolderName(folderName);
  if (!clean) return { ok: false, message: "Folder name is required." };
  const folders = getWorldFolders(collection);
  const exists = folders.some((folder) => folder.toLowerCase() === clean.toLowerCase());
  if (exists) return { ok: true, message: `Folder "${clean}" already exists.` };
  folders.push(clean);
  folders.sort((a, b) => a.localeCompare(b));
  state.meta.worldFolders[collection] = folders;
  saveState();
  return { ok: true, message: `Added folder "${clean}".` };
}

function addWorldFolderFromDraft(collection) {
  if (!isWorldCollection(collection)) return;
  const draftInput = appEl.querySelector(`[data-world-folder-draft="${collection}"]`);
  const fromDom = draftInput instanceof HTMLInputElement ? draftInput.value : "";
  const clean = normalizeWorldFolderName(fromDom || ui.worldFolderDraft?.[collection] || "");
  if (!clean) {
    ui.worldMessages[collection] = "Type a folder name first (example: Rivergate).";
    render();
    return;
  }
  const result = addWorldFolder(collection, clean);
  ui.worldMessages[collection] = result.message;
  if (result.ok && ui.worldNewFolder && collection in ui.worldNewFolder) {
    ui.worldNewFolder[collection] = clean;
  }
  if (result.ok && ui.worldFolderDraft && collection in ui.worldFolderDraft) {
    ui.worldFolderDraft[collection] = "";
  }
  render();
}

function getSelectedWorldEntry(collection, items) {
  if (!Array.isArray(items) || !items.length) {
    if (ui.worldSelection && collection in ui.worldSelection) {
      ui.worldSelection[collection] = "";
    }
    return null;
  }
  if (!ui.worldSelection || typeof ui.worldSelection !== "object") {
    ui.worldSelection = { npcs: "", quests: "", locations: "" };
  }
  const selectedId = str(ui.worldSelection[collection]);
  const found = items.find((item) => item.id === selectedId);
  if (found) return found;
  ui.worldSelection[collection] = items[0].id;
  return items[0];
}

function setWorldSelection(collection, id) {
  if (!ui.worldSelection || typeof ui.worldSelection !== "object") {
    ui.worldSelection = { npcs: "", quests: "", locations: "" };
  }
  if (!(collection in ui.worldSelection)) return;
  ui.worldSelection[collection] = id;
  render();
}

function renderNpcDetails(npc) {
  const folderOptions = renderWorldFolderOptions("npcs", npc.folder, true);
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(npc.name || "Unnamed NPC")}</span>
        <span class="entry-meta">${escapeHtml(npc.disposition || "No disposition")}</span>
      </div>
      <div class="row">
        <label>Folder
          <select data-collection="npcs" data-id="${npc.id}" data-field="folder">
            ${folderOptions}
          </select>
        </label>
        <label>Name
          <input data-collection="npcs" data-id="${npc.id}" data-field="name" value="${escapeHtml(npc.name || "")}" />
        </label>
      </div>
      <div class="row">
        <label>Role
          <input data-collection="npcs" data-id="${npc.id}" data-field="role" value="${escapeHtml(npc.role || "")}" />
        </label>
        <label>Agenda
          <input data-collection="npcs" data-id="${npc.id}" data-field="agenda" value="${escapeHtml(npc.agenda || "")}" />
        </label>
        <label>Disposition
          <input data-collection="npcs" data-id="${npc.id}" data-field="disposition" value="${escapeHtml(npc.disposition || "")}" />
        </label>
      </div>
      <label>Notes
        <textarea data-collection="npcs" data-id="${npc.id}" data-field="notes">${escapeHtml(npc.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="npcs" data-id="${npc.id}">Delete NPC</button>
      </div>
    </article>
  `;
}

function renderQuestDetails(quest) {
  const folderOptions = renderWorldFolderOptions("quests", quest.folder, true);
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(quest.title || "Untitled Quest")}</span>
        <span class="entry-meta">${escapeHtml(quest.giver || "No giver")}</span>
      </div>
      <div class="row">
        <label>Folder
          <select data-collection="quests" data-id="${quest.id}" data-field="folder">
            ${folderOptions}
          </select>
        </label>
        <label>Title
          <input data-collection="quests" data-id="${quest.id}" data-field="title" value="${escapeHtml(quest.title || "")}" />
        </label>
      </div>
      <div class="row">
        <label>Status
          <select data-collection="quests" data-id="${quest.id}" data-field="status">
            ${["open", "in-progress", "blocked", "completed", "failed"]
              .map((status) => `<option value="${status}" ${quest.status === status ? "selected" : ""}>${status}</option>`)
              .join("")}
          </select>
        </label>
        <label>Quest Giver
          <input data-collection="quests" data-id="${quest.id}" data-field="giver" value="${escapeHtml(quest.giver || "")}" />
        </label>
        <label>Stakes
          <input data-collection="quests" data-id="${quest.id}" data-field="stakes" value="${escapeHtml(quest.stakes || "")}" />
        </label>
      </div>
      <label>Objective
        <textarea data-collection="quests" data-id="${quest.id}" data-field="objective">${escapeHtml(quest.objective || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="quests" data-id="${quest.id}">Delete Quest</button>
      </div>
    </article>
  `;
}

function renderLocationDetails(location) {
  const folderOptions = renderWorldFolderOptions("locations", location.folder, true);
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(location.name || "Unnamed Location")}</span>
        <span class="entry-meta">${escapeHtml(location.hex || "No hex")}</span>
      </div>
      <div class="row">
        <label>Folder
          <select data-collection="locations" data-id="${location.id}" data-field="folder">
            ${folderOptions}
          </select>
        </label>
        <label>Name
          <input data-collection="locations" data-id="${location.id}" data-field="name" value="${escapeHtml(location.name || "")}" />
        </label>
      </div>
      <label>Hex / Region
        <input data-collection="locations" data-id="${location.id}" data-field="hex" value="${escapeHtml(location.hex || "")}" />
      </label>
      <label>What Changed
        <textarea data-collection="locations" data-id="${location.id}" data-field="whatChanged">${escapeHtml(location.whatChanged || "")}</textarea>
      </label>
      <label>Notes
        <textarea data-collection="locations" data-id="${location.id}" data-field="notes">${escapeHtml(location.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="locations" data-id="${location.id}">Delete Location</button>
      </div>
    </article>
  `;
}

async function loadKingdomRulesData() {
  try {
    const response = await fetch(new URL("./kingdom-rules-data.json", import.meta.url));
    if (!response.ok) throw new Error(`Unable to load kingdom rules data (${response.status}).`);
    const parsed = await response.json();
    if (!Array.isArray(parsed?.profiles) || !parsed.profiles.length) throw new Error("No kingdom rules profiles found.");
    return parsed;
  } catch {
    return createFallbackKingdomRulesData();
  }
}

function createFallbackKingdomRulesData() {
  return {
    loadedAt: new Date().toISOString().slice(0, 10),
    latestProfileId: "fallback",
    profiles: [
      {
        id: "fallback",
        label: "Kingdom Rules Profile",
        shortLabel: "Kingdom",
        version: "local-fallback",
        status: "fallback",
        summary: "Fallback kingdom rules profile used because the shared rules data file could not be loaded.",
        sources: [],
        automationNotes: [],
        quickStart: ["Load the rules data file to unlock the full kingdom guide."],
        turnStructure: [
          { phase: "Upkeep", summary: "Review current kingdom state." },
          { phase: "Activities", summary: "Assign actions and record outcomes." },
          { phase: "Event", summary: "Resolve the kingdom event." }
        ],
        actionLimits: [],
        creationChanges: [],
        advancement: [],
        mathAdjustments: [],
        leadershipRules: [],
        leadershipRoles: [],
        economyAndXP: [],
        activitiesAdded: [],
        activitiesChanged: [],
        settlementRules: [],
        constructionRules: [],
        structureBonuses: [],
        clarifications: [],
        aiContextSummary: [],
        helpPrompts: []
      }
    ]
  };
}

function getKingdomRulesProfiles() {
  return Array.isArray(kingdomRulesData?.profiles) ? kingdomRulesData.profiles : [];
}

function getDefaultKingdomProfileId() {
  const profiles = getKingdomRulesProfiles();
  const wanted = str(kingdomRulesData?.latestProfileId);
  if (wanted && profiles.some((profile) => profile.id === wanted)) return wanted;
  return profiles[0]?.id || "fallback";
}

function getKingdomProfileById(profileId) {
  const clean = str(profileId);
  return getKingdomRulesProfiles().find((profile) => str(profile?.id) === clean) || getKingdomRulesProfiles()[0] || null;
}

function getActiveKingdomProfile() {
  return getKingdomProfileById(state?.kingdom?.profileId || getDefaultKingdomProfileId());
}

function getControlDcForLevel(profile, level) {
  const normalizedLevel = Math.max(1, Number.parseInt(String(level || "1"), 10) || 1);
  const table = Array.isArray(profile?.advancement) ? profile.advancement : [];
  return table.find((entry) => Number.parseInt(String(entry?.level || "0"), 10) === normalizedLevel)?.controlDC || 14;
}

function createStarterKingdomState() {
  const profile = getKingdomProfileById(getDefaultKingdomProfileId());
  return {
    profileId: profile?.id || getDefaultKingdomProfileId(),
    name: "Stolen Lands Charter",
    charter: "Open charter",
    government: "Council",
    heartland: "Grassland",
    capital: "TBD",
    currentTurnLabel: "Turn 1",
    currentDate: "",
    level: 1,
    size: 1,
    controlDC: getControlDcForLevel(profile, 1),
    resourceDie: "d4",
    resourcePoints: 0,
    xp: 0,
    trainedSkills: ["Agriculture", "Politics", "Trade", "Wilderness"],
    abilities: {
      culture: 0,
      economy: 0,
      loyalty: 0,
      stability: 0
    },
    commodities: {
      food: 0,
      lumber: 0,
      luxuries: 0,
      ore: 0,
      stone: 0
    },
    consumption: 0,
    renown: 1,
    fame: 0,
    infamy: 0,
    unrest: 0,
    ruin: {
      corruption: 0,
      crime: 0,
      decay: 0,
      strife: 0,
      threshold: 5
    },
    notes: "Track capital growth, local influence, and construction queue here.",
    leaders: [
      {
        id: uid(),
        role: "Ruler",
        name: "Unassigned",
        type: "PC",
        leadershipBonus: 1,
        relevantSkills: "Diplomacy, Politics Lore",
        specializedSkills: "Industry, Politics, Statecraft",
        notes: "Set once the party chooses invested roles."
      }
    ],
    settlements: [
      {
        id: uid(),
        name: "Capital Site",
        size: "Village",
        influence: 1,
        civicStructure: "Town Hall",
        resourceDice: 0,
        consumption: 0,
        notes: "First permanent seat of government."
      }
    ],
    regions: [
      {
        id: uid(),
        hex: "A1",
        status: "Claimed",
        terrain: "Plains",
        workSite: "",
        notes: "Starting heartland."
      }
    ],
    turns: [],
    pendingProjects: [
      "Choose all eight leadership roles.",
      "Settle final charter / government / heartland choices in the sheet.",
      "Create the first real settlement record once the capital is founded."
    ]
  };
}

function normalizeKingdomState(input) {
  const base = createStarterKingdomState();
  const out = {
    ...base,
    ...(input && typeof input === "object" ? input : {})
  };
  out.profileId = str(out.profileId) || base.profileId;
  out.name = str(out.name) || base.name;
  out.charter = str(out.charter);
  out.government = str(out.government);
  out.heartland = str(out.heartland);
  out.capital = str(out.capital);
  out.currentTurnLabel = str(out.currentTurnLabel) || "Turn 1";
  out.currentDate = str(out.currentDate);
  out.level = Math.max(1, Number.parseInt(String(out.level || "1"), 10) || 1);
  out.size = Math.max(1, Number.parseInt(String(out.size || "1"), 10) || 1);
  out.controlDC = Math.max(10, Number.parseInt(String(out.controlDC || getControlDcForLevel(getKingdomProfileById(out.profileId), out.level)), 10) || 14);
  out.resourceDie = ["d4", "d6", "d8", "d10", "d12"].includes(str(out.resourceDie)) ? str(out.resourceDie) : "d4";
  out.resourcePoints = Number.parseInt(String(out.resourcePoints || "0"), 10) || 0;
  out.xp = Number.parseInt(String(out.xp || "0"), 10) || 0;
  out.trainedSkills = Array.isArray(out.trainedSkills)
    ? out.trainedSkills.map((skill) => str(skill)).filter(Boolean)
    : str(out.trainedSkills)
        .split(",")
        .map((skill) => str(skill).trim())
        .filter(Boolean);
  out.abilities = {
    culture: Number.parseInt(String(out?.abilities?.culture || "0"), 10) || 0,
    economy: Number.parseInt(String(out?.abilities?.economy || "0"), 10) || 0,
    loyalty: Number.parseInt(String(out?.abilities?.loyalty || "0"), 10) || 0,
    stability: Number.parseInt(String(out?.abilities?.stability || "0"), 10) || 0
  };
  out.commodities = {
    food: Number.parseInt(String(out?.commodities?.food || "0"), 10) || 0,
    lumber: Number.parseInt(String(out?.commodities?.lumber || "0"), 10) || 0,
    luxuries: Number.parseInt(String(out?.commodities?.luxuries || "0"), 10) || 0,
    ore: Number.parseInt(String(out?.commodities?.ore || "0"), 10) || 0,
    stone: Number.parseInt(String(out?.commodities?.stone || "0"), 10) || 0
  };
  out.consumption = Math.max(0, Number.parseInt(String(out.consumption || "0"), 10) || 0);
  out.renown = Math.max(0, Number.parseInt(String(out.renown || "0"), 10) || 0);
  out.fame = Math.max(0, Number.parseInt(String(out.fame || "0"), 10) || 0);
  out.infamy = Math.max(0, Number.parseInt(String(out.infamy || "0"), 10) || 0);
  out.unrest = Math.max(0, Number.parseInt(String(out.unrest || "0"), 10) || 0);
  out.ruin = {
    corruption: Math.max(0, Number.parseInt(String(out?.ruin?.corruption || "0"), 10) || 0),
    crime: Math.max(0, Number.parseInt(String(out?.ruin?.crime || "0"), 10) || 0),
    decay: Math.max(0, Number.parseInt(String(out?.ruin?.decay || "0"), 10) || 0),
    strife: Math.max(0, Number.parseInt(String(out?.ruin?.strife || "0"), 10) || 0),
    threshold: Math.max(1, Number.parseInt(String(out?.ruin?.threshold || "5"), 10) || 5)
  };
  out.notes = str(out.notes);
  out.pendingProjects = Array.isArray(out.pendingProjects) ? out.pendingProjects.map((entry) => str(entry)).filter(Boolean) : [];
  out.leaders = Array.isArray(out.leaders) ? out.leaders.map((leader) => ({ ...leader, id: str(leader?.id) || uid(), updatedAt: str(leader?.updatedAt) || "" })) : [];
  out.settlements = Array.isArray(out.settlements)
    ? out.settlements.map((settlement) => ({ ...settlement, id: str(settlement?.id) || uid(), updatedAt: str(settlement?.updatedAt) || "" }))
    : [];
  out.regions = Array.isArray(out.regions) ? out.regions.map((region) => ({ ...region, id: str(region?.id) || uid(), updatedAt: str(region?.updatedAt) || "" })) : [];
  out.turns = Array.isArray(out.turns) ? out.turns.map((turn) => ({ ...turn, id: str(turn?.id) || uid(), updatedAt: str(turn?.updatedAt) || "" })) : [];
  return out;
}

function getKingdomState() {
  if (!state.kingdom || typeof state.kingdom !== "object" || Array.isArray(state.kingdom)) {
    state.kingdom = createStarterKingdomState();
  }
  return state.kingdom;
}

function buildKingdomAiContext(kingdom, profile) {
  const data = kingdom && typeof kingdom === "object" ? kingdom : getKingdomState();
  const rulesProfile = profile || getActiveKingdomProfile();
  return {
    name: data.name,
    currentTurnLabel: data.currentTurnLabel,
    currentDate: data.currentDate,
    level: data.level,
    size: data.size,
    controlDC: data.controlDC,
    resourceDie: data.resourceDie,
    resourcePoints: data.resourcePoints,
    trainedSkills: [...(data.trainedSkills || [])].slice(0, 16),
    abilities: { ...(data.abilities || {}) },
    commodities: { ...(data.commodities || {}) },
    consumption: data.consumption,
    renown: data.renown,
    fame: data.fame,
    infamy: data.infamy,
    unrest: data.unrest,
    ruin: { ...(data.ruin || {}) },
    notes: str(data.notes).slice(0, 900),
    pendingProjects: [...(data.pendingProjects || [])].slice(0, 8),
    leaders: (data.leaders || []).slice(0, 8),
    settlements: (data.settlements || []).slice(0, 8),
    regions: (data.regions || []).slice(0, 10),
    recentTurns: (data.turns || []).slice(0, 6),
    rulesProfile: {
      id: rulesProfile?.id || "",
      label: rulesProfile?.label || "",
      summary: str(rulesProfile?.summary || "").slice(0, 420),
      turnStructure: (rulesProfile?.turnStructure || []).map((entry) => `${entry.phase}: ${entry.summary}`).slice(0, 5),
      aiSummary: [...(rulesProfile?.aiContextSummary || [])].slice(0, 8)
    }
  };
}

function applyKingdomOverviewForm(fields) {
  const kingdom = getKingdomState();
  const profileId = str(fields.profileId) || kingdom.profileId || getDefaultKingdomProfileId();
  const profile = getKingdomProfileById(profileId);
  kingdom.profileId = profile?.id || getDefaultKingdomProfileId();
  kingdom.name = str(fields.name);
  kingdom.charter = str(fields.charter);
  kingdom.government = str(fields.government);
  kingdom.heartland = str(fields.heartland);
  kingdom.capital = str(fields.capital);
  kingdom.currentTurnLabel = str(fields.currentTurnLabel) || kingdom.currentTurnLabel;
  kingdom.currentDate = str(fields.currentDate);
  kingdom.level = Math.max(1, Number.parseInt(String(fields.level || kingdom.level || "1"), 10) || 1);
  kingdom.size = Math.max(1, Number.parseInt(String(fields.size || kingdom.size || "1"), 10) || 1);
  kingdom.controlDC = Math.max(
    10,
    Number.parseInt(String(fields.controlDC || getControlDcForLevel(profile, kingdom.level)), 10) || getControlDcForLevel(profile, kingdom.level)
  );
  kingdom.resourceDie = ["d4", "d6", "d8", "d10", "d12"].includes(str(fields.resourceDie)) ? str(fields.resourceDie) : kingdom.resourceDie;
  kingdom.resourcePoints = Number.parseInt(String(fields.resourcePoints || kingdom.resourcePoints || "0"), 10) || 0;
  kingdom.xp = Number.parseInt(String(fields.xp || kingdom.xp || "0"), 10) || 0;
  kingdom.trainedSkills = str(fields.trainedSkills)
    .split(",")
    .map((skill) => str(skill).trim())
    .filter(Boolean);
  kingdom.abilities = {
    culture: Number.parseInt(String(fields.culture || kingdom.abilities.culture || "0"), 10) || 0,
    economy: Number.parseInt(String(fields.economy || kingdom.abilities.economy || "0"), 10) || 0,
    loyalty: Number.parseInt(String(fields.loyalty || kingdom.abilities.loyalty || "0"), 10) || 0,
    stability: Number.parseInt(String(fields.stability || kingdom.abilities.stability || "0"), 10) || 0
  };
  kingdom.commodities = {
    food: Number.parseInt(String(fields.food || kingdom.commodities.food || "0"), 10) || 0,
    lumber: Number.parseInt(String(fields.lumber || kingdom.commodities.lumber || "0"), 10) || 0,
    luxuries: Number.parseInt(String(fields.luxuries || kingdom.commodities.luxuries || "0"), 10) || 0,
    ore: Number.parseInt(String(fields.ore || kingdom.commodities.ore || "0"), 10) || 0,
    stone: Number.parseInt(String(fields.stone || kingdom.commodities.stone || "0"), 10) || 0
  };
  kingdom.consumption = Math.max(0, Number.parseInt(String(fields.consumption || kingdom.consumption || "0"), 10) || 0);
  kingdom.renown = Math.max(0, Number.parseInt(String(fields.renown || kingdom.renown || "0"), 10) || 0);
  kingdom.fame = Math.max(0, Number.parseInt(String(fields.fame || kingdom.fame || "0"), 10) || 0);
  kingdom.infamy = Math.max(0, Number.parseInt(String(fields.infamy || kingdom.infamy || "0"), 10) || 0);
  kingdom.unrest = Math.max(0, Number.parseInt(String(fields.unrest || kingdom.unrest || "0"), 10) || 0);
  kingdom.ruin = {
    corruption: Math.max(0, Number.parseInt(String(fields.corruption || kingdom.ruin.corruption || "0"), 10) || 0),
    crime: Math.max(0, Number.parseInt(String(fields.crime || kingdom.ruin.crime || "0"), 10) || 0),
    decay: Math.max(0, Number.parseInt(String(fields.decay || kingdom.ruin.decay || "0"), 10) || 0),
    strife: Math.max(0, Number.parseInt(String(fields.strife || kingdom.ruin.strife || "0"), 10) || 0),
    threshold: Math.max(1, Number.parseInt(String(fields.ruinThreshold || kingdom.ruin.threshold || "5"), 10) || 5)
  };
  kingdom.notes = str(fields.notes);
}

function createKingdomLeader(fields) {
  const kingdom = getKingdomState();
  kingdom.leaders.unshift({
    id: uid(),
    role: str(fields.role) || "Leader",
    name: str(fields.name) || "Unnamed leader",
    type: str(fields.type) || "NPC",
    leadershipBonus: Number.parseInt(String(fields.leadershipBonus || "0"), 10) || 0,
    relevantSkills: str(fields.relevantSkills),
    specializedSkills: str(fields.specializedSkills),
    notes: str(fields.notes),
    updatedAt: new Date().toISOString()
  });
}

function createKingdomSettlement(fields) {
  const kingdom = getKingdomState();
  kingdom.settlements.unshift({
    id: uid(),
    name: str(fields.name) || "Unnamed settlement",
    size: str(fields.size) || "Village",
    influence: Math.max(0, Number.parseInt(String(fields.influence || "0"), 10) || 0),
    civicStructure: str(fields.civicStructure),
    resourceDice: Math.max(0, Number.parseInt(String(fields.resourceDice || "0"), 10) || 0),
    consumption: Math.max(0, Number.parseInt(String(fields.consumption || "0"), 10) || 0),
    notes: str(fields.notes),
    updatedAt: new Date().toISOString()
  });
}

function createKingdomRegion(fields) {
  const kingdom = getKingdomState();
  kingdom.regions.unshift({
    id: uid(),
    hex: str(fields.hex) || "Unknown hex",
    status: str(fields.status) || "Claimed",
    terrain: str(fields.terrain),
    workSite: str(fields.workSite),
    notes: str(fields.notes),
    updatedAt: new Date().toISOString()
  });
}

function applyKingdomTurnForm(fields) {
  const kingdom = getKingdomState();
  const title = str(fields.title) || `Turn ${kingdom.turns.length + 1}`;
  const rpDelta = Number.parseInt(String(fields.rpDelta || "0"), 10) || 0;
  const unrestDelta = Number.parseInt(String(fields.unrestDelta || "0"), 10) || 0;
  const renownDelta = Number.parseInt(String(fields.renownDelta || "0"), 10) || 0;
  const fameDelta = Number.parseInt(String(fields.fameDelta || "0"), 10) || 0;
  const infamyDelta = Number.parseInt(String(fields.infamyDelta || "0"), 10) || 0;
  const corruptionDelta = Number.parseInt(String(fields.corruptionDelta || "0"), 10) || 0;
  const crimeDelta = Number.parseInt(String(fields.crimeDelta || "0"), 10) || 0;
  const decayDelta = Number.parseInt(String(fields.decayDelta || "0"), 10) || 0;
  const strifeDelta = Number.parseInt(String(fields.strifeDelta || "0"), 10) || 0;
  const foodDelta = Number.parseInt(String(fields.foodDelta || "0"), 10) || 0;
  const lumberDelta = Number.parseInt(String(fields.lumberDelta || "0"), 10) || 0;
  const luxuriesDelta = Number.parseInt(String(fields.luxuriesDelta || "0"), 10) || 0;
  const oreDelta = Number.parseInt(String(fields.oreDelta || "0"), 10) || 0;
  const stoneDelta = Number.parseInt(String(fields.stoneDelta || "0"), 10) || 0;
  kingdom.currentTurnLabel = title;
  kingdom.currentDate = str(fields.date) || kingdom.currentDate;
  kingdom.resourcePoints += rpDelta;
  kingdom.unrest = Math.max(0, kingdom.unrest + unrestDelta);
  kingdom.renown = Math.max(0, kingdom.renown + renownDelta);
  kingdom.fame = Math.max(0, kingdom.fame + fameDelta);
  kingdom.infamy = Math.max(0, kingdom.infamy + infamyDelta);
  kingdom.ruin.corruption = Math.max(0, kingdom.ruin.corruption + corruptionDelta);
  kingdom.ruin.crime = Math.max(0, kingdom.ruin.crime + crimeDelta);
  kingdom.ruin.decay = Math.max(0, kingdom.ruin.decay + decayDelta);
  kingdom.ruin.strife = Math.max(0, kingdom.ruin.strife + strifeDelta);
  kingdom.commodities.food += foodDelta;
  kingdom.commodities.lumber += lumberDelta;
  kingdom.commodities.luxuries += luxuriesDelta;
  kingdom.commodities.ore += oreDelta;
  kingdom.commodities.stone += stoneDelta;
  kingdom.turns.unshift({
    id: uid(),
    title,
    date: str(fields.date),
    rpDelta,
    unrestDelta,
    renownDelta,
    fameDelta,
    infamyDelta,
    summary: str(fields.summary),
    risks: str(fields.risks),
    updatedAt: new Date().toISOString()
  });
  const latest = getLatestSession();
  if (latest) {
    latest.kingdomTurn = title;
    latest.updatedAt = new Date().toISOString();
  }
  const pending = str(fields.pendingProject);
  if (pending) {
    kingdom.pendingProjects.unshift(pending);
    kingdom.pendingProjects = [...new Set(kingdom.pendingProjects.map((entry) => str(entry)).filter(Boolean))].slice(0, 16);
  }
}

function appendKingdomAiNote(text) {
  const kingdom = getKingdomState();
  const stamp = new Date().toLocaleString();
  const block = `[AI ${stamp}]\n${str(text)}`;
  kingdom.notes = kingdom.notes ? `${kingdom.notes}\n\n${block}`.trim() : block;
  saveState();
}

function renderKingdom() {
  const kingdom = getKingdomState();
  const profile = getActiveKingdomProfile();
  const sourceLines = Array.isArray(profile?.sources)
    ? profile.sources
        .map((source) => `${source.title}${source.role ? ` (${source.role})` : ""}`)
        .filter(Boolean)
    : [];
  return `
    <div class="page-stack kingdom-page">
      ${renderPageIntro(
        "Kingdom",
        "Track the kingdom sheet, leaders, settlements, regions, turn flow, and the active kingdom-rules profile in one place."
      )}
      <section class="panel flow-panel">
        <div class="entry-head">
          <div>
            <h2 style="margin:0;">Rules Profile</h2>
            <p class="small" style="margin:6px 0 0;">${escapeHtml(profile?.summary || "No kingdom rules profile loaded.")}</p>
          </div>
          <div class="kingdom-profile-pill">${escapeHtml(profile?.shortLabel || profile?.label || "Unknown profile")}</div>
        </div>
        ${ui.kingdomMessage ? `<p class="small">${escapeHtml(ui.kingdomMessage)}</p>` : ""}
        <div class="kingdom-chip-row">
          <span class="chip">Version ${escapeHtml(str(profile?.version || "unknown"))}</span>
          <span class="chip">Turn ${escapeHtml(kingdom.currentTurnLabel || "Turn 1")}</span>
          <span class="chip">Control DC ${escapeHtml(String(kingdom.controlDC || 14))}</span>
          <span class="chip">Resource Die ${escapeHtml(kingdom.resourceDie || "d4")}</span>
          <span class="chip">Settlements ${escapeHtml(String(kingdom.settlements.length))}</span>
          <span class="chip">Regions ${escapeHtml(String(kingdom.regions.length))}</span>
        </div>
        ${
          sourceLines.length
            ? `
              <details class="kingdom-guide-panel">
                <summary>Profile Sources</summary>
                <ul class="flow-list">
                  ${sourceLines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}
                </ul>
              </details>
            `
            : ""
        }
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">1</span>
          <h2>Kingdom Sheet</h2>
        </div>
        <form data-form="kingdom-overview">
          <div class="row">
            <label>Rules Profile
              <select name="profileId">
                ${getKingdomRulesProfiles()
                  .map(
                    (entry) =>
                      `<option value="${escapeHtml(entry.id)}" ${entry.id === kingdom.profileId ? "selected" : ""}>${escapeHtml(entry.label)}</option>`
                  )
                  .join("")}
              </select>
            </label>
            <label>Kingdom Name
              <input name="name" value="${escapeHtml(kingdom.name || "")}" placeholder="Stolen Lands Charter" />
            </label>
            <label>Capital
              <input name="capital" value="${escapeHtml(kingdom.capital || "")}" placeholder="Tuskfall" />
            </label>
          </div>
          <div class="row">
            <label>Charter
              <input name="charter" value="${escapeHtml(kingdom.charter || "")}" placeholder="Open charter" />
            </label>
            <label>Government
              <input name="government" value="${escapeHtml(kingdom.government || "")}" placeholder="Council" />
            </label>
            <label>Heartland
              <input name="heartland" value="${escapeHtml(kingdom.heartland || "")}" placeholder="Grassland" />
            </label>
          </div>
          <div class="row">
            <label>Current Turn Label
              <input name="currentTurnLabel" value="${escapeHtml(kingdom.currentTurnLabel || "")}" placeholder="Turn 3" />
            </label>
            <label>Current Date
              <input name="currentDate" value="${escapeHtml(kingdom.currentDate || "")}" placeholder="4712-09-01" />
            </label>
            <label>Level
              <input name="level" type="number" min="1" max="20" value="${escapeHtml(String(kingdom.level || 1))}" />
            </label>
            <label>Size
              <input name="size" type="number" min="1" value="${escapeHtml(String(kingdom.size || 1))}" />
            </label>
          </div>
          <div class="row">
            <label>Control DC
              <input name="controlDC" type="number" min="10" value="${escapeHtml(String(kingdom.controlDC || 14))}" />
            </label>
            <label>Resource Die
              <select name="resourceDie">
                ${["d4", "d6", "d8", "d10", "d12"]
                  .map((die) => `<option value="${die}" ${kingdom.resourceDie === die ? "selected" : ""}>${die}</option>`)
                  .join("")}
              </select>
            </label>
            <label>Resource Points
              <input name="resourcePoints" type="number" value="${escapeHtml(String(kingdom.resourcePoints || 0))}" />
            </label>
            <label>XP
              <input name="xp" type="number" value="${escapeHtml(String(kingdom.xp || 0))}" />
            </label>
          </div>
          <div class="row">
            <label>Culture
              <input name="culture" type="number" value="${escapeHtml(String(kingdom.abilities.culture || 0))}" />
            </label>
            <label>Economy
              <input name="economy" type="number" value="${escapeHtml(String(kingdom.abilities.economy || 0))}" />
            </label>
            <label>Loyalty
              <input name="loyalty" type="number" value="${escapeHtml(String(kingdom.abilities.loyalty || 0))}" />
            </label>
            <label>Stability
              <input name="stability" type="number" value="${escapeHtml(String(kingdom.abilities.stability || 0))}" />
            </label>
          </div>
          <div class="row">
            <label>Food
              <input name="food" type="number" value="${escapeHtml(String(kingdom.commodities.food || 0))}" />
            </label>
            <label>Lumber
              <input name="lumber" type="number" value="${escapeHtml(String(kingdom.commodities.lumber || 0))}" />
            </label>
            <label>Luxuries
              <input name="luxuries" type="number" value="${escapeHtml(String(kingdom.commodities.luxuries || 0))}" />
            </label>
            <label>Ore
              <input name="ore" type="number" value="${escapeHtml(String(kingdom.commodities.ore || 0))}" />
            </label>
            <label>Stone
              <input name="stone" type="number" value="${escapeHtml(String(kingdom.commodities.stone || 0))}" />
            </label>
          </div>
          <div class="row">
            <label>Consumption
              <input name="consumption" type="number" min="0" value="${escapeHtml(String(kingdom.consumption || 0))}" />
            </label>
            <label>Renown
              <input name="renown" type="number" min="0" value="${escapeHtml(String(kingdom.renown || 0))}" />
            </label>
            <label>Fame
              <input name="fame" type="number" min="0" value="${escapeHtml(String(kingdom.fame || 0))}" />
            </label>
            <label>Infamy
              <input name="infamy" type="number" min="0" value="${escapeHtml(String(kingdom.infamy || 0))}" />
            </label>
            <label>Unrest
              <input name="unrest" type="number" min="0" value="${escapeHtml(String(kingdom.unrest || 0))}" />
            </label>
          </div>
          <div class="row">
            <label>Corruption
              <input name="corruption" type="number" min="0" value="${escapeHtml(String(kingdom.ruin.corruption || 0))}" />
            </label>
            <label>Crime
              <input name="crime" type="number" min="0" value="${escapeHtml(String(kingdom.ruin.crime || 0))}" />
            </label>
            <label>Decay
              <input name="decay" type="number" min="0" value="${escapeHtml(String(kingdom.ruin.decay || 0))}" />
            </label>
            <label>Strife
              <input name="strife" type="number" min="0" value="${escapeHtml(String(kingdom.ruin.strife || 0))}" />
            </label>
            <label>Ruin Threshold
              <input name="ruinThreshold" type="number" min="1" value="${escapeHtml(String(kingdom.ruin.threshold || 5))}" />
            </label>
          </div>
          <label>Trained Skills (comma separated)
            <input name="trainedSkills" value="${escapeHtml((kingdom.trainedSkills || []).join(", "))}" placeholder="Agriculture, Politics, Trade, Wilderness" />
          </label>
          <label>Kingdom Notes
            <textarea name="notes" placeholder="Track active plans, open rulings, and the state of the kingdom here.">${escapeHtml(kingdom.notes || "")}</textarea>
          </label>
          <div class="toolbar">
            <button class="btn btn-primary" type="submit">Save Kingdom Sheet</button>
          </div>
        </form>
      </section>

      <section class="kingdom-overview-grid">
        <article class="panel">
          <h2>Leaders</h2>
          <form data-form="kingdom-leader">
            <div class="row">
              <label>Role
                <select name="role">
                  ${["Ruler", "Counselor", "Emissary", "General", "Magister", "Treasurer", "Viceroy", "Warden"]
                    .map((role) => `<option value="${role}">${role}</option>`)
                    .join("")}
                </select>
              </label>
              <label>Name
                <input name="name" placeholder="Amiri" />
              </label>
              <label>Type
                <select name="type">
                  <option value="PC">PC</option>
                  <option value="NPC">NPC</option>
                </select>
              </label>
              <label>Leadership Bonus
                <input name="leadershipBonus" type="number" min="0" max="4" value="1" />
              </label>
            </div>
            <label>Relevant Skills
              <input name="relevantSkills" placeholder="Diplomacy, Politics Lore" />
            </label>
            <label>Specialized Kingdom Skills
              <input name="specializedSkills" placeholder="Politics, Statecraft, Trade" />
            </label>
            <label>Notes
              <textarea name="notes" placeholder="Why this leader is good in this role, house rulings, companion details..."></textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Add Leader</button>
            </div>
          </form>
          <div class="card-list">
            ${kingdom.leaders.length ? kingdom.leaders.map((leader) => renderKingdomLeaderEntry(leader)).join("") : `<p class="empty">No leaders tracked yet.</p>`}
          </div>
        </article>

        <article class="panel">
          <h2>Settlements</h2>
          <form data-form="kingdom-settlement">
            <div class="row">
              <label>Name
                <input name="name" placeholder="Tuskfall" />
              </label>
              <label>Size
                <select name="size">
                  ${["Village", "Town", "City", "Metropolis"].map((size) => `<option value="${size}">${size}</option>`).join("")}
                </select>
              </label>
              <label>Influence
                <input name="influence" type="number" min="0" value="1" />
              </label>
            </div>
            <div class="row">
              <label>Civic Structure
                <select name="civicStructure">
                  ${["", "Town Hall", "Castle", "Palace"].map((value) => `<option value="${value}">${value || "None"}</option>`).join("")}
                </select>
              </label>
              <label>Resource Dice
                <input name="resourceDice" type="number" min="0" value="0" />
              </label>
              <label>Consumption
                <input name="consumption" type="number" min="0" value="0" />
              </label>
            </div>
            <label>Notes
              <textarea name="notes" placeholder="Infrastructure, civic limits, item bonuses, special buildings..."></textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Add Settlement</button>
            </div>
          </form>
          <div class="card-list">
            ${kingdom.settlements.length
              ? kingdom.settlements.map((settlement) => renderKingdomSettlementEntry(settlement)).join("")
              : `<p class="empty">No settlements tracked yet.</p>`}
          </div>
        </article>
      </section>

      <section class="kingdom-overview-grid">
        <article class="panel">
          <h2>Regions / Hexes</h2>
          <form data-form="kingdom-region">
            <div class="row">
              <label>Hex
                <input name="hex" placeholder="B3" />
              </label>
              <label>Status
                <select name="status">
                  ${["Claimed", "Reconnoitered", "Work Site", "Settlement", "Contested"].map((status) => `<option value="${status}">${status}</option>`).join("")}
                </select>
              </label>
              <label>Terrain
                <input name="terrain" placeholder="Forest" />
              </label>
              <label>Work Site
                <input name="workSite" placeholder="Lumber Camp" />
              </label>
            </div>
            <label>Notes
              <textarea name="notes" placeholder="Terrain features, refuge, danger, or why this hex matters."></textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Add Region Record</button>
            </div>
          </form>
          <div class="card-list">
            ${kingdom.regions.length ? kingdom.regions.map((region) => renderKingdomRegionEntry(region)).join("") : `<p class="empty">No regions tracked yet.</p>`}
          </div>
        </article>

        <article class="panel">
          <h2>Run Kingdom Turn</h2>
          <p class="small">Use the active rules profile to resolve the turn, then record the deltas here so the kingdom sheet stays current.</p>
          <form data-form="kingdom-turn">
            <div class="row">
              <label>Turn Title
                <input name="title" placeholder="Turn 4 - Harvest Preparations" />
              </label>
              <label>Date
                <input name="date" placeholder="4712-11-01" />
              </label>
              <label>Pending Project
                <input name="pendingProject" placeholder="Finish Town Hall foundation" />
              </label>
            </div>
            <div class="row">
              <label>RP Delta
                <input name="rpDelta" type="number" value="0" />
              </label>
              <label>Unrest Delta
                <input name="unrestDelta" type="number" value="0" />
              </label>
              <label>Renown Delta
                <input name="renownDelta" type="number" value="0" />
              </label>
              <label>Fame Delta
                <input name="fameDelta" type="number" value="0" />
              </label>
              <label>Infamy Delta
                <input name="infamyDelta" type="number" value="0" />
              </label>
            </div>
            <div class="row">
              <label>Food Delta
                <input name="foodDelta" type="number" value="0" />
              </label>
              <label>Lumber Delta
                <input name="lumberDelta" type="number" value="0" />
              </label>
              <label>Luxuries Delta
                <input name="luxuriesDelta" type="number" value="0" />
              </label>
              <label>Ore Delta
                <input name="oreDelta" type="number" value="0" />
              </label>
              <label>Stone Delta
                <input name="stoneDelta" type="number" value="0" />
              </label>
            </div>
            <div class="row">
              <label>Corruption Delta
                <input name="corruptionDelta" type="number" value="0" />
              </label>
              <label>Crime Delta
                <input name="crimeDelta" type="number" value="0" />
              </label>
              <label>Decay Delta
                <input name="decayDelta" type="number" value="0" />
              </label>
              <label>Strife Delta
                <input name="strifeDelta" type="number" value="0" />
              </label>
            </div>
            <label>Turn Summary
              <textarea name="summary" placeholder="What happened in Upkeep, Activities, Construction, and Event?"></textarea>
            </label>
            <label>Risks / Follow-Ups
              <textarea name="risks" placeholder="What needs attention next turn?"></textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Apply Kingdom Turn</button>
            </div>
          </form>
          <div class="card-list">
            ${kingdom.turns.length ? kingdom.turns.map((turn) => renderKingdomTurnEntry(turn)).join("") : `<p class="empty">No kingdom turns recorded yet.</p>`}
          </div>
        </article>
      </section>

      <section class="panel kingdom-guide-panel">
        <h2>Kingdom Guide</h2>
        <p class="small">This guide is built from the active rules profile so Loremaster and the kingdom sheet stay aligned.</p>
        ${renderKingdomGuide(profile, kingdom)}
      </section>
    </div>
  `;
}

function renderKingdomGuide(profile, kingdom) {
  const pcLeaders = (kingdom?.leaders || []).filter((leader) => str(leader?.type).toUpperCase() === "PC").length;
  const npcLeaders = Math.max(0, (kingdom?.leaders || []).length - pcLeaders);
  const sections = [
    renderKingdomGuideSection("Rules Stack", `
      <p class="small">${escapeHtml(profile?.summary || "No kingdom rules profile loaded.")}</p>
      ${renderKingdomGuideList(
        (profile?.sources || []).map((source) => `${source.title}${source.role ? ` (${source.role})` : ""}`),
        { className: "kingdom-source-list" }
      )}
    `),
    renderKingdomGuideSection("Quick Start", renderKingdomGuideList(profile?.quickStart || [], { ordered: true })),
    renderKingdomGuideSection(
      "Turn Structure",
      `
        <div class="kingdom-chip-row">
          <span class="chip">Current Turn ${escapeHtml(kingdom?.currentTurnLabel || "Turn 1")}</span>
          <span class="chip">PC Leader Actions ${escapeHtml(String(pcLeaders * 3))}</span>
          <span class="chip">NPC Leader Actions ${escapeHtml(String(npcLeaders * 2))}</span>
          <span class="chip">Pending Projects ${escapeHtml(String((kingdom?.pendingProjects || []).length))}</span>
        </div>
        <div class="card-list kingdom-guide-cards">
          ${(profile?.turnStructure || [])
            .map(
              (entry) => `
                <article class="entry">
                  <div class="entry-head">
                    <span class="entry-title">${escapeHtml(entry?.phase || "Phase")}</span>
                  </div>
                  <p>${escapeHtml(entry?.summary || "")}</p>
                </article>
              `
            )
            .join("")}
        </div>
      `
    ),
    renderKingdomGuideSection(
      "Action Economy",
      `
        <p class="small">The remastered profile collapses turn actions into an activity economy: each PC leader gets 3 actions, each NPC leader gets 2, and civic structures can add settlement actions.</p>
        ${renderKingdomGuideList(profile?.actionLimits || [])}
      `
    ),
    renderKingdomGuideSection("Kingdom Creation", renderKingdomGuideList(profile?.creationChanges || [])),
    renderKingdomGuideSection("Math And Scaling", renderKingdomGuideList(profile?.mathAdjustments || [])),
    renderKingdomGuideSection("Leadership Rules", renderKingdomGuideList(profile?.leadershipRules || [])),
    renderKingdomGuideSection(
      "Leadership Roles",
      `
        <div class="card-list kingdom-guide-cards">
          ${(profile?.leadershipRoles || [])
            .map(
              (role) => `
                <article class="entry">
                  <div class="entry-head">
                    <span class="entry-title">${escapeHtml(role?.role || "Role")}</span>
                  </div>
                  <p><strong>Relevant Skills:</strong> ${escapeHtml((role?.relevantSkills || []).join(", ") || "None listed.")}</p>
                  <p><strong>Specialized Kingdom Skills:</strong> ${escapeHtml((role?.specializedSkills || []).join(", ") || "None listed.")}</p>
                </article>
              `
            )
            .join("")}
        </div>
      `
    ),
    renderKingdomGuideSection("Economy And XP", renderKingdomGuideList(profile?.economyAndXP || [])),
    renderKingdomGuideSection(
      "Activities",
      `
        ${renderKingdomGuideList(
          (profile?.activitiesAdded || []).map(
            (entry) => `${entry.name}${entry.source ? ` (${entry.source})` : ""}: ${entry.summary || ""}`
          )
        )}
        ${renderKingdomGuideList(profile?.activitiesChanged || [])}
      `
    ),
    renderKingdomGuideSection("Settlements", renderKingdomGuideList(profile?.settlementRules || [])),
    renderKingdomGuideSection("Construction", renderKingdomGuideList(profile?.constructionRules || [])),
    renderKingdomGuideSection("Structure Bonus Notes", renderKingdomGuideList(profile?.structureBonuses || [])),
    renderKingdomGuideSection("Clarifications", renderKingdomGuideList(profile?.clarifications || [])),
    renderKingdomGuideSection(
      "Advancement Table",
      `
        <div class="card-list kingdom-advancement-grid">
          ${(profile?.advancement || [])
            .map(
              (entry) => `
                <article class="entry">
                  <div class="entry-head">
                    <span class="entry-title">Level ${escapeHtml(String(entry?.level || "?"))}</span>
                    <span class="entry-meta">Control DC ${escapeHtml(String(entry?.controlDC || "?"))}</span>
                  </div>
                  ${renderKingdomGuideList(entry?.features || [])}
                </article>
              `
            )
            .join("")}
        </div>
      `
    ),
    renderKingdomGuideSection("Current Watchlist", renderKingdomGuideList(kingdom?.pendingProjects || [])),
    renderKingdomGuideSection("AI Prompt Ideas", renderKingdomGuideList(profile?.helpPrompts || [])),
  ];
  return `<div class="kingdom-guide-grid">${sections.filter(Boolean).join("")}</div>`;
}

function renderKingdomGuideSection(title, body) {
  const cleanBody = str(body);
  if (!cleanBody) return "";
  return `
    <details class="session-edit-panel kingdom-guide-section" open>
      <summary>${escapeHtml(title)}</summary>
      <div class="kingdom-guide-body">${body}</div>
    </details>
  `;
}

function renderKingdomGuideList(items, options = {}) {
  const entries = Array.isArray(items) ? items.map((entry) => str(entry)).filter(Boolean) : [];
  if (!entries.length) return "";
  const tag = options.ordered ? "ol" : "ul";
  const className = options.className || "flow-list";
  return `<${tag} class="${className}">${entries.map((entry) => `<li>${escapeHtml(entry)}</li>`).join("")}</${tag}>`;
}

function renderKingdomLeaderEntry(leader) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(leader.name || "Unnamed leader")}</span>
        <span class="entry-meta">${escapeHtml(leader.role || "Role")} • ${escapeHtml(leader.type || "NPC")} • +${escapeHtml(String(leader.leadershipBonus || 0))}</span>
      </div>
      <div class="row">
        <label>Role
          <input data-collection="kingdomLeaders" data-id="${leader.id}" data-field="role" value="${escapeHtml(leader.role || "")}" />
        </label>
        <label>Name
          <input data-collection="kingdomLeaders" data-id="${leader.id}" data-field="name" value="${escapeHtml(leader.name || "")}" />
        </label>
        <label>Type
          <select data-collection="kingdomLeaders" data-id="${leader.id}" data-field="type">
            ${["PC", "NPC"].map((value) => `<option value="${value}" ${leader.type === value ? "selected" : ""}>${value}</option>`).join("")}
          </select>
        </label>
        <label>Leadership Bonus
          <input data-collection="kingdomLeaders" data-id="${leader.id}" data-field="leadershipBonus" type="number" min="0" max="4" value="${escapeHtml(
            String(leader.leadershipBonus || 0)
          )}" />
        </label>
      </div>
      <label>Relevant Skills
        <input data-collection="kingdomLeaders" data-id="${leader.id}" data-field="relevantSkills" value="${escapeHtml(leader.relevantSkills || "")}" />
      </label>
      <label>Specialized Skills
        <input data-collection="kingdomLeaders" data-id="${leader.id}" data-field="specializedSkills" value="${escapeHtml(leader.specializedSkills || "")}" />
      </label>
      <label>Notes
        <textarea data-collection="kingdomLeaders" data-id="${leader.id}" data-field="notes">${escapeHtml(leader.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="kingdomLeaders" data-id="${leader.id}">Delete Leader</button>
      </div>
    </article>
  `;
}

function renderKingdomSettlementEntry(settlement) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(settlement.name || "Unnamed settlement")}</span>
        <span class="entry-meta">${escapeHtml(settlement.size || "Settlement")} • influence ${escapeHtml(String(settlement.influence || 0))}</span>
      </div>
      <div class="row">
        <label>Name
          <input data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="name" value="${escapeHtml(settlement.name || "")}" />
        </label>
        <label>Size
          <select data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="size">
            ${["Village", "Town", "City", "Metropolis"]
              .map((value) => `<option value="${value}" ${settlement.size === value ? "selected" : ""}>${value}</option>`)
              .join("")}
          </select>
        </label>
        <label>Influence
          <input data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="influence" type="number" min="0" value="${escapeHtml(
            String(settlement.influence || 0)
          )}" />
        </label>
      </div>
      <div class="row">
        <label>Civic Structure
          <select data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="civicStructure">
            ${["", "Town Hall", "Castle", "Palace"]
              .map((value) => `<option value="${value}" ${settlement.civicStructure === value ? "selected" : ""}>${value || "None"}</option>`)
              .join("")}
          </select>
        </label>
        <label>Resource Dice
          <input data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="resourceDice" type="number" min="0" value="${escapeHtml(
            String(settlement.resourceDice || 0)
          )}" />
        </label>
        <label>Consumption
          <input data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="consumption" type="number" min="0" value="${escapeHtml(
            String(settlement.consumption || 0)
          )}" />
        </label>
      </div>
      <label>Notes
        <textarea data-collection="kingdomSettlements" data-id="${settlement.id}" data-field="notes">${escapeHtml(settlement.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="kingdomSettlements" data-id="${settlement.id}">Delete Settlement</button>
      </div>
    </article>
  `;
}

function renderKingdomRegionEntry(region) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(region.hex || "Unknown hex")}</span>
        <span class="entry-meta">${escapeHtml(region.status || "Status unknown")} • ${escapeHtml(region.terrain || "terrain n/a")}</span>
      </div>
      <div class="row">
        <label>Hex
          <input data-collection="kingdomRegions" data-id="${region.id}" data-field="hex" value="${escapeHtml(region.hex || "")}" />
        </label>
        <label>Status
          <input data-collection="kingdomRegions" data-id="${region.id}" data-field="status" value="${escapeHtml(region.status || "")}" />
        </label>
        <label>Terrain
          <input data-collection="kingdomRegions" data-id="${region.id}" data-field="terrain" value="${escapeHtml(region.terrain || "")}" />
        </label>
        <label>Work Site
          <input data-collection="kingdomRegions" data-id="${region.id}" data-field="workSite" value="${escapeHtml(region.workSite || "")}" />
        </label>
      </div>
      <label>Notes
        <textarea data-collection="kingdomRegions" data-id="${region.id}" data-field="notes">${escapeHtml(region.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="kingdomRegions" data-id="${region.id}">Delete Region</button>
      </div>
    </article>
  `;
}

function renderKingdomTurnEntry(turn) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(turn.title || "Kingdom Turn")}</span>
        <span class="entry-meta">${escapeHtml(turn.date || "No date")} • RP ${turn.rpDelta >= 0 ? "+" : ""}${escapeHtml(String(turn.rpDelta || 0))} • Unrest ${
          turn.unrestDelta >= 0 ? "+" : ""
        }${escapeHtml(String(turn.unrestDelta || 0))}</span>
      </div>
      <label>Title
        <input data-collection="kingdomTurns" data-id="${turn.id}" data-field="title" value="${escapeHtml(turn.title || "")}" />
      </label>
      <div class="row">
        <label>Date
          <input data-collection="kingdomTurns" data-id="${turn.id}" data-field="date" value="${escapeHtml(turn.date || "")}" />
        </label>
        <label>Summary
          <input data-collection="kingdomTurns" data-id="${turn.id}" data-field="summary" value="${escapeHtml(turn.summary || "")}" />
        </label>
      </div>
      <label>Risks / Follow-Ups
        <textarea data-collection="kingdomTurns" data-id="${turn.id}" data-field="risks">${escapeHtml(turn.risks || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="kingdomTurns" data-id="${turn.id}">Delete Turn</button>
      </div>
    </article>
  `;
}

function renderPdfIntel() {
  const indexedFiles = Array.isArray(state?.meta?.pdfIndexedFiles)
    ? state.meta.pdfIndexedFiles.map((name) => str(name)).filter(Boolean)
    : [];
  const selectedSummaryFile =
    str(ui.pdfSummaryFile) && indexedFiles.includes(str(ui.pdfSummaryFile)) ? str(ui.pdfSummaryFile) : indexedFiles[0] || "";
  if (!ui.pdfSummaryBusy && selectedSummaryFile && !ui.pdfSummaryFile) {
    ui.pdfSummaryFile = selectedSummaryFile;
  }
  const storedSummary = getPdfSummaryByFileName(selectedSummaryFile);
  const summaryOutput = str(ui.pdfSummaryOutput) || str(storedSummary?.summary);
  const summaryStamp = str(storedSummary?.updatedAt || "");
  const summaryProgressTotal = Math.max(1, Number.parseInt(String(ui.pdfSummaryProgressTotal || "0"), 10) || 1);
  const summaryProgressCurrent = Math.max(
    0,
    Math.min(
      Number.parseInt(String(ui.pdfSummaryProgressCurrent || "0"), 10) || 0,
      summaryProgressTotal
    )
  );
  const summaryProgressPercent = Math.round((summaryProgressCurrent / summaryProgressTotal) * 100);
  const summaryProgressActive = ui.pdfSummaryBusy || !!str(ui.pdfSummaryProgressLabel);
  const summaryProgressLabel = str(ui.pdfSummaryProgressLabel) || "Working...";

  const status = `
    <p class="small">
      Folder: <span class="mono">${escapeHtml(state.meta.pdfFolder || "Not set")}</span><br />
      Last Indexed: <span class="mono">${escapeHtml(state.meta.pdfIndexedAt || "Never")}</span><br />
      Indexed Files: <span class="mono">${escapeHtml(String(state.meta.pdfIndexedCount || 0))}</span>
    </p>
  `;

  if (!desktopApi) {
    return `
      <div class="page-stack">
        ${renderPageIntro("PDF Intel", "Index and search your local TTRPG PDFs so rules/lore checks stay fast at the table.")}
        <section class="panel flow-panel">
          <h2>Run Order</h2>
          <ol class="flow-list">
            <li><strong>Step 1:</strong> open desktop app build.</li>
            <li><strong>Step 2:</strong> index your PDF folder.</li>
            <li><strong>Step 3:</strong> search rules/lore while prepping.</li>
          </ol>
        </section>
        <section class="panel step-card">
          <div class="step-head">
            <span class="step-badge">!</span>
            <h2>Desktop Required</h2>
          </div>
          <p>This feature is available only in the desktop build.</p>
          <p class="small">Run this app through Electron to index/search your local PDFs.</p>
        </section>
      </div>
    `;
  }

  return `
    <div class="page-stack">
      ${renderPageIntro("PDF Intel", "Index local PDFs once, then search quickly for rules, lore, and chapter details.")}
      <section class="panel flow-panel">
        <h2>Run Order</h2>
        <ol class="flow-list">
          <li><strong>Step 1:</strong> choose folder and index PDFs.</li>
          <li><strong>Step 2:</strong> run focused keyword searches.</li>
          <li><strong>Step 3:</strong> jump directly to matched pages and pull what you need into prep.</li>
        </ol>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">1</span>
          <h2>Index Your PDF Library</h2>
        </div>
        <label>PDF Folder
          <input id="pdf-folder-input" value="${escapeHtml(state.meta.pdfFolder || "")}" placeholder="C:\\Users\\Chris Bender\\OneDrive\\Desktop\\TTRPG-PDFs" />
        </label>
        <div class="toolbar">
          <button class="btn btn-secondary" data-action="pdf-choose-folder">Choose Folder</button>
          <button class="btn btn-primary" data-action="pdf-index" ${ui.pdfBusy ? "disabled" : ""}>
            ${ui.pdfBusy ? "Indexing..." : "Index PDFs"}
          </button>
        </div>
        ${status}
        ${ui.pdfMessage ? `<p class="small">${escapeHtml(ui.pdfMessage)}</p>` : ""}
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">2</span>
          <h2>Search Indexed PDFs</h2>
        </div>
        <form data-form="pdf-search">
          <div class="row">
            <label>Search Query
              <input name="query" value="${escapeHtml(ui.pdfSearchQuery)}" placeholder="e.g., travel hazards, downtime, undead, faction politics" />
            </label>
            <label>Max Results
              <select name="limit">
                <option value="10">10</option>
                <option value="20" selected>20</option>
                <option value="40">40</option>
              </select>
            </label>
          </div>
          <button class="btn btn-primary" type="submit">Search</button>
        </form>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">3</span>
          <h2>Use Search Results</h2>
        </div>
        <p class="small">Open the exact matched page first, then use the snippet to quickly verify context. Hybrid matches combine keyword and semantic retrieval when a local embedding model is available.</p>
        <div class="card-list" style="margin-top:12px;">
          ${
            ui.pdfSearchResults.length
              ? ui.pdfSearchResults
                  .map(
                    (r) => `
                    <article class="entry">
                      <div class="entry-head">
                        <span class="entry-title">${escapeHtml(r.fileName)}</span>
                        <span class="entry-meta">Page ${escapeHtml(String(r.page || 1))} | ${escapeHtml(
                          sentenceCaseAndPunctuation(String(r.searchMode || "lexical")).replace(/\.$/, "")
                        )} | Score: ${escapeHtml(String(r.score))}</span>
                      </div>
                      <p>${escapeHtml(r.snippet)}</p>
                      <div class="toolbar">
                        <button class="btn btn-primary" data-action="pdf-open-path-page" data-path="${encodeURIComponent(
                          r.path
                        )}" data-page="${escapeHtml(String(r.page || 1))}">Open Page</button>
                      </div>
                    </article>`
                  )
                  .join("")
              : `<p class="empty">No search results yet.</p>`
          }
        </div>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">4</span>
          <h2>Summarize Indexed PDF</h2>
        </div>
        <p class="small">Build a persistent GM-ready brief from indexed text, then reuse it across tabs. Search/RAG works after indexing even if you never run this step.</p>
        <div class="row">
          <label>Indexed File
            <select data-pdf-summary-file>
              ${
                indexedFiles.length
                  ? indexedFiles
                      .map(
                        (name) =>
                          `<option value="${escapeHtml(name)}" ${name === selectedSummaryFile ? "selected" : ""}>${escapeHtml(
                            name
                          )}</option>`
                      )
                      .join("")
                  : `<option value="">No indexed files</option>`
              }
            </select>
          </label>
        </div>
        <div class="toolbar">
          <button class="btn btn-primary" data-action="pdf-summarize-selected" ${
            ui.pdfSummaryBusy || !indexedFiles.length ? "disabled" : ""
          }>
            ${ui.pdfSummaryBusy ? "Summarizing..." : "Summarize Selected PDF"}
          </button>
          <button class="btn btn-secondary" data-action="pdf-summarize-refresh" ${
            ui.pdfSummaryBusy || !indexedFiles.length ? "disabled" : ""
          }>Refresh Summary</button>
        </div>
        ${
          summaryProgressActive
            ? `<div class="summary-progress">
                <div class="summary-progress-meta">
                  <span class="small">${escapeHtml(summaryProgressLabel)}</span>
                  <span class="small mono">${escapeHtml(
                    `${summaryProgressCurrent}/${summaryProgressTotal} (${summaryProgressPercent}%)`
                  )}</span>
                </div>
                <progress class="summary-progress-bar" max="${summaryProgressTotal}" value="${summaryProgressCurrent}"></progress>
              </div>`
            : ""
        }
        ${
          summaryStamp
            ? `<p class="small">Summary Updated: <span class="mono">${escapeHtml(summaryStamp)}</span></p>`
            : `<p class="small">No saved summary yet for this file.</p>`
        }
        <textarea readonly class="session-textarea-nextprep">${escapeHtml(
          summaryOutput || "Run Summarize Selected PDF to generate a persistent summary."
        )}</textarea>
      </section>
    </div>
  `;
}

function renderFoundry() {
  return `
    <div class="page-stack">
      ${renderPageIntro("Foundry Export", "Export clean JSON packs from your campaign tracker into Foundry VTT with one click.")}
      <section class="panel flow-panel">
        <h2>Run Order</h2>
        <ol class="flow-list">
          <li><strong>Step 1:</strong> verify what will export (counts below).</li>
          <li><strong>Step 2:</strong> export NPCs, Quests, Locations, or Full Pack.</li>
          <li><strong>Step 3:</strong> import JSON into Foundry and review journals/actors.</li>
        </ol>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">1</span>
          <h2>Verify Export Scope</h2>
        </div>
        <section class="grid grid-3">
          <article class="step-sub">
            <h3>NPC Actors</h3>
            <p class="mono">${state.npcs.length}</p>
          </article>
          <article class="step-sub">
            <h3>Quest Journals</h3>
            <p class="mono">${state.quests.length}</p>
          </article>
          <article class="step-sub">
            <h3>Location Journals</h3>
            <p class="mono">${state.locations.length}</p>
          </article>
        </section>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">2</span>
          <h2>Run Exports</h2>
        </div>
        <p class="small">NPCs export as Actors. Quests and Locations export as JournalEntries.</p>
        <div class="toolbar">
          <button class="btn btn-primary" data-action="export-foundry" data-kind="npcs">Export NPC Actors</button>
          <button class="btn btn-primary" data-action="export-foundry" data-kind="quests">Export Quest Journals</button>
          <button class="btn btn-primary" data-action="export-foundry" data-kind="locations">Export Location Journals</button>
          <button class="btn btn-secondary" data-action="export-foundry" data-kind="all">Export Full Pack</button>
        </div>
      </section>

      <section class="panel step-card">
        <div class="step-head">
          <span class="step-badge">3</span>
          <h2>Foundry Import Checklist</h2>
        </div>
        <ul class="list">
          <li>Open your Foundry world and create/import destination folders first.</li>
          <li>Import exported JSON files and verify actor/journal names and images.</li>
          <li>Spot-check one NPC and one quest journal before session day.</li>
        </ul>
      </section>
    </div>
  `;
}

function sessionEntry(s) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(s.title)}</span>
        <span class="entry-meta">${escapeHtml(s.date || "")} • ${escapeHtml(s.arc || "No arc")}</span>
      </div>
      ${renderSessionReadableView(s)}
      <details class="session-edit-panel">
        <summary>Edit Raw Fields</summary>
        <label>Summary
          <textarea class="session-textarea-summary" data-collection="sessions" data-id="${s.id}" data-field="summary">${escapeHtml(
            s.summary || ""
          )}</textarea>
        </label>
        <label>Next Prep
          <textarea class="session-textarea-nextprep" data-collection="sessions" data-id="${s.id}" data-field="nextPrep">${escapeHtml(
            s.nextPrep || ""
          )}</textarea>
        </label>
      </details>
      <div class="toolbar">
        <button class="btn btn-secondary" data-action="session-export-packet-one" data-id="${s.id}">Export Packet</button>
        <button class="btn btn-secondary" data-action="session-wrapup-one" data-id="${s.id}">Smart Wrap-Up</button>
        <button class="btn btn-secondary" data-action="session-wizard-open-one" data-id="${s.id}">Wizard</button>
        <button class="btn btn-danger" data-action="delete" data-collection="sessions" data-id="${s.id}">Delete</button>
      </div>
    </article>
  `;
}

function renderSessionReadableView(session) {
  const summary = str(session?.summary);
  const nextPrep = str(session?.nextPrep);
  const blocks = parseSessionReadableBlocks(nextPrep);
  const summaryHtml = renderReadableContent(summary || "No summary captured yet.");
  return `
    <section class="session-readable">
      <h4 class="session-readable-title">Readable View</h4>
      <article class="session-readable-block tone-summary">
        <div class="session-readable-label">Summary</div>
        <div class="session-readable-content">${summaryHtml}</div>
      </article>
      ${
        blocks.length
          ? blocks
              .map(
                (block) => `
              <details class="session-readable-block ${escapeHtml(block.tone)}" open>
                <summary>${escapeHtml(block.title)}</summary>
                <div class="session-readable-content">${renderReadableContent(block.body)}</div>
              </details>
            `
              )
              .join("")
          : `<article class="session-readable-block tone-base">
              <div class="session-readable-label">Next Prep</div>
              <div class="session-readable-content">${renderReadableContent(nextPrep || "No prep note yet.")}</div>
            </article>`
      }
    </section>
  `;
}

function renderCopilotOutputPanel(text) {
  const clean = str(text);
  if (!clean) return "";
  return `
    <section class="copilot-readable">
      <article class="session-readable-block tone-session">
        <div class="session-readable-label">Readable Output</div>
        <div class="session-readable-content copilot-readable-content">${renderReadableContent(clean)}</div>
      </article>
      <details class="session-edit-panel">
        <summary>Raw Output</summary>
        <textarea class="copilot-output" readonly>${escapeHtml(clean)}</textarea>
      </details>
    </section>
  `;
}

function parseSessionReadableBlocks(nextPrep) {
  const text = String(nextPrep || "");
  if (!text) return [];
  const definitions = [
    { key: "SMART_WRAPUP", title: "Smart Wrap-Up", tone: "tone-wrap" },
    { key: "SMART_SCENES", title: "Scene Openers", tone: "tone-scenes" },
    { key: "AI_DASHBOARD", title: "AI Prep Plan", tone: "tone-dashboard" },
    { key: "AI_SESSION", title: "AI Session Prep", tone: "tone-session" },
    { key: "AI_FOUNDRY", title: "AI Foundry Notes", tone: "tone-foundry" },
    { key: "AUTO_LINKS", title: "Auto-Connected Links", tone: "tone-links" },
  ];
  const blocks = [];
  let remainder = text;
  for (const def of definitions) {
    const section = extractMarkedSessionSection(text, def.key);
    if (!section) continue;
    blocks.push({
      title: def.title,
      tone: def.tone,
      body: section,
    });
    remainder = stripMarkedSessionSection(remainder, def.key);
  }
  const base = remainder
    .replace(/<!--\s*[A-Z0-9_]+\s*-->/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  if (base) {
    blocks.push({
      title: "Base Prep Notes",
      tone: "tone-base",
      body: base,
    });
  }
  return blocks;
}

function extractMarkedSessionSection(text, key) {
  const source = String(text || "");
  if (!source) return "";
  const start = `<!-- ${key}_START -->`;
  const end = `<!-- ${key}_END -->`;
  const regex = new RegExp(`${escapeRegex(start)}([\\s\\S]*?)${escapeRegex(end)}`, "m");
  const match = source.match(regex);
  if (!match) return "";
  return String(match[1] || "")
    .replace(/\r/g, "")
    .replace(/^\s+|\s+$/g, "");
}

function stripMarkedSessionSection(text, key) {
  const source = String(text || "");
  if (!source) return "";
  const start = `<!-- ${key}_START -->`;
  const end = `<!-- ${key}_END -->`;
  const regex = new RegExp(`${escapeRegex(start)}[\\s\\S]*?${escapeRegex(end)}`, "m");
  return source.replace(regex, "").trim();
}

function renderReadableContent(text) {
  const lines = String(text || "")
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.trim());
  if (!lines.some(Boolean)) return `<p class="small">No content yet.</p>`;

  const html = [];
  let inList = false;
  const closeList = () => {
    if (!inList) return;
    html.push("</ul>");
    inList = false;
  };

  for (let index = 0; index < lines.length; index += 1) {
    const rawLine = lines[index];
    const line = rawLine.replace(/^•\s+/, "- ");
    if (!line) {
      closeList();
      continue;
    }

    if (isMarkdownTableStart(lines, index)) {
      closeList();
      const table = renderReadableTable(lines, index);
      html.push(table.html);
      index = table.nextIndex;
      continue;
    }

    const heading = line.match(/^\*\*(.+?)\*\*$/) || line.match(/^#{1,4}\s+(.+)$/);
    if (heading) {
      closeList();
      html.push(`<h5 class="prep-heading">${formatReadableInline(heading[1])}</h5>`);
      continue;
    }

    if (/^(-|\*|\d+\.)\s+/.test(line)) {
      if (!inList) {
        html.push('<ul class="prep-list">');
        inList = true;
      }
      html.push(`<li>${formatReadableInline(line.replace(/^(-|\*|\d+\.)\s+/, ""))}</li>`);
      continue;
    }

    closeList();
    html.push(`<p>${formatReadableInline(line, true)}</p>`);
  }
  closeList();
  return html.join("");
}

function isMarkdownTableStart(lines, index) {
  const current = str(lines?.[index]);
  const divider = str(lines?.[index + 1]);
  return isMarkdownTableRow(current) && isMarkdownTableDivider(divider);
}

function isMarkdownTableRow(line) {
  const text = String(line || "").trim();
  return text.startsWith("|") && text.endsWith("|") && text.slice(1, -1).includes("|");
}

function isMarkdownTableDivider(line) {
  const cells = splitMarkdownTableRow(line);
  return cells.length > 0 && cells.every((cell) => /^:?-{3,}:?$/.test(cell));
}

function splitMarkdownTableRow(line) {
  return String(line || "")
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function renderReadableTable(lines, startIndex) {
  const headerCells = splitMarkdownTableRow(lines[startIndex]);
  const rows = [];
  let cursor = startIndex + 2;
  while (cursor < lines.length && isMarkdownTableRow(lines[cursor])) {
    rows.push(splitMarkdownTableRow(lines[cursor]));
    cursor += 1;
  }

  const html = `
    <div class="readable-table-wrap">
      <table class="readable-table">
        <thead>
          <tr>${headerCells.map((cell) => `<th>${formatReadableInline(cell, true)}</th>`).join("")}</tr>
        </thead>
        <tbody>
          ${rows
            .map((cells) => `<tr>${cells.map((cell) => `<td>${formatReadableInline(cell, true)}</td>`).join("")}</tr>`)
            .join("")}
        </tbody>
      </table>
    </div>
  `;
  return {
    html,
    nextIndex: cursor - 1,
  };
}

function formatReadableInline(text, keepBreaks = false) {
  const parts = String(text || "").split(/<br\s*\/?>/gi);
  const joined = parts
    .map((part) => {
      const escaped = escapeHtml(part.trim());
      return escaped
        .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
        .replace(/`([^`]+)`/g, '<span class="mono">$1</span>');
    })
    .join(keepBreaks ? "<br />" : "; ");
  return joined || escapeHtml(String(text || ""));
}

function ensureChecklistChecks() {
  const checks = state?.meta?.checklistChecks;
  if (!checks || typeof checks !== "object" || Array.isArray(checks)) return {};
  return checks;
}

function ensureCustomChecklistItems() {
  const items = state?.meta?.customChecklistItems;
  if (!Array.isArray(items)) return [];
  return items
    .map((item) => ({
      id: str(item?.id),
      label: normalizeChecklistLabel(item?.label),
    }))
    .filter((item) => item.id && item.label);
}

function ensureChecklistOverrides() {
  const overrides = state?.meta?.checklistOverrides;
  if (!overrides || typeof overrides !== "object" || Array.isArray(overrides)) return {};
  return overrides;
}

function normalizeChecklistLabel(value) {
  return str(value).replace(/\s+/g, " ");
}

function isCustomChecklistId(id) {
  return str(id).startsWith("custom-check-");
}

function updateChecklistLabel(id, nextValue) {
  const itemId = str(id);
  if (!itemId) return;
  const clean = normalizeChecklistLabel(nextValue);
  if (isCustomChecklistId(itemId)) {
    if (!clean) return;
    const items = ensureCustomChecklistItems().map((item) => (item.id === itemId ? { ...item, label: clean } : item));
    state.meta.customChecklistItems = items;
  } else {
    const overrides = ensureChecklistOverrides();
    if (!clean) {
      delete overrides[itemId];
    } else {
      overrides[itemId] = clean;
    }
    state.meta.checklistOverrides = overrides;
  }
  saveState();
}

function ensureChecklistArchived() {
  const archived = state?.meta?.checklistArchived;
  if (!archived || typeof archived !== "object" || Array.isArray(archived)) return {};
  const out = {};
  for (const [id, value] of Object.entries(archived)) {
    if (!id) continue;
    if (value === true) {
      out[id] = { label: "", archivedAt: "" };
      continue;
    }
    if (!value || typeof value !== "object" || Array.isArray(value)) continue;
    out[id] = {
      label: normalizeChecklistLabel(value.label),
      archivedAt: str(value.archivedAt),
    };
  }
  return out;
}

function getArchivedChecklistItems(allItems, archivedMap) {
  const map = new Map((allItems || []).map((item) => [item.id, item.label]));
  const out = [];
  for (const [id, meta] of Object.entries(archivedMap || {})) {
    const label = normalizeChecklistLabel(meta?.label) || normalizeChecklistLabel(map.get(id));
    if (!label) continue;
    out.push({
      id,
      label,
      archivedAt: str(meta?.archivedAt),
    });
  }
  out.sort((a, b) => safeDate(b.archivedAt) - safeDate(a.archivedAt) || a.label.localeCompare(b.label));
  return out;
}

function archiveCompletedChecklistItems() {
  const visible = generateSmartChecklist();
  const checks = ensureChecklistChecks();
  const archived = ensureChecklistArchived();
  let moved = 0;
  for (const item of visible) {
    if (!checks[item.id]) continue;
    archived[item.id] = {
      label: normalizeChecklistLabel(item.label),
      archivedAt: new Date().toISOString(),
    };
    delete checks[item.id];
    moved += 1;
  }
  if (!moved) {
    ui.sessionMessage = "No completed checklist items to archive.";
    render();
    return;
  }
  state.meta.checklistArchived = archived;
  state.meta.checklistChecks = checks;
  saveState();
  ui.sessionMessage = `Archived ${moved} completed checklist item(s).`;
  render();
}

async function generateChecklistWithAi() {
  if (!desktopApi?.generateLocalAiText) {
    ui.sessionMessage = "Local AI is not available in this runtime.";
    render();
    return;
  }
  if (ui.checklistAiBusy) return;

  const config = ensureAiConfig();
  const context = collectAiCampaignContext();
  const latest = getLatestSession();
  const prompt = [
    "Generate 8 concise prep checklist items for my next tabletop session.",
    "Each line should be one actionable checklist item for a GM.",
    "Keep lines short and specific (no numbering).",
    "Do not include markdown headings.",
    `Latest session title: ${str(latest?.title) || "unknown"}`,
    `Latest summary: ${str(latest?.summary) || "none"}`,
    `Latest prep notes: ${str(latest?.nextPrep) || "none"}`,
  ].join("\n");

  ui.checklistAiBusy = true;
  ui.sessionMessage = "AI generating checklist items...";
  render();

  try {
    const response = await desktopApi.generateLocalAiText({
      mode: "prep",
      input: prompt,
      context: {
        ...context,
        activeTab: "sessions",
        tabLabel: "Session Runner",
        tabContext: "Generate next-session prep checklist items.",
      },
      config,
    });

    const processed = processAiOutputWithFallback({
      rawText: response?.text || "",
      mode: "prep",
      input: prompt,
      source: "checklist",
      tabId: "sessions",
    });
    const usedFallback = processed.usedFallback || response?.usedFallback === true;
    const parsed = parseChecklistLines(processed.text || "");
    if (!parsed.length) {
      ui.sessionMessage = usedFallback
        ? "AI output looked like instruction text. DM Helper fallback generated checklist content."
        : "AI returned no checklist items.";
      return;
    }

    const existing = ensureCustomChecklistItems();
    const existingLabels = new Set(generateSmartChecklist({ includeArchived: true }).map((item) => item.label.toLowerCase()));
    const additions = [];
    for (const label of parsed) {
      const key = label.toLowerCase();
      if (existingLabels.has(key)) continue;
      existingLabels.add(key);
      additions.push({ id: `custom-check-${uid()}`, label });
    }

    if (!additions.length) {
      ui.sessionMessage = "AI checklist items were duplicates of current list.";
      return;
    }

    state.meta.customChecklistItems = [...existing, ...additions];
    saveState();
    ui.sessionMessage = usedFallback
      ? `AI output looked like instruction text. Added ${additions.length} fallback checklist item(s).`
      : `AI added ${additions.length} checklist item(s).`;
  } catch (err) {
    const message = recordAiError("Checklist generation", err);
    ui.sessionMessage = `AI checklist generation failed: ${message}`;
  } finally {
    ui.checklistAiBusy = false;
    render();
  }
}

function parseChecklistLines(text) {
  const lines = String(text || "")
    .split(/\n+/)
    .map((line) => line.replace(/^\s*(?:[-*]|\d+[.)])\s*/, "").trim())
    .map((line) => normalizeChecklistLabel(line))
    .filter(Boolean)
    .filter((line) => line.length >= 6 && line.length <= 180);

  const seen = new Set();
  const out = [];
  for (const line of lines) {
    const key = line.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(line.endsWith(".") ? line : `${line}.`);
    if (out.length >= 12) break;
  }
  return out;
}

function getLatestSession() {
  const sorted = [...state.sessions].sort((a, b) => sessionSortKey(b) - sessionSortKey(a));
  return sorted[0] || null;
}

function getPrepQueueMode() {
  const mode = Number.parseInt(String(state?.meta?.prepQueueMode || "60"), 10);
  if (mode === 30 || mode === 90) return mode;
  return 60;
}

function setPrepQueueMode(mode) {
  const normalized = mode === 30 || mode === 90 ? mode : 60;
  state.meta.prepQueueMode = normalized;
  saveState();
  render();
}

function ensurePrepQueueChecks() {
  const checks = state?.meta?.prepQueueChecks;
  if (!checks || typeof checks !== "object" || Array.isArray(checks)) return {};
  return checks;
}

function ensureAiConfig() {
  const base = {
    endpoint: "http://127.0.0.1:11434",
    model: "llama3.1:8b",
    temperature: 0.2,
    maxOutputTokens: 320,
    timeoutSec: 120,
    compactContext: true,
    autoRunTabs: true,
    usePdfContext: true,
    aiProfile: "fast",
  };
  const current =
    state?.meta?.aiConfig && typeof state.meta.aiConfig === "object" && !Array.isArray(state.meta.aiConfig)
      ? state.meta.aiConfig
      : {};
  const temperatureRaw = Number.parseFloat(String(current.temperature ?? base.temperature));
  const maxOutputTokensRaw = Number.parseInt(String(current.maxOutputTokens ?? base.maxOutputTokens), 10);
  const timeoutSecRaw = Number.parseInt(String(current.timeoutSec ?? base.timeoutSec), 10);
  const merged = {
    endpoint: str(current.endpoint) || base.endpoint,
    model: str(current.model) || base.model,
    temperature: Number.isFinite(temperatureRaw) ? Math.max(0, Math.min(temperatureRaw, 2)) : base.temperature,
    maxOutputTokens: Number.isFinite(maxOutputTokensRaw) ? Math.max(64, Math.min(maxOutputTokensRaw, 2048)) : base.maxOutputTokens,
    timeoutSec: Number.isFinite(timeoutSecRaw) ? Math.max(15, Math.min(timeoutSecRaw, 1200)) : base.timeoutSec,
    compactContext: current.compactContext === false ? false : true,
    autoRunTabs: current.autoRunTabs === false ? false : true,
    usePdfContext: current.usePdfContext === false ? false : true,
    aiProfile: ["fast", "deep", "custom"].includes(str(current.aiProfile).toLowerCase())
      ? str(current.aiProfile).toLowerCase()
      : base.aiProfile,
  };
  state.meta.aiConfig = merged;
  return merged;
}

function ensureAiHistory() {
  const raw = Array.isArray(state?.meta?.aiHistory) ? state.meta.aiHistory : [];
  const cleaned = [];
  for (const entry of raw) {
    const role = str(entry?.role).toLowerCase();
    const text = normalizeAiHistoryText(entry?.text);
    if (!text) continue;
    if (role !== "user" && role !== "assistant") continue;
    if (role === "assistant" && isLikelyInstructionEcho(text)) continue;
    const tabId = str(entry?.tabId) || "dashboard";
    const mode = str(entry?.mode) || "assistant";
    cleaned.push({
      id: str(entry?.id) || `ai-turn-${uid()}`,
      tabId,
      role,
      mode,
      text: text.slice(0, 1800),
      at: str(entry?.at) || new Date().toISOString(),
    });
  }
  state.meta.aiHistory = cleaned.slice(-AI_HISTORY_LIMIT);
  return state.meta.aiHistory;
}

function addAiHistoryTurn({ tabId, role, mode, text }) {
  const message = normalizeAiHistoryText(text);
  if (!message) return;
  const normalizedRole = str(role).toLowerCase();
  if (normalizedRole !== "user" && normalizedRole !== "assistant") return;
  const history = ensureAiHistory();
  history.push({
    id: `ai-turn-${uid()}`,
    tabId: str(tabId) || activeTab || "dashboard",
    role: normalizedRole,
    mode: str(mode) || "assistant",
    text: message.slice(0, 1800),
    at: new Date().toISOString(),
  });
  state.meta.aiHistory = history.slice(-AI_HISTORY_LIMIT);
  saveState();
}

function normalizeAiHistoryText(text) {
  const source = str(text).replace(/\r\n?/g, "\n");
  if (!source) return "";
  const lines = source.split("\n").map((line) => line.replace(/[ \t]+/g, " ").trimEnd());
  const cleaned = [];
  let lastBlank = false;
  for (const line of lines) {
    const normalized = line.trim() ? line : "";
    if (!normalized) {
      if (lastBlank) continue;
      lastBlank = true;
      cleaned.push("");
      continue;
    }
    lastBlank = false;
    cleaned.push(normalized);
  }
  return cleaned.join("\n").trim().slice(0, 1800);
}

function getRecentAiHistory(tabId, limit = 10) {
  const history = ensureAiHistory();
  const max = Number.parseInt(String(limit || "10"), 10);
  const target = Number.isFinite(max) ? Math.max(1, Math.min(max, 40)) : 10;
  if (!history.length) return [];

  const reversed = [...history].reverse();
  const picked = [];
  for (const entry of reversed) {
    if (entry.tabId !== tabId) continue;
    picked.push(entry);
    if (picked.length >= target) break;
  }
  if (picked.length < target) {
    for (const entry of reversed) {
      if (entry.tabId === tabId) continue;
      picked.push(entry);
      if (picked.length >= target) break;
    }
  }
  return picked.reverse();
}

function getAiHistoryEntryById(entryId) {
  const target = str(entryId);
  if (!target) return null;
  return ensureAiHistory().find((entry) => entry.id === target) || null;
}

function buildAiModelOptions(currentModel, models) {
  const normalizedCurrent = str(currentModel);
  const unique = [];
  const seen = new Set();
  for (const model of models || []) {
    const clean = str(model);
    if (!clean) continue;
    const key = clean.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(clean);
  }
  if (normalizedCurrent && !seen.has(normalizedCurrent.toLowerCase())) {
    unique.unshift(normalizedCurrent);
  }
  if (!unique.length) {
    const fallbackModel = normalizedCurrent || "llama3.1:8b";
    const fallbackLabel = normalizedCurrent ? getAiModelDisplayName(fallbackModel) : "No models loaded";
    return `<option value="${escapeHtml(fallbackModel)}">${escapeHtml(fallbackLabel)}</option>`;
  }
  return unique
    .map(
      (model) =>
        `<option value="${escapeHtml(model)}" ${model.toLowerCase() === normalizedCurrent.toLowerCase() ? "selected" : ""}>${escapeHtml(
          getAiModelDisplayName(model)
        )}</option>`
    )
    .join("");
}

function getAiModelDisplayName(model) {
  const raw = str(model).trim();
  if (!raw) return "";
  return AI_MODEL_LABELS[raw.toLowerCase()] || prettifyAiModelId(raw);
}

function prettifyAiModelId(model) {
  const raw = str(model).trim();
  if (!raw) return "";
  const [base, tag = ""] = raw.split(":");
  const prettyBase = base
    .replace(/[-_]+/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .map((token) => {
      if (/^pf2e$/i.test(token)) return "PF2e";
      if (/^gpt$/i.test(token)) return "GPT";
      if (/^oss$/i.test(token)) return "OSS";
      if (/^cpu$/i.test(token)) return "CPU";
      if (/^qwen/i.test(token)) return token.replace(/^qwen/i, "Qwen");
      if (/^llama/i.test(token)) return token.replace(/^llama/i, "Llama");
      if (/^\d+(\.\d+)?b$/i.test(token)) return token.toUpperCase();
      return token.charAt(0).toUpperCase() + token.slice(1).toLowerCase();
    })
    .join(" ");

  if (!tag || tag.toLowerCase() === "latest") return prettyBase;
  if (/^\d+(\.\d+)?b$/i.test(tag)) return `${prettyBase} (${tag.toUpperCase()})`;
  return `${prettyBase} (${tag.replace(/[-_]+/g, " ").toUpperCase()})`;
}

function replaceAiModelLabelsInText(text) {
  let output = str(text);
  const entries = Object.entries(AI_MODEL_LABELS).sort((a, b) => b[0].length - a[0].length);
  for (const [rawModel, label] of entries) {
    output = output.replace(new RegExp(escapeRegex(rawModel), "gi"), label);
  }
  return output;
}

function renderAiSelectedModelHelp(model) {
  const raw = str(model).trim();
  if (!raw) return "";
  const friendly = getAiModelDisplayName(raw);
  const suffix = friendly !== raw ? ` <span class="mono">(${escapeHtml(raw)})</span>` : "";
  return `<p class="small">Selected model: <strong>${escapeHtml(friendly)}</strong>${suffix}</p>`;
}

function pickInstalledModelByPreference(preferences, fallbackModel = "") {
  const installed = Array.isArray(ui.aiModels) ? ui.aiModels.map((model) => str(model)).filter(Boolean) : [];
  if (!installed.length) return str(fallbackModel);

  const lowered = installed.map((model) => ({ raw: model, key: model.toLowerCase() }));
  for (const preference of preferences) {
    const pref = str(preference).toLowerCase();
    if (!pref) continue;
    const exact = lowered.find((entry) => entry.key === pref);
    if (exact) return exact.raw;
    const partial = lowered.find((entry) => entry.key.includes(pref));
    if (partial) return partial.raw;
  }
  return str(fallbackModel) || installed[0];
}

function renderAiProfileControls(aiConfig) {
  const profile = str(aiConfig?.aiProfile || "custom").toLowerCase();
  const fastActive = profile === "fast";
  const deepActive = profile === "deep";
  const profileLabel = fastActive ? "Fast Mode" : deepActive ? "Deep Mode" : "Custom";
  return `
    <div class="toolbar" style="margin-top:8px;">
      <button class="btn ${fastActive ? "btn-primary" : "btn-secondary"}" data-action="ai-profile-fast">Fast Mode</button>
      <button class="btn ${deepActive ? "btn-primary" : "btn-secondary"}" data-action="ai-profile-deep">Deep Mode</button>
    </div>
    <p class="small">Active profile: ${escapeHtml(profileLabel)}. Fast favors speed and shorter replies. Deep favors richer prep and longer context.</p>
  `;
}

function applyAiProfile(profile) {
  const normalized = str(profile).toLowerCase();
  if (normalized !== "fast" && normalized !== "deep") return;

  const config = ensureAiConfig();
  const next = { ...config };

  if (normalized === "fast") {
    next.model = pickInstalledModelByPreference(
      [
        "lorebound-pf2e-fast:latest",
        "lorebound-pf2e-minimal:latest",
        "gpt-oss-20b-fast:latest",
        "lorebound-pf2e-ultra-fast:latest",
        "llama3.1:8b",
      ],
      next.model
    );
    next.temperature = 0.2;
    next.maxOutputTokens = 260;
    next.timeoutSec = 180;
    next.compactContext = true;
    next.autoRunTabs = false;
    next.usePdfContext = true;
    next.aiProfile = "fast";
  } else {
    next.model = pickInstalledModelByPreference(
      [
        "lorebound-pf2e:latest",
        "lorebound-pf2e-qwen:latest",
        "gpt-oss:20b",
        "gpt-oss-20b-optimized:latest",
        "gpt-oss-20b-fast:latest",
      ],
      next.model
    );
    next.temperature = 0.2;
    next.maxOutputTokens = 700;
    next.timeoutSec = 420;
    next.compactContext = false;
    next.autoRunTabs = false;
    next.usePdfContext = true;
    next.aiProfile = "deep";
  }

  state.meta.aiConfig = next;
  saveState();
  ui.copilotMessage = `${normalized === "fast" ? "Fast" : "Deep"} Mode applied using model "${next.model}".`;
  ui.aiMessage = ui.copilotMessage;
  render();
}

async function refreshLocalAiModels(silent = false) {
  if (!desktopApi?.listLocalAiModels) {
    if (!silent) {
      ui.copilotMessage = "Model list is unavailable in this runtime.";
      render();
    }
    return;
  }

  const config = ensureAiConfig();
  ui.aiBusy = true;
  if (!silent) {
    ui.copilotMessage = "Loading local AI models...";
    render();
  }
  try {
    const result = await desktopApi.listLocalAiModels(config);
    ui.aiModels = Array.isArray(result?.models) ? result.models : [];
    clearAiError();
    if (!silent) {
      ui.copilotMessage = `Loaded ${ui.aiModels.length} local model(s).`;
    }
  } catch (err) {
    const message = !silent ? recordAiError("Model refresh", err) : readableError(err);
    if (!silent) {
      ui.copilotMessage = `Model refresh failed: ${message}`;
    }
  } finally {
    ui.aiBusy = false;
    render();
  }
}

function getTabLabel(tabId) {
  return tabs.find((tab) => tab.id === tabId)?.label || "Unknown";
}

function getGlobalCopilotPlaceholder(tabId) {
  if (tabId === "sessions") return "Ask for recap + next prep beats for the latest session.";
  if (tabId === "capture") return "Ask to transform live capture into clean session notes.";
  if (tabId === "kingdom") return "Ask for kingdom-turn help, action order, leader assignments, or settlement advice.";
  if (tabId === "npcs") return "Describe an NPC concept and ask for table-ready details.";
  if (tabId === "quests") return "Describe a quest idea and ask for objective/stakes.";
  if (tabId === "locations") return "Describe a hex/location and ask for a usable scene brief.";
  if (tabId === "pdf") return "Ask for best PDF search queries for your next session.";
  if (tabId === "foundry") return "Ask for Foundry handoff checklist for this session.";
  if (tabId === "writing") return "Ask for rewrite help, stronger wording, and clean structure.";
  return "Ask a GM question or chat naturally (example: hello, help me prep tonight).";
}

function getGlobalCopilotApplyLabel(tabId) {
  if (tabId === "sessions") return "Apply to Latest Session";
  if (tabId === "capture") return "Add as Live Capture";
  if (tabId === "kingdom") return "Append Kingdom Notes";
  if (tabId === "npcs") return "Create NPC(s)";
  if (tabId === "quests") return "Create Quest";
  if (tabId === "locations") return "Create Location";
  if (tabId === "pdf") return "Use as PDF Query";
  if (tabId === "foundry") return "Attach Foundry Notes";
  if (tabId === "writing") return "Send to Writing Helper";
  return "Attach to Latest Prep";
}

function getGlobalCopilotMode(tabId) {
  if (tabId === "npcs") return "npc";
  if (tabId === "quests") return "quest";
  if (tabId === "locations") return "location";
  if (tabId === "sessions" || tabId === "capture" || tabId === "writing" || tabId === "kingdom") return "session";
  return "prep";
}

function isStructuredWorldTab(tabId) {
  return tabId === "npcs" || tabId === "quests" || tabId === "locations";
}

function getSeedTabForMode(mode) {
  if (mode === "npc") return "npcs";
  if (mode === "quest") return "quests";
  if (mode === "location") return "locations";
  return "dashboard";
}

function inferStructuredModeFromInput(inputText) {
  const lower = str(inputText).toLowerCase();
  if (!lower) return "";
  const createVerb = /\b(create|make|build|draft|invent|design|write|come up with|describe)\b/;
  if (createVerb.test(lower) && /\bnpc\b/.test(lower)) return "npc";
  if (createVerb.test(lower) && /\bquest\b/.test(lower)) return "quest";
  if (createVerb.test(lower) && /\b(location|village|town|city|settlement|hex|place)\b/.test(lower)) return "location";
  return "";
}

function compactCopilotRequestText(inputText, max = 420) {
  const clean = str(inputText).replace(/\s+/g, " ");
  const limit = Number.isFinite(Number(max)) ? Math.max(160, Number(max)) : 420;
  if (clean.length <= limit) return clean;
  let cut = clean.slice(0, limit);
  const breakpoints = [". ", "? ", "! ", "; ", ", "];
  let pivot = -1;
  for (const token of breakpoints) {
    const next = cut.lastIndexOf(token);
    if (next > pivot) pivot = next + token.trimEnd().length;
  }
  if (pivot > Math.floor(limit * 0.55)) {
    cut = cut.slice(0, pivot);
  }
  return `${cut.trim()}...`;
}

function isPdfGroundedQuestion(inputText) {
  const lower = str(inputText).toLowerCase();
  if (!lower) return false;
  return /\b(selected pdf|this book|the book|book|pdf|adventure|module|chapter|section|main threat|run chapter|run it|run this)\b/.test(
    lower
  );
}

function isPdfGroundedWorldRequest(tabId, inputText) {
  return isStructuredWorldTab(tabId) && isPdfGroundedQuestion(inputText);
}

function shouldUseCopilotChatMode(tabId, inputText) {
  const text = str(inputText);
  if (!text) return false;
  const lower = text.toLowerCase();
  if (
    tabId === "sessions" &&
    /\b(idea|ideas|hook|hooks|scene|scenes|encounter|encounters|village|town|quest|npc|prep|session|run|opening)\b/.test(
      lower
    )
  ) {
    return false;
  }
  if (isPdfGroundedQuestion(lower)) {
    return false;
  }
  if (/^(hi|hello|hey|yo)\b/.test(lower)) return true;
  if (/\?$/.test(text)) return true;
  if (/\b(how are you|can you|could you|would you|help me|what|why|explain|brainstorm)\b/.test(lower)) return true;
  if (
    tabId === "dashboard" &&
    text.length <= 180 &&
    !/\b(top priorities|prep queue|opening scene|time-boxed|action plan)\b/i.test(text)
  ) {
    return true;
  }
  return false;
}

function buildGlobalCopilotRequest(tabId, userInput, autoRun) {
  const seedPrompt = buildGlobalCopilotSeedPrompt(tabId);
  const baseMode = getGlobalCopilotMode(tabId);
  const cleanInput = str(userInput);
  if (autoRun) {
    return {
      mode: baseMode,
      input: seedPrompt,
      isChat: false,
    };
  }
  if (!cleanInput) {
    return {
      mode: baseMode,
      input: seedPrompt,
      isChat: false,
    };
  }
  const inferredMode = inferStructuredModeFromInput(cleanInput);
  if (isPdfGroundedWorldRequest(tabId, cleanInput) && !isCopilotSmallTalkInput(cleanInput)) {
    return {
      mode: baseMode,
      input: buildPdfGroundedWorldPrompt(tabId, cleanInput),
      isChat: false,
    };
  }
  if (inferredMode && !isCopilotSmallTalkInput(cleanInput)) {
    const compactInput = compactCopilotRequestText(cleanInput, inferredMode === "npc" ? 420 : 340);
    return {
      mode: inferredMode,
      input: `${buildGlobalCopilotSeedPrompt(getSeedTabForMode(inferredMode))}\n${getStructuredWorldDetailInstruction(
        getSeedTabForMode(inferredMode)
      )}\n\nAdditional request:\n${compactInput}`,
      isChat: false,
    };
  }
  if (isStructuredWorldTab(tabId) && !isCopilotSmallTalkInput(cleanInput)) {
    const compactInput = compactCopilotRequestText(cleanInput, tabId === "npcs" ? 420 : 340);
    return {
      mode: baseMode,
      input: `${seedPrompt}\n${getStructuredWorldDetailInstruction(tabId)}\n\nAdditional request:\n${compactInput}`,
      isChat: false,
    };
  }
  if ((tabId === "pdf" || isPdfGroundedQuestion(cleanInput)) && !isCopilotSmallTalkInput(cleanInput)) {
    const compactInput = compactCopilotRequestText(cleanInput, 420);
    return {
      mode: "prep",
      input: [
        "Use the selected PDF summary and indexed PDF snippets to answer the GM's question.",
        "If the available PDF context is thin or missing, say that clearly instead of guessing.",
        "Return:",
        "Book Takeaways:",
        "- 3 to 6 bullets grounded in the PDF context",
        "How To Run It:",
        "- 4 to 8 GM-facing bullets",
        "Next PDF Queries:",
        "- bullet",
        "- bullet",
        "",
        "GM question:",
        compactInput,
      ].join("\n"),
      isChat: false,
    };
  }
  if (shouldUseCopilotChatMode(tabId, cleanInput)) {
    return {
      mode: "assistant",
      input: cleanInput,
      isChat: true,
    };
  }
  if (tabId === "dashboard") {
    return {
      mode: baseMode,
      input: cleanInput,
      isChat: false,
    };
  }
  return {
    mode: baseMode,
    input: `${seedPrompt}\n\nAdditional request:\n${cleanInput}`,
    isChat: false,
  };
}

function isCopilotSmallTalkInput(inputText) {
  const text = str(inputText);
  if (!text) return false;
  const lower = text.toLowerCase();
  const words = lower.split(/\s+/).filter(Boolean);
  if (words.length > 8) return false;
  if (/^(hi|hello|hey|yo|sup|howdy)\b/.test(lower)) return true;
  if (/^(how are you|hows it going|what are you|who are you|thanks|thank you)\b/.test(lower)) return true;
  if (/^(good morning|good afternoon|good evening)\b/.test(lower)) return true;
  return false;
}

function getStructuredWorldDetailInstruction(tabId) {
  if (tabId === "npcs") {
    return "Keep every field concrete, story-tied, and table-ready. Under Notes include 6 to 8 short bullets covering core want, leverage, pressure, voice, first impression, hidden truth or complication, and how to use the NPC at the table.";
  }
  if (tabId === "quests") {
    return "Keep every field concrete, story-tied, and table-ready. Make the stakes specific and the next actionable beat obvious.";
  }
  if (tabId === "locations") {
    return "Keep every field concrete, story-tied, and table-ready. Make the change, immediate tension, and next clue clearly usable at the table.";
  }
  return "Keep every field concrete, story-tied, and table-ready.";
}

function buildPdfGroundedWorldPrompt(tabId, userInput) {
  const compactInput = compactCopilotRequestText(userInput, tabId === "npcs" ? 520 : 420);
  if (tabId === "npcs") {
    return [
      "Use the selected PDF summary and indexed PDF snippets as the source of truth.",
      "Base the answer on named or clearly implied people, factions, allies, rivals, and pressures from that book.",
      "If the GM asks for a few or multiple NPCs, return 2 to 4 NPCs separated by a line containing only ---.",
      "If the book context clearly supports fewer than requested, return only the strongest confirmed NPCs and mark inferred roles as inferred instead of inventing unsupported lore.",
      "For each NPC, return exactly this structure:",
      "Name:",
      "Role:",
      "Agenda:",
      "Disposition:",
      "Notes:",
      "- Book anchor:",
      "- Why the party works with or against them:",
      "- Current pressure or fear:",
      "- Voice and mannerisms:",
      "- First impression or look:",
      "- Hidden truth or complication:",
      "- Best way to use them in the next session:",
      "",
      "GM request:",
      compactInput,
    ].join("\n");
  }
  return [
    "Use the selected PDF summary and indexed PDF snippets as the source of truth.",
    "Keep every detail grounded in that book and say when a detail is not confirmed instead of guessing.",
    buildGlobalCopilotSeedPrompt(tabId),
    getStructuredWorldDetailInstruction(tabId),
    "",
    "GM request:",
    compactInput,
  ].join("\n");
}

function buildGlobalCopilotSeedPrompt(tabId) {
  if (tabId === "dashboard") {
    return [
      "Create a high-signal GM action plan for the next session.",
      "Return:",
      "1) Top Priorities (max 5 bullets)",
      "2) 60-Minute Prep Queue (time-boxed bullets)",
      "3) Opening Scene Suggestion (2-4 sentences)",
    ].join("\n");
  }
  if (tabId === "sessions") {
    return [
      "Based on latest campaign context, produce session notes in this exact structure:",
      "Summary: (4-6 sentences with concrete outcomes)",
      "Next Prep:",
      "- 6 to 10 actionable bullets",
      "Scene Openers:",
      "- 3 opening beats with sensory detail + immediate tension",
    ].join("\n");
  }
  if (tabId === "capture") {
    return [
      "Convert recent live capture entries into clean GM notes.",
      "Return:",
      "Summary:",
      "Follow-up Tasks:",
      "- bullet",
      "- bullet",
    ].join("\n");
  }
  if (tabId === "kingdom") {
    return [
      "Using the active V&K kingdom rules profile and current kingdom state, help the GM run the next kingdom turn.",
      "Return:",
      "Kingdom Turn Focus:",
      "- 3 to 6 bullets",
      "Recommended Action Order:",
      "1. one concrete action",
      "2. one concrete action",
      "3. one concrete action",
      "Risks To Watch:",
      "- bullet",
      "- bullet",
      "What To Record In DM Helper:",
      "- bullet",
      "- bullet",
    ].join("\n");
  }
  if (tabId === "npcs") {
    return [
      "Create one table-ready NPC and return fields exactly in this structure:",
      "Name:",
      "Role:",
      "Agenda:",
      "Disposition:",
      "Notes:",
      "- Core want:",
      "- Leverage over the party or locals:",
      "- Current pressure or fear:",
      "- Voice and mannerisms:",
      "- First impression or look:",
      "- Hidden truth or complication:",
      "- Best way to use them in the next session:",
    ].join("\n");
  }
  if (tabId === "quests") {
    return [
      "Create one quest and return fields exactly:",
      "Title:",
      "Status:",
      "Objective:",
      "Giver:",
      "Stakes:",
    ].join("\n");
  }
  if (tabId === "locations") {
    return [
      "Create one location update and return fields exactly:",
      "Name:",
      "Hex:",
      "What Changed:",
      "Notes:",
    ].join("\n");
  }
  if (tabId === "pdf") {
    return [
      "Suggest high-value PDF search targets for next prep.",
      "Return:",
      "Query:",
      "Backup Queries:",
      "- bullet",
      "- bullet",
      "Why:",
    ].join("\n");
  }
  if (tabId === "foundry") {
    return [
      "Create a Foundry handoff checklist for next session.",
      "Return only concise bullet points.",
    ].join("\n");
  }
  return "Rewrite and structure this into practical GM prep notes.";
}

function buildGlobalCopilotContext(tabId) {
  const context = collectAiCampaignContext();
  const latest = getLatestSession();
  const indexedFiles = Array.isArray(state?.meta?.pdfIndexedFiles)
    ? state.meta.pdfIndexedFiles.map((name) => str(name)).filter(Boolean)
    : [];
  const selectedPdfFile =
    str(ui.pdfSummaryFile) && indexedFiles.includes(str(ui.pdfSummaryFile)) ? str(ui.pdfSummaryFile) : indexedFiles[0] || "";
  const summaryEntry = getPdfSummaryByFileName(selectedPdfFile);
  const selectedPdfSummary = str(summaryEntry?.summary).replace(/\s+/g, " ").slice(0, 900);
  const aiHistory = getRecentAiHistory(tabId, 10)
    .filter((turn) => turn.role === "user" || !isLikelyInstructionEcho(turn.text))
    .map((turn) => ({
      role: turn.role,
      text: turn.text,
      tabId: turn.tabId,
      at: turn.at,
    }));
  const recentCapture = [...(state.liveCapture || [])]
    .sort((a, b) => safeDate(b.timestamp) - safeDate(a.timestamp))
    .slice(0, 14)
    .map((entry) => `${entry.kind || "Note"}: ${entry.note || ""}`);

  let tabContext = "";
  if (tabId === "dashboard") {
    const activeQuestTitles = state.quests
      .filter((q) => q.status !== "completed" && q.status !== "failed")
      .slice(0, 6)
      .map((q) => q.title);
    tabContext = `Active quests: ${activeQuestTitles.join("; ") || "None"}.`;
  } else if (tabId === "sessions") {
    tabContext = `Latest session summary: ${str(latest?.summary)} | Next prep: ${str(latest?.nextPrep)}`;
  } else if (tabId === "capture") {
    tabContext = `Recent live capture: ${recentCapture.join(" | ") || "No capture entries yet."}`;
  } else if (tabId === "kingdom") {
    const kingdom = getKingdomState();
    const profile = getActiveKingdomProfile();
    tabContext = `Kingdom: ${kingdom.name || "Unnamed kingdom"} | Turn: ${kingdom.currentTurnLabel || "Not set"} | Level ${kingdom.level} | Size ${kingdom.size} | Control DC ${
      kingdom.controlDC
    } | Unrest ${kingdom.unrest} | Renown ${kingdom.renown} | Fame ${kingdom.fame} | Infamy ${kingdom.infamy} | Settlements ${
      kingdom.settlements.length
    } | Claimed regions ${kingdom.regions.length} | Active profile: ${profile?.shortLabel || profile?.label || "Unknown"}`;
  } else if (tabId === "npcs") {
    tabContext = `Current NPC names: ${state.npcs.slice(0, 15).map((n) => n.name).join(", ") || "None"}${
      selectedPdfFile ? ` | Selected PDF: ${selectedPdfFile} | Selected PDF summary: ${selectedPdfSummary || "No summary yet."}` : ""
    }`;
  } else if (tabId === "quests") {
    tabContext = `Current quests: ${state.quests.slice(0, 15).map((q) => `${q.title} (${q.status})`).join("; ") || "None"}${
      selectedPdfFile ? ` | Selected PDF: ${selectedPdfFile} | Selected PDF summary: ${selectedPdfSummary || "No summary yet."}` : ""
    }`;
  } else if (tabId === "locations") {
    tabContext = `Known locations: ${state.locations.slice(0, 15).map((l) => l.name).join(", ") || "None"}${
      selectedPdfFile ? ` | Selected PDF: ${selectedPdfFile} | Selected PDF summary: ${selectedPdfSummary || "No summary yet."}` : ""
    }`;
  } else if (tabId === "pdf") {
    const snippets = ui.pdfSearchResults.slice(0, 4).map((r) => `${r.fileName}: ${r.snippet}`);
    const summaryText = selectedPdfSummary.slice(0, 420);
    tabContext = `Selected PDF: ${selectedPdfFile || "(none)"} | PDF query: ${ui.pdfSearchQuery || "(none)"} | Snippets: ${
      snippets.join(" || ") || "No snippets yet."
    } | Selected PDF summary: ${summaryText || "No summary yet."}`;
  } else if (tabId === "foundry") {
    tabContext = `Foundry export counts: NPCs ${state.npcs.length}, Quests ${state.quests.length}, Locations ${state.locations.length}.`;
  } else if (tabId === "writing") {
    tabContext = `Writing Helper mode: ${ui.writingDraft.mode}. Current output length: ${str(ui.writingDraft.output).length}.`;
  }

  return {
    ...context,
    activeTab: tabId,
    tabLabel: getTabLabel(tabId),
    tabContext,
    selectedPdfFile,
    selectedPdfSummary,
    aiHistory,
  };
}

function buildMinimalCopilotContext(tabId) {
  return {
    latestSession: null,
    openQuests: [],
    npcs: [],
    locations: [],
    aiHistory: [],
    activeTab: tabId,
    tabLabel: getTabLabel(tabId),
    tabContext: "Small-talk chat. Reply briefly and naturally.",
  };
}

function pickCopilotRecoveryModel(currentModel) {
  const current = str(currentModel).toLowerCase();
  const preferred = [
    "lorebound-pf2e-fast:latest",
    "gpt-oss-20b-fast:latest",
    "lorebound-pf2e-ultra-fast:latest",
    "lorebound-pf2e:latest",
    "llama3.1:8b",
  ];
  const installed = Array.isArray(ui.aiModels) ? ui.aiModels.map((model) => str(model)).filter(Boolean) : [];
  if (installed.length) {
    for (const model of preferred) {
      if (model.toLowerCase() === current) continue;
      if (installed.some((item) => item.toLowerCase() === model.toLowerCase())) return model;
    }
    const fallbackInstalled = installed.find((item) => item.toLowerCase() !== current);
    return fallbackInstalled || "";
  }
  return preferred.find((model) => model.toLowerCase() !== current) || "";
}

async function runCopilotAiAttempt({ mode, input, config, tabId, userInput, contextOverride = null }) {
  const result = await desktopApi.generateLocalAiText({
    mode,
    input,
    config,
    context: contextOverride || buildGlobalCopilotContext(tabId),
  });
  const processed = processAiOutputWithFallback({
    rawText: result?.text || "",
    mode,
    input: userInput || input,
    source: "copilot",
    tabId,
  });
  const usedFallback = processed.usedFallback || result?.usedFallback === true;
  const fallbackReason = str(result?.fallbackReason || "");
  return {
    result,
    processed,
    usedFallback,
    fallbackReason,
  };
}

async function maybeAutoRunCopilotOnTabChange(trigger = "tab-switch") {
  const config = ensureAiConfig();
  if (!config.autoRunTabs) return;
  if (!desktopApi?.generateLocalAiText) return;
  if (ui.copilotBusy || ui.aiBusy) return;
  await runGlobalAiCopilot({ autoRun: true, trigger });
}

async function runGlobalAiCopilot(options = {}) {
  if (!desktopApi?.generateLocalAiText) {
    ui.copilotMessage = "Desktop local AI bridge is not available in this runtime.";
    render();
    return;
  }

  const autoRun = options?.autoRun === true;
  const config = ensureAiConfig();
  const userInput = str(ui.copilotDraft.input);
  const request = buildGlobalCopilotRequest(activeTab, userInput, autoRun);
  const mode = request.mode;
  const input = request.input;
  const effectiveConfig = { ...config };
  if (!request.isChat && isStructuredWorldTab(activeTab)) {
    const worldOutputFloor = activeTab === "npcs" ? 420 : 280;
    const worldOutputCeiling = activeTab === "npcs" ? 640 : 480;
    effectiveConfig.maxOutputTokens = Math.min(
      Math.max(Number(effectiveConfig.maxOutputTokens || 0) || 0, worldOutputFloor),
      worldOutputCeiling
    );
    effectiveConfig.timeoutSec = Math.max(Number(effectiveConfig.timeoutSec || 0), 300);
    effectiveConfig.compactContext = true;
  }
  if (!request.isChat && (activeTab === "pdf" || isPdfGroundedQuestion(userInput || input))) {
    effectiveConfig.maxOutputTokens = Math.max(Number(effectiveConfig.maxOutputTokens || 0) || 0, 900);
    effectiveConfig.timeoutSec = Math.max(Number(effectiveConfig.timeoutSec || 0), 420);
    effectiveConfig.temperature = Math.min(Number(effectiveConfig.temperature || 0.2) || 0.2, 0.15);
    effectiveConfig.compactContext = true;
  }
  if (!request.isChat && activeTab === "sessions") {
    effectiveConfig.maxOutputTokens = Math.max(Number(effectiveConfig.maxOutputTokens || 0) || 0, 700);
  }
  if (/20b/i.test(str(effectiveConfig.model)) && Number(effectiveConfig.timeoutSec || 0) < 300) {
    effectiveConfig.timeoutSec = 300;
  }
  effectiveConfig.timeoutMs = Math.max(15000, Number(effectiveConfig.timeoutSec || 0) * 1000);
  if (!autoRun && userInput) {
    addAiHistoryTurn({
      tabId: activeTab,
      role: "user",
      mode,
      text: userInput,
    });
  }

  const isSmallTalk = request.isChat && isCopilotSmallTalkInput(userInput);
  const contextOverride = isSmallTalk ? buildMinimalCopilotContext(activeTab) : null;
  const requestId = (Number(ui.copilotRequestSeq) || 0) + 1;
  ui.copilotRequestSeq = requestId;
  ui.copilotActiveRequestId = requestId;
  const replacingInFlight = ui.copilotBusy;

  ui.copilotBusy = true;
  ui.copilotDraft.output = "";
  ui.copilotPendingFallbackMemory = null;
  ui.copilotMessage = replacingInFlight && !autoRun
    ? "Previous request replaced. Generating new reply..."
    : autoRun
      ? `Auto-running for ${getTabLabel(activeTab)}...`
      : `Generating for ${getTabLabel(activeTab)}...`;
  render();
  try {
    const primaryAttempt = await runCopilotAiAttempt({
      mode,
      input,
      config: effectiveConfig,
      tabId: activeTab,
      userInput,
      contextOverride,
    });
    if (requestId !== ui.copilotActiveRequestId) return;
    let finalAttempt = primaryAttempt;
    let recoveredWithModel = "";

    if (
      primaryAttempt.usedFallback &&
      (
        primaryAttempt.fallbackReason === "empty" ||
        primaryAttempt.fallbackReason === "instruction" ||
        primaryAttempt.fallbackReason === "weak"
      )
    ) {
      const recoveryModel = pickCopilotRecoveryModel(config.model);
      if (recoveryModel) {
        const recoveryConfig = { ...effectiveConfig, model: recoveryModel };
        if (/20b/i.test(recoveryModel) && Number(recoveryConfig.timeoutSec || 0) < 300) {
          recoveryConfig.timeoutSec = 300;
        }
        recoveryConfig.timeoutMs = Math.max(15000, Number(recoveryConfig.timeoutSec || 0) * 1000);
        const retryAttempt = await runCopilotAiAttempt({
          mode,
          input,
          config: recoveryConfig,
          tabId: activeTab,
          userInput,
          contextOverride,
        });
        if (requestId !== ui.copilotActiveRequestId) return;
        if (!retryAttempt.usedFallback) {
          finalAttempt = retryAttempt;
          recoveredWithModel = recoveryModel;
        }
      }
    }

    const isPdfFocusedRequest =
      !request.isChat &&
      (
        activeTab === "pdf" ||
        isPdfGroundedQuestion(userInput) ||
        isPdfGroundedQuestion(input)
      );

    const shouldForcePdfRetry =
      isPdfFocusedRequest &&
      (
        finalAttempt.usedFallback ||
        activeTab === "pdf" ||
        activeTab === "npcs" ||
        str(finalAttempt.processed.text).length < (activeTab === "npcs" ? 320 : 220)
      );

    if (shouldForcePdfRetry) {
      const selectedPdfFile = str(buildGlobalCopilotContext(activeTab)?.selectedPdfFile) || "the selected PDF";
      const pdfRetryPrompt = buildPdfFocusedRetryPrompt(activeTab, selectedPdfFile, userInput || input);
      const pdfRetryConfig = {
        ...effectiveConfig,
        maxOutputTokens: Math.max(Number(effectiveConfig.maxOutputTokens || 0) || 0, activeTab === "npcs" ? 1300 : 1100),
        timeoutSec: Math.max(Number(effectiveConfig.timeoutSec || 0), 480),
      };
      pdfRetryConfig.timeoutMs = Math.max(15000, Number(pdfRetryConfig.timeoutSec || 0) * 1000);
      const pdfRetryAttempt = await runCopilotAiAttempt({
        mode: activeTab === "npcs" ? "npc" : "prep",
        input: pdfRetryPrompt,
        config: pdfRetryConfig,
        tabId: activeTab,
        userInput,
      });
      if (requestId !== ui.copilotActiveRequestId) return;
      if (
        !pdfRetryAttempt.usedFallback &&
        (
          finalAttempt.usedFallback ||
          str(pdfRetryAttempt.processed.text).length > str(finalAttempt.processed.text).length
        )
      ) {
        finalAttempt = pdfRetryAttempt;
      }
    }

    const shouldForceSessionExpand =
      activeTab === "sessions" &&
      !finalAttempt.usedFallback &&
      str(finalAttempt.processed.text).length < 220 &&
      /\b(idea|ideas|hook|hooks|scene|scenes|encounter|encounters|village|town|quest|npc|prep|session|run|opening)\b/i.test(
        userInput
      );

    if (shouldForceSessionExpand) {
      const expandedRequest = [
        buildGlobalCopilotSeedPrompt("sessions"),
        "",
        "GM request:",
        userInput,
        "",
        "Expand with concrete details and at least 6 actionable bullets in Next Prep.",
      ].join("\n");
      const expansionAttempt = await runCopilotAiAttempt({
        mode: "session",
        input: expandedRequest,
        config: effectiveConfig,
        tabId: activeTab,
        userInput,
      });
      if (requestId !== ui.copilotActiveRequestId) return;
      if (!expansionAttempt.usedFallback && str(expansionAttempt.processed.text).length > str(finalAttempt.processed.text).length) {
        finalAttempt = expansionAttempt;
      }
    }

    ui.copilotDraft.output = finalAttempt.processed.text;
    if (!autoRun && !finalAttempt.usedFallback) {
      addAiHistoryTurn({
        tabId: activeTab,
        role: "assistant",
        mode,
        text: finalAttempt.processed.text,
      });
      ui.copilotPendingFallbackMemory = null;
    }
    clearAiError();
    if (recoveredWithModel) {
      ui.copilotPendingFallbackMemory = null;
      const nextConfig = ensureAiConfig();
      nextConfig.model = recoveredWithModel;
      state.meta.aiConfig = nextConfig;
      saveState();
      ui.copilotMessage = `Recovered reply using ${getAiModelDisplayName(recoveredWithModel)}. Default model switched to this one.`;
    } else if (finalAttempt.usedFallback) {
      ui.copilotPendingFallbackMemory = !autoRun
        ? {
            tabId: activeTab,
            mode,
            text: finalAttempt.processed.text,
          }
        : null;
      const reason = finalAttempt.fallbackReason;
      if (request.isChat) {
        ui.copilotMessage =
          reason === "empty"
            ? "Local AI returned empty output. Showing built-in fallback instead of a real model reply (not saved to memory)."
            : reason === "instruction"
              ? "Local AI returned instruction text, not a usable answer. Showing built-in fallback (not saved to memory)."
              : reason === "weak"
                ? "Local AI returned an incomplete answer. Showing built-in fallback (not saved to memory)."
                : "Fallback used for chat response.";
      } else {
        ui.copilotMessage = `Fallback used for ${getTabLabel(activeTab)} output (not saved to memory).`;
      }
      if (activeTab === "pdf") {
        ui.copilotMessage += " This fallback is generic and not grounded in PDF/book content.";
      }
    } else {
      ui.copilotPendingFallbackMemory = null;
      if (request.isChat) {
        ui.copilotMessage = `Chat response generated with ${getAiModelDisplayName(str(finalAttempt.result?.model) || effectiveConfig.model)}.`;
      } else {
        ui.copilotMessage = autoRun
          ? `Auto-generated ${getTabLabel(activeTab)} output.`
          : "Generated.";
      }
    }
  } catch (err) {
    if (requestId !== ui.copilotActiveRequestId) return;
    const message = recordAiError("AI generate", err);
    if (/timed out/i.test(message)) {
      ui.copilotDraft.output = generateFallbackAiOutput({
        mode,
        input: userInput || input,
        tabId: activeTab,
      });
      ui.copilotMessage = `Local AI timed out. Showing built-in fallback instead (not saved to memory).${
        activeTab === "pdf" ? " This fallback is generic and not grounded in PDF/book content." : ""
      } ${message}`;
    } else {
      ui.copilotDraft.output = "";
      ui.copilotMessage = `AI generate failed: ${message}`;
    }
    ui.copilotPendingFallbackMemory = null;
  } finally {
    if (requestId === ui.copilotActiveRequestId) {
      ui.copilotBusy = false;
      render();
    }
  }
}

function buildPdfFocusedRetryPrompt(tabId, selectedPdfFile, requestText) {
  const cleanRequest = str(requestText);
  if (tabId === "npcs") {
    return [
      "Use only the selected PDF context already provided with this request.",
      `Selected PDF: ${selectedPdfFile}`,
      "Create table-ready NPCs grounded in the book.",
      "If the GM asked for a few or multiple NPCs, return 2 to 4 NPCs separated by a line containing only ---.",
      "Prefer named or clearly implied recurring figures the party is likely to meet early.",
      "Do not invent unsupported lore. If a role is inferred from the book context, mark it as inferred in Notes.",
      "For each NPC, answer in this exact structure:",
      "Name:",
      "Role:",
      "Agenda:",
      "Disposition:",
      "Notes:",
      "- Book anchor:",
      "- Why the party works with or against them:",
      "- Current pressure or fear:",
      "- Voice and mannerisms:",
      "- First impression or look:",
      "- Hidden truth or complication:",
      "- Best way to use them in the next session:",
      "",
      "GM request:",
      cleanRequest,
    ].join("\n");
  }
  return [
    "Use only the selected PDF context already provided with this request.",
    `Selected PDF: ${selectedPdfFile}`,
    "Answer in this exact structure:",
    "Main Threat Summary:",
    "- 2 to 4 bullets grounded in the PDF",
    "5 Ways To Run Chapter One:",
    "1. one concrete GM approach",
    "2. one concrete GM approach",
    "3. one concrete GM approach",
    "4. one concrete GM approach",
    "5. one concrete GM approach",
    "If a detail is not confirmed by the indexed PDF context, say it is not confirmed instead of guessing.",
    "",
    "GM request:",
    cleanRequest,
  ].join("\n");
}

async function copyGlobalAiOutput() {
  const text = str(ui.copilotDraft.output);
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    ui.copilotMessage = "Loremaster output copied.";
  } catch {
    ui.copilotMessage = "Copy failed. Select output manually and copy.";
  }
  render();
}

async function applyGlobalAiOutput() {
  const text = str(ui.copilotDraft.output);
  if (!text) {
    ui.copilotMessage = "No Loremaster output to apply.";
    render();
    return;
  }

  if (activeTab === "npcs") {
    createNpcFromAi(text);
    return;
  }
  if (activeTab === "quests") {
    createQuestFromAi(text);
    return;
  }
  if (activeTab === "locations") {
    createLocationFromAi(text);
    return;
  }
  if (activeTab === "capture") {
    createCaptureEntry("AI", text, getResolvedCaptureSessionId());
    ui.copilotMessage = "Added AI output to live capture.";
    render();
    return;
  }
  if (activeTab === "kingdom") {
    appendKingdomAiNote(text);
    ui.copilotMessage = "Appended AI output to kingdom notes.";
    render();
    return;
  }
  if (activeTab === "writing") {
    ui.writingDraft.output = text;
    ui.copilotMessage = "Sent AI output to Writing Helper.";
    render();
    return;
  }
  if (activeTab === "pdf") {
    const query = extractQueryFromAiOutput(text);
    if (!query) {
      ui.copilotMessage = "Could not extract a query from AI output.";
      render();
      return;
    }
    ui.pdfSearchQuery = query;
    ui.copilotMessage = `Using AI query: ${query}`;
    if (desktopApi) {
      await runPdfSearch(query, 20);
      return;
    }
    render();
    return;
  }

  const latest = ensureLatestSessionForAi();
  if (!latest) {
    ui.copilotMessage = "No session available to attach AI output.";
    render();
    return;
  }

  if (activeTab === "sessions") {
    const parsedSummary = extractLabeledBlock(text, "Summary");
    const parsedPrep = extractLabeledBlock(text, "Next Prep");
    if (parsedSummary) {
      latest.summary = parsedSummary;
    } else {
      latest.summary = text;
    }
    if (parsedPrep) {
      latest.nextPrep = injectOrReplaceAiSection(latest.nextPrep, "AI_SESSION", "AI Session Prep", parsedPrep);
    } else {
      latest.nextPrep = injectOrReplaceAiSection(latest.nextPrep, "AI_SESSION", "AI Session Prep", text);
    }
    latest.updatedAt = new Date().toISOString();
    saveState();
    ui.copilotMessage = `Applied AI output to "${latest.title}".`;
    render();
    return;
  }

  if (activeTab === "foundry") {
    latest.nextPrep = injectOrReplaceAiSection(latest.nextPrep, "AI_FOUNDRY", "AI Foundry Handoff", text);
    latest.updatedAt = new Date().toISOString();
    saveState();
    ui.copilotMessage = `Attached Foundry handoff notes to "${latest.title}".`;
    render();
    return;
  }

  latest.nextPrep = injectOrReplaceAiSection(latest.nextPrep, "AI_DASHBOARD", "AI Prep Plan", text);
  latest.updatedAt = new Date().toISOString();
  saveState();
  ui.copilotMessage = `Attached AI prep plan to "${latest.title}".`;
  render();
}

function ensureLatestSessionForAi() {
  const latest = getLatestSession();
  if (latest) return latest;
  const now = new Date().toISOString();
  const created = {
    id: uid(),
    title: "Session (AI Draft)",
    date: "",
    arc: "",
    kingdomTurn: "",
    summary: "",
    nextPrep: "",
    createdAt: now,
    updatedAt: now,
  };
  state.sessions.unshift(created);
  saveState();
  return created;
}

function injectOrReplaceAiSection(currentText, key, title, body) {
  const start = `<!-- ${key}_START -->`;
  const end = `<!-- ${key}_END -->`;
  const section = `${start}\n### ${title}\n${str(body)}\n${end}`;
  const markerRegex = new RegExp(`${escapeRegex(start)}[\\s\\S]*?${escapeRegex(end)}`, "m");
  const base = str(currentText);
  if (!base) return section;
  if (markerRegex.test(base)) return base.replace(markerRegex, section).trim();
  return `${base}\n\n${section}`;
}

function extractLabeledBlock(text, label) {
  const source = String(text || "");
  if (!source) return "";
  const regex = new RegExp(`${escapeRegex(label)}\\s*:\\s*([\\s\\S]*?)(?=\\n[A-Za-z][A-Za-z ]{1,28}:|$)`, "i");
  const match = source.match(regex);
  return match ? str(match[1]) : "";
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

function extractQueryFromAiOutput(text) {
  const labeled = extractLabeledBlock(text, "Query");
  if (labeled) return labeled.split(/\n/)[0].trim().slice(0, 120);
  const firstLine = String(text || "")
    .split(/\n+/)
    .map((line) => line.trim())
    .find((line) => line.length > 2 && line.length < 120);
  return str(firstLine).replace(/^[-*]\s*/, "");
}

function splitNpcEntriesFromAi(text) {
  const source = str(text).replace(/\r\n?/g, "\n").trim();
  if (!source) return [];
  const lines = source.split("\n");
  const blocks = [];
  let current = [];
  for (const rawLine of lines) {
    const line = str(rawLine);
    if (/^---+\s*$/.test(line)) {
      if (current.length) {
        blocks.push(current.join("\n").trim());
        current = [];
      }
      continue;
    }
    if (/^Name\s*:/i.test(line) && current.some((entry) => /^Name\s*:/i.test(entry))) {
      blocks.push(current.join("\n").trim());
      current = [line];
      continue;
    }
    current.push(line);
  }
  if (current.length) blocks.push(current.join("\n").trim());
  return blocks.filter(Boolean);
}

function createNpcFromAi(text) {
  const blocks = splitNpcEntriesFromAi(text);
  const now = new Date().toISOString();
  const createdNames = [];
  for (const block of blocks.length ? blocks : [text]) {
    const name = extractLabeledBlock(block, "Name") || guessTitleFromText(block, "AI NPC");
    const role = extractLabeledBlock(block, "Role");
    const agenda = extractLabeledBlock(block, "Agenda");
    const disposition = extractLabeledBlock(block, "Disposition") || "Neutral";
    const notes = buildNpcNotesFromAi(block) || block;
    if (!name && !notes) continue;
    state.npcs.unshift({
      id: uid(),
      name,
      role,
      agenda,
      disposition,
      notes,
      createdAt: now,
      updatedAt: now,
    });
    createdNames.push(name);
  }
  if (!createdNames.length) {
    ui.copilotMessage = "Could not extract any NPCs from the AI output.";
    render();
    return;
  }
  saveState();
  ui.copilotMessage = createdNames.length === 1
    ? `Created NPC: ${createdNames[0]}`
    : `Created ${createdNames.length} NPCs: ${createdNames.slice(0, 3).join(", ")}${createdNames.length > 3 ? "..." : ""}`;
  render();
}

function createQuestFromAi(text) {
  const statusRaw = (extractLabeledBlock(text, "Status") || "open").toLowerCase();
  const status = ["open", "in-progress", "blocked", "completed", "failed"].includes(statusRaw) ? statusRaw : "open";
  const title = extractLabeledBlock(text, "Title") || guessTitleFromText(text, "AI Quest");
  const objective = extractLabeledBlock(text, "Objective");
  const giver = extractLabeledBlock(text, "Giver");
  const stakes = extractLabeledBlock(text, "Stakes") || text;
  const now = new Date().toISOString();
  state.quests.unshift({
    id: uid(),
    title,
    status,
    objective,
    giver,
    stakes,
    createdAt: now,
    updatedAt: now,
  });
  saveState();
  ui.copilotMessage = `Created Quest: ${title}`;
  render();
}

function createLocationFromAi(text) {
  const name = extractLabeledBlock(text, "Name") || guessTitleFromText(text, "AI Location");
  const hex = extractLabeledBlock(text, "Hex");
  const whatChanged = extractLabeledBlock(text, "What Changed");
  const notes = extractLabeledBlock(text, "Notes") || text;
  const now = new Date().toISOString();
  state.locations.unshift({
    id: uid(),
    name,
    hex,
    whatChanged,
    notes,
    createdAt: now,
    updatedAt: now,
  });
  saveState();
  ui.copilotMessage = `Created Location: ${name}`;
  render();
}

function guessTitleFromText(text, fallback) {
  const line = String(text || "")
    .split(/\n+/)
    .map((entry) => entry.replace(/^[-*#\s]+/, "").trim())
    .find(Boolean);
  return str(line).slice(0, 80) || fallback;
}

function escapeRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function generatePrepQueue(mode) {
  const latest = getLatestSession();
  const sourceText = `${latest?.summary || ""} ${latest?.nextPrep || ""}`;
  const items = [];
  const seen = new Set();

  const add = (label, minutes) => {
    const clean = str(label);
    const mins = Number.parseInt(String(minutes || "0"), 10);
    if (!clean || Number.isNaN(mins) || mins <= 0) return;
    const key = clean.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    items.push({
      id: `prep-${slugify(clean)}`,
      label: clean,
      minutes: mins,
    });
  };

  add("Read recap and pick the exact opening scene.", 6);
  add("Verify Foundry scene, token lights, and initiative setup.", 8);

  const openQuests = state.quests
    .filter((q) => q.status !== "completed" && q.status !== "failed")
    .sort((a, b) => relevanceScore(sourceText, b.title) - relevanceScore(sourceText, a.title))
    .slice(0, 4);
  for (const quest of openQuests) {
    add(`Prep one concrete beat for "${quest.title}".`, 8);
  }

  const npcFocus = getMentionedOrRecent(state.npcs, "name", sourceText, 3);
  for (const npc of npcFocus) {
    add(`Prep voice + agenda for ${npc.name}.`, 6);
  }

  const locationFocus = getMentionedOrRecent(state.locations, "name", sourceText, 2);
  for (const location of locationFocus) {
    add(`Update consequence state at ${location.name}.`, 5);
  }

  if ((state.meta.pdfIndexedCount || 0) > 0) {
    const terms = suggestSearchTerms(sourceText, 3);
    if (terms.length) {
      add(`Run PDF Intel checks: ${terms.join(", ")}.`, 8);
    } else {
      add("Run one PDF Intel rules check for expected edge-case.", 6);
    }
  } else {
    add("Index PDFs in PDF Intel (one-time setup/refresh).", 12);
  }

  if (latest && (str(latest.kingdomTurn) || hasKingdomSignals(sourceText))) {
    add("Resolve campaign bookkeeping prep (resources, unrest, upkeep).", 10);
  }

  add("Prep one fallback encounter + one social wildcard.", 10);

  const budget = mode === 30 ? 30 : mode === 90 ? 90 : 60;
  const queued = [];
  let used = 0;
  for (const item of items) {
    if (used + item.minutes <= budget || queued.length < 3) {
      queued.push(item);
      used += item.minutes;
    }
  }

  return queued;
}

function sessionSortKey(session) {
  const dateKey = safeDate(session?.date || "");
  if (dateKey > 0) return dateKey;
  const created = Date.parse(session?.createdAt || "");
  return Number.isNaN(created) ? 0 : created;
}

function generateSmartChecklist(options = {}) {
  const includeArchived = options?.includeArchived === true;
  const latest = getLatestSession();
  const latestText = `${latest?.summary || ""} ${latest?.nextPrep || ""}`;
  const items = [];
  const seen = new Set();

  const add = (label) => {
    const clean = str(label);
    if (!clean) return;
    const key = clean.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    items.push({ id: `check-${slugify(clean)}`, label: clean });
  };

  add("Confirm Foundry scene, token vision, and initiative tools are ready.");
  add("Read a 60-90 second recap of last session out loud.");
  add("Prepare one backup encounter and one social complication.");

  if (latest && (str(latest.kingdomTurn) || hasKingdomSignals(latestText))) {
    add("Run campaign bookkeeping first (resources, unrest, claims, build queue).");
  }

  const activeQuests = state.quests
    .filter((q) => q.status !== "completed" && q.status !== "failed")
    .sort((a, b) => relevanceScore(latestText, b.title) - relevanceScore(latestText, a.title))
    .slice(0, 3);

  for (const quest of activeQuests) {
    add(`Prep next concrete beat for quest: ${quest.title}.`);
  }

  const npcFocus = getMentionedOrRecent(state.npcs, "name", latestText, 3);
  for (const npc of npcFocus) {
    add(`Refresh voice + motivation for NPC: ${npc.name}.`);
  }

  const locationFocus = getMentionedOrRecent(state.locations, "name", latestText, 2);
  for (const location of locationFocus) {
    add(`Update consequences and sensory detail for location: ${location.name}.`);
  }

  if ((state.meta.pdfIndexedCount || 0) > 0) {
    const terms = suggestSearchTerms(latestText, 3);
    if (terms.length) {
      add(`PDF Intel search before session: ${terms.join(", ")}.`);
    } else {
      add("Use PDF Intel to verify one rule likely to come up this session.");
    }
  } else {
    add("Index PDFs in PDF Intel before final prep pass.");
  }

  const generated = items.slice(0, 12);
  const custom = ensureCustomChecklistItems();
  const overrides = ensureChecklistOverrides();
  const archived = ensureChecklistArchived();
  const combined = [...generated, ...custom];

  return combined
    .map((item) => {
      const override = normalizeChecklistLabel(overrides[item.id]);
      return {
        id: item.id,
        label: override || item.label,
      };
    })
    .filter((item) => normalizeChecklistLabel(item.label))
    .filter((item) => (includeArchived ? true : !archived[item.id]));
}

function generateWrapUpForLatestSession() {
  const latest = getLatestSession();
  if (!latest) {
    ui.sessionMessage = "No sessions available for smart wrap-up.";
    render();
    return;
  }
  generateWrapUpForSession(latest.id);
}

function openSessionCloseWizard(defaultSessionId = "") {
  const latest = getLatestSession();
  ui.wizardOpen = true;
  ui.wizardDraft = {
    sessionId: defaultSessionId || latest?.id || "",
    highlights: "",
    cliffhanger: "",
    playerIntent: "",
  };
  render();
}

function closeSessionCloseWizard() {
  ui.wizardOpen = false;
  ui.wizardDraft = {
    sessionId: "",
    highlights: "",
    cliffhanger: "",
    playerIntent: "",
  };
  render();
}

function generateWrapUpForSession(sessionId, options = {}) {
  const { wizardAnswers = null, sceneOpeners = [], silent = false } = options;
  const session = state.sessions.find((s) => s.id === sessionId);
  if (!session) {
    if (!silent) {
      ui.sessionMessage = "Could not find that session.";
      render();
    }
    return null;
  }

  const bullets = buildWrapUpBullets(session, wizardAnswers);
  let nextPrep = injectOrReplaceSmartWrapSection(
    session.nextPrep || "",
    buildSmartWrapSection(bullets)
  );

  if (sceneOpeners.length) {
    nextPrep = injectOrReplaceSceneOpenersSection(
      nextPrep,
      buildSceneOpenersSection(sceneOpeners)
    );
  }

  session.nextPrep = nextPrep;
  session.updatedAt = new Date().toISOString();

  saveState();
  const result = { session, bullets, sceneOpeners };

  if (!silent) {
    ui.sessionMessage = `Smart wrap-up generated for "${session.title}" (${bullets.length} bullets).`;
    render();
  }

  return result;
}

function buildWrapUpBullets(session, wizardAnswers = null) {
  const sourceText = `${session.summary || ""} ${session.nextPrep || ""}`;
  const wizardText = wizardAnswers
    ? `${wizardAnswers.highlights || ""} ${wizardAnswers.cliffhanger || ""} ${wizardAnswers.playerIntent || ""}`
    : "";
  const fullSource = `${sourceText} ${wizardText}`.trim();
  const bullets = [];
  const seen = new Set();

  const add = (bullet) => {
    const clean = str(bullet);
    if (!clean) return;
    const key = clean.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    bullets.push(clean);
  };

  const activeQuests = state.quests
    .filter((q) => q.status !== "completed" && q.status !== "failed")
    .map((q) => ({ quest: q, score: relevanceScore(fullSource, q.title) }))
    .sort((a, b) => b.score - a.score)
    .slice(0, 3)
    .map((entry) => entry.quest);

  for (const quest of activeQuests) {
    add(`Advance "${quest.title}" with one clear opening scene and consequence.`);
  }

  const npcFocus = getMentionedOrRecent(state.npcs, "name", fullSource, 2);
  for (const npc of npcFocus) {
    add(`Decide ${npc.name}'s immediate reaction and ask for next session.`);
  }

  const locationFocus = getMentionedOrRecent(state.locations, "name", fullSource, 2);
  for (const location of locationFocus) {
    add(`Update world-state at ${location.name} and prep one reveal.`);
  }

  if (str(session.kingdomTurn) || hasKingdomSignals(fullSource)) {
    add("Resolve campaign bookkeeping at start of next session (resources, unrest, buildings, claims).");
  }

  if (wizardAnswers) {
    if (str(wizardAnswers.highlights)) {
      add(`Carry forward this key beat: ${condenseLine(wizardAnswers.highlights)}.`);
    }
    if (str(wizardAnswers.cliffhanger)) {
      add(`Open next session by resolving: ${condenseLine(wizardAnswers.cliffhanger)}.`);
    }
    if (str(wizardAnswers.playerIntent)) {
      add(`Prioritize player-declared intent: ${condenseLine(wizardAnswers.playerIntent)}.`);
    }
  }

  if ((state.meta.pdfIndexedCount || 0) > 0) {
    const terms = suggestSearchTerms(
      `${fullSource} ${activeQuests.map((q) => q.title).join(" ")}`,
      4
    );
    if (terms.length) {
      add(`Use PDF Intel to verify rules/lore for: ${terms.join(", ")}.`);
    }
  } else {
    add("Index your PDFs before next prep to speed up rules checks.");
  }

  add("Prepare one fallback encounter and one non-combat complication.");

  return bullets.slice(0, 8);
}

function buildSmartWrapSection(bullets) {
  const stamp = new Date().toISOString().slice(0, 10);
  const lines = bullets.map((b) => `- ${b}`).join("\n");
  return `<!-- SMART_WRAPUP_START -->
### Smart Wrap-Up (${stamp})
${lines}
<!-- SMART_WRAPUP_END -->`;
}

function buildSceneOpenersSection(openers) {
  const stamp = new Date().toISOString().slice(0, 10);
  const lines = openers.map((opener, i) => `${i + 1}. ${opener}`).join("\n");
  return `<!-- SMART_SCENES_START -->
### Suggested Scene Openers (${stamp})
${lines}
<!-- SMART_SCENES_END -->`;
}

function injectOrReplaceSmartWrapSection(currentText, smartSection) {
  const markerRegex = /<!-- SMART_WRAPUP_START -->[\s\S]*?<!-- SMART_WRAPUP_END -->/m;
  if (markerRegex.test(currentText)) {
    return currentText.replace(markerRegex, smartSection).trim();
  }
  const base = str(currentText);
  return base ? `${smartSection}\n\n${base}` : smartSection;
}

function injectOrReplaceSceneOpenersSection(currentText, sceneSection) {
  const markerRegex = /<!-- SMART_SCENES_START -->[\s\S]*?<!-- SMART_SCENES_END -->/m;
  if (markerRegex.test(currentText)) {
    return currentText.replace(markerRegex, sceneSection).trim();
  }
  const base = str(currentText);
  return base ? `${sceneSection}\n\n${base}` : sceneSection;
}

function generateSceneOpeners(session, wizardAnswers) {
  const source = `${session.summary || ""} ${session.nextPrep || ""} ${wizardAnswers?.highlights || ""} ${
    wizardAnswers?.cliffhanger || ""
  } ${wizardAnswers?.playerIntent || ""}`;
  const quest = state.quests
    .filter((q) => q.status !== "completed" && q.status !== "failed")
    .sort((a, b) => relevanceScore(source, b.title) - relevanceScore(source, a.title))[0];
  const npc = getMentionedOrRecent(state.npcs, "name", source, 1)[0];
  const location = getMentionedOrRecent(state.locations, "name", source, 1)[0];

  const questTitle = quest?.title || "the current frontier threat";
  const npcName = npc?.name || "a known local contact";
  const locationName = location?.name || "the nearest frontier settlement";
  const cliff = str(wizardAnswers?.cliffhanger)
    ? condenseLine(wizardAnswers.cliffhanger)
    : "the unresolved pressure from last session";
  const playerIntent = str(wizardAnswers?.playerIntent)
    ? condenseLine(wizardAnswers.playerIntent)
    : "the party's stated next objective";

  const options = [];
  const seen = new Set();
  const add = (line) => {
    const clean = str(line);
    if (!clean) return;
    const key = clean.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    options.push(clean);
  };

  add(
    `Cold open at ${locationName}: ${npcName} interrupts with urgent news tied to "${questTitle}".`
  );
  add(
    `Immediate consequence opener: "${cliff}" escalates before the party can settle in.`
  );
  add(
    `Player-intent opener: frame the first scene around "${playerIntent}" and attach one new complication.`
  );

  if (hasKingdomSignals(source) || str(session.kingdomTurn)) {
    add("Kingdom pressure opener: start with a council/upkeep decision before adventure scenes.");
  }

  return options.slice(0, 3);
}

function condenseLine(text) {
  const clean = str(text).replace(/\s+/g, " ");
  if (!clean) return "";
  const sentence = clean.split(/[.!?]/)[0] || clean;
  return sentence.slice(0, 140).replace(/[,;:\- ]+$/g, "");
}

function runWritingHelper() {
  const input = str(ui.writingDraft.input);
  if (!input) {
    ui.writingDraft.output = "";
    ui.sessionMessage = "Writing Helper: add some draft text first.";
    render();
    return;
  }

  const mode = ui.writingDraft.mode || "session";
  const cleaned = basicAutoCorrect(input);
  const output = generateStructuredText(cleaned, mode);
  ui.writingDraft.output = output;
  ui.sessionMessage = "Writing Helper generated cleaned text.";
  render();
}

async function testLocalAiConnection() {
  if (!desktopApi?.testLocalAi) {
    const message = "Desktop local AI bridge is not available in this runtime.";
    ui.aiMessage = message;
    ui.copilotMessage = message;
    render();
    return;
  }

  const config = ensureAiConfig();
  ui.aiTestAt = new Date().toISOString();
  ui.aiTestStatus = `Running connection test (${getTabLabel(activeTab)})...`;
  ui.aiBusy = true;
  ui.aiMessage = "Testing local AI connection...";
  ui.copilotMessage = "Testing local AI connection...";
  render();
  try {
    const result = await desktopApi.testLocalAi(config);
    const message = str(result?.message) || "Local AI connection ok.";
    ui.aiMessage = message;
    ui.copilotMessage = message;
    ui.aiTestStatus = `Passed: ${message}`;
    ui.aiTestAt = new Date().toISOString();
    ui.aiModels = Array.isArray(result?.models) ? result.models : ui.aiModels;
    clearAiError();
  } catch (err) {
    const message = recordAiError("AI connection test", err);
    ui.aiMessage = `AI test failed: ${message}`;
    ui.copilotMessage = `AI test failed: ${message}`;
    ui.aiTestStatus = `Failed: ${message}`;
    ui.aiTestAt = new Date().toISOString();
  } finally {
    ui.aiBusy = false;
    render();
  }
}

async function runWritingHelperWithLocalAi() {
  if (!desktopApi?.generateLocalAiText) {
    ui.aiMessage = "Desktop local AI bridge is not available in this runtime.";
    render();
    return;
  }

  const input = str(ui.writingDraft.input);
  if (!input) {
    ui.sessionMessage = "Writing Helper: add some draft text first.";
    render();
    return;
  }

  const mode = ui.writingDraft.mode || "session";
  const config = ensureAiConfig();
  const context = collectAiCampaignContext();

  ui.aiBusy = true;
  ui.aiMessage = "Generating with local AI...";
  render();

  try {
    const response = await desktopApi.generateLocalAiText({
      mode,
      input,
      context,
      config,
    });
    const processed = processAiOutputWithFallback({
      rawText: response?.text || "",
      mode,
      input,
      source: "writing",
      tabId: "writing",
    });
    const finalOutput = processed.text;
    const usedFallback = processed.usedFallback || response?.usedFallback === true;

    ui.writingDraft.output = finalOutput;
    ui.sessionMessage = usedFallback
      ? `AI returned instruction-style text, so DM Helper generated a usable ${mode} draft automatically.`
      : `Local AI generated text using ${str(response?.model) || config.model}.`;
    if (ui.writingDraft.autoLink) {
      const autoResult = autoConnectWritingOutputToLatestSession({ silent: true, source: "AI output" });
      if (autoResult.applied) {
        ui.sessionMessage = `${ui.sessionMessage} Auto-connected ${autoResult.totalLinks} reference(s) to latest session prep.`;
      }
    }
    ui.aiMessage = `Connected to ${str(response?.endpoint) || config.endpoint}`;
    clearAiError();
  } catch (err) {
    const message = recordAiError("Writing helper generation", err);
    ui.sessionMessage = `Local AI generation failed: ${message}`;
  } finally {
    ui.aiBusy = false;
    render();
  }
}

function collectAiCampaignContext() {
  const latest = getLatestSession();
  const kingdom = getKingdomState();
  const kingdomProfile = getActiveKingdomProfile();
  const openQuests = state.quests.filter((q) => q.status !== "completed" && q.status !== "failed").slice(0, 6);
  const recentSessions = [...state.sessions]
    .sort((a, b) => safeDate(b.date || b.updatedAt || b.createdAt) - safeDate(a.date || a.updatedAt || a.createdAt))
    .slice(0, 6);
  const indexedFiles = Array.isArray(state?.meta?.pdfIndexedFiles)
    ? state.meta.pdfIndexedFiles.map((name) => str(name)).filter(Boolean)
    : [];
  const summaryBriefs = Object.values(getPdfSummaryMap())
    .filter((item) => item && typeof item === "object")
    .map((item) => ({
      fileName: str(item.fileName),
      summary: str(item.summary).replace(/\s+/g, " ").slice(0, 420),
      updatedAt: str(item.updatedAt),
    }))
    .filter((item) => item.fileName && item.summary)
    .slice(0, 12);
  return {
    latestSession: latest
      ? {
          title: latest.title,
          summary: latest.summary,
          nextPrep: latest.nextPrep,
          arc: latest.arc,
          kingdomTurn: latest.kingdomTurn,
        }
      : null,
    recentSessions: recentSessions.map((session) => ({
      title: session.title,
      date: session.date,
      summary: session.summary,
      nextPrep: session.nextPrep,
      arc: session.arc,
    })),
    openQuests: openQuests.map((q) => ({ title: q.title, objective: q.objective, stakes: q.stakes })),
    quests: state.quests.slice(0, 12).map((q) => ({
      title: q.title,
      status: q.status,
      objective: q.objective,
      giver: q.giver,
      stakes: q.stakes,
    })),
    npcs: state.npcs.slice(0, 12).map((n) => ({
      name: n.name,
      role: n.role,
      agenda: n.agenda,
      disposition: n.disposition,
      notes: n.notes,
    })),
    locations: state.locations.slice(0, 10).map((l) => ({
      name: l.name,
      hex: l.hex,
      whatChanged: l.whatChanged,
      notes: l.notes,
    })),
    kingdom: buildKingdomAiContext(kingdom, kingdomProfile),
    pdfIndexedFileCount: Number.parseInt(String(state?.meta?.pdfIndexedCount || indexedFiles.length || 0), 10) || 0,
    pdfIndexedFiles: indexedFiles.slice(0, 60),
    pdfSummaryBriefs: summaryBriefs,
  };
}

async function copyWritingOutput() {
  const text = str(ui.writingDraft.output);
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    ui.sessionMessage = "Writing Helper output copied.";
  } catch {
    ui.sessionMessage = "Copy failed. Select output manually and copy.";
  }
  render();
}

function applyWritingOutputToLatestSession(field) {
  const text = str(ui.writingDraft.output);
  if (!text) {
    ui.sessionMessage = "No Writing Helper output to apply.";
    render();
    return;
  }
  const latest = getLatestSession();
  if (!latest) {
    ui.sessionMessage = "No session found to apply output.";
    render();
    return;
  }
  if (field === "summary") {
    latest.summary = text;
  } else {
    latest.nextPrep = text;
  }
  latest.updatedAt = new Date().toISOString();
  saveState();
  ui.sessionMessage = `Applied Writing Helper output to "${latest.title}" (${field}).`;
  render();
}

function autoConnectWritingOutputToLatestSession(options = {}) {
  const text = str(ui.writingDraft.output);
  if (!text) {
    if (!options.silent) {
      ui.sessionMessage = "No Writing Helper output to auto-connect.";
      render();
    }
    return { applied: false, totalLinks: 0 };
  }

  const latest = getLatestSession();
  if (!latest) {
    if (!options.silent) {
      ui.sessionMessage = "No session found for auto-connect.";
      render();
    }
    return { applied: false, totalLinks: 0 };
  }

  const links = collectEntityLinksFromText(text);
  const totalLinks = links.npcs.length + links.quests.length + links.locations.length;
  if (!totalLinks) {
    if (!options.silent) {
      ui.sessionMessage = "Auto-connect found no matching NPC/quest/location names.";
      render();
    }
    return { applied: false, totalLinks: 0 };
  }

  const sourceLabel = str(options.source) || "Writing Helper output";
  const section = buildAutoLinksSection(links, sourceLabel);
  latest.nextPrep = injectOrReplaceAutoLinksSection(latest.nextPrep || "", section);
  latest.updatedAt = new Date().toISOString();
  saveState();

  if (!options.silent) {
    ui.sessionMessage = `Auto-connected ${totalLinks} reference(s) into latest session prep.`;
    render();
  }

  return { applied: true, totalLinks };
}

function collectEntityLinksFromText(text) {
  return {
    npcs: findEntityMentions(text, state.npcs, "name", 6),
    quests: findEntityMentions(text, state.quests, "title", 6),
    locations: findEntityMentions(text, state.locations, "name", 6),
  };
}

function findEntityMentions(text, entities, field, limit) {
  const source = str(text).toLowerCase();
  if (!source) return [];
  const scored = [];
  for (const entity of entities || []) {
    const name = str(entity?.[field]);
    if (!name) continue;
    const nameLower = name.toLowerCase();
    const directHit = source.includes(nameLower);
    const score = relevanceScore(source, name);
    if (!directHit && score < 2) continue;
    scored.push({ name, score: directHit ? score + 2 : score });
  }
  scored.sort((a, b) => b.score - a.score || a.name.localeCompare(b.name));
  const seen = new Set();
  const out = [];
  for (const item of scored) {
    const key = item.name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(item.name);
    if (out.length >= limit) break;
  }
  return out;
}

function buildAutoLinksSection(links, sourceLabel) {
  const stamp = new Date().toISOString().slice(0, 10);
  const npcLine = links.npcs.length ? links.npcs.join(", ") : "None";
  const questLine = links.quests.length ? links.quests.join(", ") : "None";
  const locationLine = links.locations.length ? links.locations.join(", ") : "None";
  return `<!-- AUTO_LINKS_START -->
### Auto-Linked References (${stamp})
Source: ${sourceLabel}
- NPCs: ${npcLine}
- Quests: ${questLine}
- Locations: ${locationLine}
<!-- AUTO_LINKS_END -->`;
}

function injectOrReplaceAutoLinksSection(currentText, section) {
  const markerRegex = /<!-- AUTO_LINKS_START -->[\s\S]*?<!-- AUTO_LINKS_END -->/m;
  if (markerRegex.test(currentText)) {
    return currentText.replace(markerRegex, section).trim();
  }
  const base = str(currentText);
  return base ? `${base}\n\n${section}` : section;
}

function basicAutoCorrect(text) {
  let out = String(text || "");
  out = out.replace(/\r\n/g, "\n");
  out = out
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join("\n");

  const replacements = [
    ["\\bteh\\b", "the"],
    ["\\badn\\b", "and"],
    ["\\bthier\\b", "their"],
    ["\\brecieve\\b", "receive"],
    ["\\bseperate\\b", "separate"],
    ["\\boccured\\b", "occurred"],
    ["\\bdefinately\\b", "definitely"],
    ["\\bwierd\\b", "weird"],
    ["\\bcharachter\\b", "character"],
    ["\\bcharater\\b", "character"],
    ["\\bencouter\\b", "encounter"],
    ["\\bencoutered\\b", "encountered"],
    ["\\bgoverment\\b", "government"],
    ["\\bwich\\b", "which"],
    ["\\bthru\\b", "through"],
    ["\\bcoudl\\b", "could"],
    ["\\bwoudl\\b", "would"],
    ["\\bim\\b", "I'm"],
    ["\\bidk\\b", "I don't know"],
    ["\\bdm\\b", "DM"],
    ["\\bpcs\\b", "PCs"],
    ["\\bnpcs\\b", "NPCs"],
  ];
  for (const [pattern, replacement] of replacements) {
    out = out.replace(new RegExp(pattern, "gi"), replacement);
  }

  const lines = out.split("\n").map((line) => cleanSentenceLine(line));
  return lines.join("\n");
}

function cleanSentenceLine(line) {
  let out = str(line);
  if (!out) return "";
  if (/^[-*]\s+/.test(out)) {
    const bullet = out.replace(/^[-*]\s+/, "");
    return `- ${sentenceCaseAndPunctuation(bullet)}`;
  }
  return sentenceCaseAndPunctuation(out);
}

function sentenceCaseAndPunctuation(text) {
  let out = str(text);
  if (!out) return "";
  out = out.charAt(0).toUpperCase() + out.slice(1);
  if (!/[.!?]$/.test(out)) out += ".";
  return out;
}

function splitIntoIdeaLines(text) {
  return String(text || "")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function countBulletLikeLines(text) {
  return splitIntoIdeaLines(text).filter((line) => /^[-*]\s+/.test(line) || /^\d+[.)]\s+/.test(line)).length;
}

function isClearlyTruncatedOutput(text) {
  const clean = str(text).trim();
  if (!clean) return true;
  const lines = splitIntoIdeaLines(clean);
  const lastLine = lines[lines.length - 1] || clean;
  if (/[:\-]\s*$/.test(lastLine)) return true;
  if (/[.!?]$/.test(lastLine)) return false;
  const lastWord = str(lastLine).toLowerCase().split(/\s+/).filter(Boolean).pop() || "";
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

function generateStructuredText(cleanedInput, mode) {
  const lines = splitIntoIdeaLines(cleanedInput);
  const joined = lines.join(" ");
  const lower = joined.toLowerCase();

  if (mode === "prep") {
    return lines.map((line) => (line.startsWith("- ") ? line : `- ${sentenceCaseAndPunctuation(line)}`)).join("\n");
  }

  if (mode === "recap") {
    const intro = lines[0] ? `Last session, ${lowercaseFirst(lines[0])}` : "Last session, the party pushed the story forward.";
    const middle = lines[1] ? `They also ${lowercaseFirst(lines[1])}` : "";
    const close = "Now the next chapter begins.";
    return [sentenceCaseAndPunctuation(intro), middle ? sentenceCaseAndPunctuation(middle) : "", close]
      .filter(Boolean)
      .join(" ");
  }

  if (mode === "npc") {
    const role =
      /\bwaystation\b/.test(lower)
        ? "Waystation clerk with hidden local ties"
        : /\bvillage\b/.test(lower)
          ? "Village contact who knows more than they admit"
          : "Local contact with concealed leverage";
    const pressure =
      /\bsmuggl|contraband|dock|port\b/.test(lower)
        ? "A shipment, payoff, or contact is about to expose them."
        : /\bfrontier|road|wild\b/.test(lower)
          ? "Violence on the road is closing off their safest options."
          : "A stronger faction is forcing them to choose a side too soon.";
    return [
      "Name: Mara Vens",
      `Role: ${role}`,
      "Agenda: Protect the settlement while hiding one useful truth from the party.",
      "Disposition: Guarded but helpful",
      "Notes:",
      "- Core want: Keep their position secure long enough to survive the current pressure.",
      "- Leverage over the party or locals: They control one useful rumor, contact, or point of access the party needs.",
      `- Current pressure or fear: ${pressure}`,
      "- Voice and mannerisms: Low voice, clipped answers, and long pauses before saying anything costly.",
      "- First impression or look: Travel-stained clothes, watchful eyes, and a habit of standing where they can see every exit.",
      "- Hidden truth or complication: They already made one compromise with the wrong people and are trying to keep it buried.",
      "- Best way to use them in the next session: Let them point the party toward the next lead, then reveal the harder truth only after trust, leverage, or pressure changes hands.",
    ].join("\n");
  }

  if (mode === "assistant") {
    return generateAssistantFallbackAnswer(joined || cleanedInput);
  }

  if (mode === "quest") {
    return [
      "Title: Local Trouble on the Main Road",
      "Status: open",
      "Objective: Push the party toward the immediate threat with one obstacle and one consequence for delay.",
      "Giver: A pressured local contact",
      "Stakes: If ignored, the threat spreads and costs the party trust or safety.",
    ].join("\n");
  }

  if (mode === "location") {
    return [
      "Name: Rivergate Hamlet",
      "Hex: Frontier Route",
      "What Changed: Tension rose after a recent threat or disappearance tied to the main story.",
      "Notes: Use one sensory detail, one local problem, and one clue that points toward the next scene.",
    ].join("\n");
  }

  return lines.map((line) => sentenceCaseAndPunctuation(line)).join(" ");
}

function sanitizeAiTextOutput(rawText) {
  const lines = splitIntoIdeaLines(rawText);
  if (!lines.length) return "";
  const cleaned = lines
    .filter((line) => !isConstraintInstructionLine(line))
    .filter((line) => !isLikelyDuplicateLine(line, lines));
  return cleaned.join("\n").trim();
}

function isWeakNpcOutput(text) {
  const name = extractLabeledBlock(text, "Name");
  const role = extractLabeledBlock(text, "Role");
  const agenda = extractLabeledBlock(text, "Agenda");
  const disposition = extractLabeledBlock(text, "Disposition");
  const notes = buildNpcNotesFromAi(text);
  if (!name || !role || !agenda || !disposition || !notes) return true;

  const noteLines = splitIntoIdeaLines(notes);
  const bulletCount = noteLines.filter((line) => /^[-*]\s+/.test(line)).length;
  const noteChars = notes.replace(/\s+/g, " ").trim().length;
  if (noteChars < 110) return true;
  if (bulletCount > 0 && bulletCount < 4) return true;
  return false;
}

function isLikelyWeakAiOutput(text, mode, input, tabId) {
  const clean = str(text).trim();
  if (!clean) return true;
  if (/^\*{1,2}[^*\n]{1,60}$/.test(clean)) return true;
  if (/^[A-Za-z][A-Za-z ]{1,30}:?$/.test(clean) && clean.length < 40) return true;

  const lines = splitIntoIdeaLines(clean);
  if (lines.length === 1 && clean.length < 60 && !/[.!?]$/.test(clean)) return true;

  const lowerInput = str(input).toLowerCase();
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

  if ((str(mode).toLowerCase() === "npc" || tabId === "npcs") && isWeakNpcOutput(clean)) {
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

  if (str(mode).toLowerCase() !== "assistant" && clean.length < 24) return true;
  return false;
}

function processAiOutputWithFallback({ rawText, mode, input, source, tabId }) {
  const raw = str(rawText);
  const cleaned = sanitizeAiTextOutput(raw);
  const candidate = cleaned || raw;
  if (candidate && !isLikelyInstructionEcho(candidate) && !isLikelyWeakAiOutput(candidate, mode, input, tabId)) {
    return {
      text: candidate,
      usedFallback: false,
      source,
      mode,
      tabId,
    };
  }

  return {
    text: generateFallbackAiOutput({ mode, input, tabId }),
    usedFallback: true,
    source,
    mode,
    tabId,
  };
}

function generateFallbackAiOutput({ mode, input, tabId }) {
  const normalizedMode = str(mode).toLowerCase();
  const cleanInput = basicAutoCorrect(str(input));

  if (normalizedMode === "assistant") {
    return generateAssistantFallbackAnswer(cleanInput);
  }

  if (normalizedMode === "npc" || normalizedMode === "quest" || normalizedMode === "location") {
    return generateStructuredText(cleanInput, normalizedMode);
  }

  if (tabId) {
    const copilotFallback = generateCopilotFallbackByTab(tabId, cleanInput);
    if (copilotFallback) return copilotFallback;
  }

  return generateStructuredText(cleanInput, normalizedMode || "session");
}

function generateCopilotFallbackByTab(tabId, input) {
  const latest = getLatestSession();
  const cleanInput = str(input);
  if (tabId === "dashboard") {
    const lower = cleanInput.toLowerCase();
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
      "- 20m: prepare encounter or obstacle",
      "- 15m: prep clues/reveals",
      "- 10m: prep fallback scene",
    ].join("\n");
  }
  if (tabId === "sessions") {
    return [
      "Summary:",
      sentenceCaseAndPunctuation(cleanInput || latest?.summary || "Session notes captured and next objectives clarified"),
      "",
      "Next Prep:",
      "- Open with urgency tied to a current quest.",
      "- Prepare one social beat and one challenge beat.",
      "- End with a clear hook for the next session.",
    ].join("\n");
  }
  if (tabId === "capture") {
    return [
      "Summary:",
      sentenceCaseAndPunctuation(cleanInput || "Captured notes need grouping by scene and consequence"),
      "",
      "Follow-up Tasks:",
      "- Group notes by scene.",
      "- Mark unresolved hooks.",
      "- Push key entries into latest session log.",
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
      "What Changed: New signs of hostile activity were discovered nearby.",
      "Notes: Use drifting fog, damaged supplies, and witness rumors as scene cues.",
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
      "- Verify names/titles are final before import.",
      "- Import JSON pack and spot-check journal links.",
    ].join("\n");
  }
  if (tabId === "writing") {
    return generateAssistantFallbackAnswer(cleanInput);
  }
  return "";
}

function isConstraintInstructionLine(line) {
  const text = str(line).toLowerCase();
  if (!text) return false;
  if (/^(output rules|rules|constraints)\s*:/.test(text)) return true;
  if (/^\d+\)\s*/.test(text) && /(no|keep|return|do not)/.test(text)) return true;
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
function isLikelyInstructionEcho(text) {
  const lines = splitIntoIdeaLines(text);
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
    "keep the output",
    "single response",
    "output length",
    "output rules",
    "return plain text only",
  ].filter((token) => lower.includes(token)).length;
  return signalCount >= 2;
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
  return str(line)
    .toLowerCase()
    .replace(/^\d+\)\s*/, "")
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function generateAssistantFallbackAnswer(input) {
  const prompt = str(input);
  const lower = prompt.toLowerCase();
  if (!prompt) return "Give one clear GM question and I will generate practical table-ready options.";
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

  if (isSourceScopeQuestionPrompt(lower)) {
    return [
      "I only use your campaign data, the active kingdom rules profile if one is loaded, and PDFs indexed in this app.",
      "I do not have default access to external books.",
      "Open PDF Intel to index files, then ask what books are currently indexed.",
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
      "Quick NPC prompt:",
      "- Goal: what they want in this scene.",
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
    sentenceCaseAndPunctuation(prompt),
    "Turn this into one immediate scene objective, one obstacle, and one consequence.",
  ].join("\n");
}

function isSourceScopeQuestionPrompt(lowerPrompt) {
  const lower = str(lowerPrompt).toLowerCase().trim();
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

function lowercaseFirst(text) {
  const clean = str(text);
  if (!clean) return "";
  return clean.charAt(0).toLowerCase() + clean.slice(1);
}

function exportSessionPacketForLatest() {
  const latest = getLatestSession();
  if (!latest) {
    ui.sessionMessage = "No session found to export packet.";
    render();
    return;
  }
  exportSessionPacketForSession(latest.id);
}

function exportSessionPacketForSession(sessionId) {
  const session = state.sessions.find((s) => s.id === sessionId);
  if (!session) {
    ui.sessionMessage = "Could not find session for packet export.";
    render();
    return;
  }

  const mode = getPrepQueueMode();
  const markdown = generateSessionPacketMarkdown(session, mode);
  const safeTitle = slugify(session.title || "session-packet");
  const filename = `${safeTitle}-next-session-packet-${dateStamp()}.md`;
  downloadText(markdown, filename, "text/markdown");
  ui.sessionMessage = `Exported session packet: ${filename}`;
  render();
}

function generateSessionPacketMarkdown(session, mode) {
  const sourceText = `${session.summary || ""} ${session.nextPrep || ""}`;
  const queue = generatePrepQueue(mode);
  const checklist = generateSmartChecklist().slice(0, 8);
  const openQuests = state.quests
    .filter((q) => q.status !== "completed" && q.status !== "failed")
    .sort((a, b) => relevanceScore(sourceText, b.title) - relevanceScore(sourceText, a.title))
    .slice(0, 6);
  const npcFocus = getMentionedOrRecent(state.npcs, "name", sourceText, 4);
  const locationFocus = getMentionedOrRecent(state.locations, "name", sourceText, 3);
  const sceneOpeners = getSceneOpenersForSessionPacket(session);
  const wrapBullets = getSmartWrapBulletsForSessionPacket(session);

  return `# Next Session Packet - ${session.title}

Generated: ${new Date().toLocaleString()}
Prep Mode: ${mode} minutes

## 1) Read-Aloud Recap
- ${condenseLine(session.summary) || "No recap captured yet."}

## 2) Opening Scene Options
${sceneOpeners.length ? sceneOpeners.map((line, i) => `${i + 1}. ${line}`).join("\n") : "- Create one cold open using the top open quest + top NPC."}

## 3) Smart Wrap-Up Priorities
${wrapBullets.length ? wrapBullets.map((line) => `- ${line}`).join("\n") : "- Run Smart Wrap-Up in the app for generated priorities."}

## 4) Time-Boxed Prep Queue (${mode}m)
${queue.map((task) => `- [ ] (${task.minutes}m) ${task.label}`).join("\n")}

## 5) Session Start Checklist
${checklist.map((item) => `- [ ] ${item.label}`).join("\n")}

## 6) Open Quests To Push
${openQuests.length ? openQuests.map((q) => `- ${q.title} (${q.status})`).join("\n") : "- None"}

## 7) NPC Focus Cards
${npcFocus.length ? npcFocus
    .map((npc) => `- ${npc.name}: role=${npc.role || "n/a"}, agenda=${npc.agenda || "n/a"}`)
    .join("\n") : "- None"}

## 8) Location Focus
${locationFocus.length ? locationFocus
    .map((loc) => `- ${loc.name}: ${loc.whatChanged || "No recent change logged."}`)
    .join("\n") : "- None"}

## 9) Foundry Handoff
- [ ] Import/export any updated NPC actors.
- [ ] Import/update quest + location journals.
- [ ] Confirm opening scene map, walls, and tokens.

## 10) PDF Intel Checks
${(state.meta.pdfIndexedCount || 0) > 0
    ? `- [ ] Run targeted search terms: ${suggestSearchTerms(sourceText, 4).join(", ") || "rules, travel, hazards"}`
    : "- [ ] Index PDFs first in PDF Intel tab."}
`;
}

function getSmartWrapBulletsForSessionPacket(session) {
  const text = String(session.nextPrep || "");
  const markerRegex = /<!-- SMART_WRAPUP_START -->[\s\S]*?<!-- SMART_WRAPUP_END -->/m;
  const match = text.match(markerRegex);
  if (!match) return [];
  return splitBulletLines(match[0]).slice(0, 6);
}

function getSceneOpenersForSessionPacket(session) {
  const text = String(session.nextPrep || "");
  const markerRegex = /<!-- SMART_SCENES_START -->[\s\S]*?<!-- SMART_SCENES_END -->/m;
  const match = text.match(markerRegex);
  if (match) {
    const numbered = splitNumberedLines(match[0]).slice(0, 3);
    if (numbered.length) return numbered;
  }
  return generateSceneOpeners(session, null);
}

function splitBulletLines(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.startsWith("- "))
    .map((line) => line.replace(/^- /, "").trim());
}

function splitNumberedLines(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => /^\d+\.\s+/.test(line))
    .map((line) => line.replace(/^\d+\.\s+/, "").trim());
}

function createCaptureEntry(kind, note, sessionId) {
  let text = str(note);
  if (!text) {
    const prompted = prompt(`Quick ${kind} note:`) || "";
    text = str(prompted);
  }
  if (!text) return;

  state.liveCapture.unshift({
    id: uid(),
    kind: str(kind) || "Note",
    note: text,
    sessionId: str(sessionId),
    timestamp: new Date().toISOString(),
  });
  saveState();

  ui.captureDraft.note = "";
  ui.captureMessage = `Captured ${str(kind) || "Note"} entry.`;
  render();
}

function appendCaptureToSession() {
  const targetSessionId = getResolvedCaptureSessionId();
  const session = state.sessions.find((s) => s.id === targetSessionId);
  if (!session) {
    ui.captureMessage = "No target session found for append.";
    render();
    return;
  }

  const relevant = (state.liveCapture || [])
    .filter((entry) => !entry.sessionId || entry.sessionId === session.id)
    .slice(0, 20);

  if (!relevant.length) {
    ui.captureMessage = "No capture entries available to append.";
    render();
    return;
  }

  const lines = relevant.map(
    (entry) =>
      `- [${entry.kind}] ${new Date(entry.timestamp || Date.now()).toLocaleTimeString()} - ${entry.note}`
  );
  const section = `<!-- LIVE_CAPTURE_START -->
### Live Capture Log
${lines.join("\n")}
<!-- LIVE_CAPTURE_END -->`;

  session.summary = injectOrReplaceLiveCaptureSection(session.summary || "", section);
  session.updatedAt = new Date().toISOString();
  saveState();
  ui.captureMessage = `Appended ${relevant.length} capture entries to "${session.title}".`;
  render();
}

function getResolvedCaptureSessionId() {
  const chosen = str(ui.captureDraft.sessionId);
  if (chosen) return chosen;
  return getLatestSession()?.id || "";
}

function injectOrReplaceLiveCaptureSection(currentText, section) {
  const markerRegex = /<!-- LIVE_CAPTURE_START -->[\s\S]*?<!-- LIVE_CAPTURE_END -->/m;
  if (markerRegex.test(currentText)) {
    return currentText.replace(markerRegex, section).trim();
  }
  const base = str(currentText);
  return base ? `${base}\n\n${section}` : section;
}

function hasKingdomSignals(text) {
  return /\b(kingdom|unrest|bp|build|claim|hex|edict|upkeep)\b/i.test(text || "");
}

function getMentionedOrRecent(collection, nameField, sourceText, limit) {
  const textLower = String(sourceText || "").toLowerCase();
  const withName = collection
    .filter((item) => str(item[nameField]))
    .map((item) => ({
      item,
      mentioned: textLower.includes(String(item[nameField]).toLowerCase()),
      updatedKey: Date.parse(item.updatedAt || item.createdAt || "") || 0,
    }))
    .sort((a, b) => {
      if (a.mentioned !== b.mentioned) return a.mentioned ? -1 : 1;
      return b.updatedKey - a.updatedKey;
    })
    .slice(0, limit)
    .map((entry) => entry.item);

  return withName;
}

function relevanceScore(sourceText, title) {
  const text = String(sourceText || "").toLowerCase();
  const phrase = String(title || "").toLowerCase();
  if (!phrase) return 0;

  let score = text.includes(phrase) ? 6 : 0;
  const words = phrase.split(/[^a-z0-9]+/i).filter((w) => w.length >= 4);
  for (const word of words) {
    if (text.includes(word)) score += 1;
  }
  return score;
}

function suggestSearchTerms(text, limit) {
  const stopWords = new Set([
    "about", "after", "again", "against", "along", "because", "before", "between", "during",
    "first", "foundry", "their", "there", "these", "those", "through", "under", "while",
    "where", "which", "would", "could", "should", "party", "session", "notes", "quest", "next",
    "campaign", "setting", "module",
  ]);

  const words = String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((word) => word.length >= 5 && !stopWords.has(word));

  const counts = new Map();
  for (const word of words) {
    counts.set(word, (counts.get(word) || 0) + 1);
  }

  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, limit)
    .map((entry) => entry[0]);
}

function slugify(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "")
    .slice(0, 60);
}

async function handleFormSubmit(type, form) {
  const fields = Object.fromEntries(new FormData(form).entries());

  if (type === "sessions") {
    state.sessions.unshift({
      id: uid(),
      title: str(fields.title),
      date: str(fields.date),
      arc: str(fields.arc),
      kingdomTurn: str(fields.kingdomTurn),
      summary: str(fields.summary),
      nextPrep: str(fields.nextPrep),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    saveState();
    form.reset();
    render();
    return;
  }

  if (type === "kingdom-overview") {
    applyKingdomOverviewForm(fields);
    saveState();
    ui.kingdomMessage = "Kingdom overview updated.";
    render();
    return;
  }

  if (type === "kingdom-leader") {
    createKingdomLeader(fields);
    saveState();
    form.reset();
    ui.kingdomMessage = "Kingdom leader added.";
    render();
    return;
  }

  if (type === "kingdom-settlement") {
    createKingdomSettlement(fields);
    saveState();
    form.reset();
    ui.kingdomMessage = "Settlement added.";
    render();
    return;
  }

  if (type === "kingdom-region") {
    createKingdomRegion(fields);
    saveState();
    form.reset();
    ui.kingdomMessage = "Region record added.";
    render();
    return;
  }

  if (type === "kingdom-turn") {
    applyKingdomTurnForm(fields);
    saveState();
    form.reset();
    ui.kingdomMessage = "Kingdom turn applied and recorded.";
    render();
    return;
  }

  if (type === "npcs") {
    const id = uid();
    const folder = normalizeWorldFolderName(fields.folder);
    if (folder) addWorldFolder("npcs", folder);
    state.npcs.unshift({
      id,
      name: str(fields.name),
      role: str(fields.role),
      agenda: str(fields.agenda),
      disposition: str(fields.disposition),
      notes: str(fields.notes),
      folder,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    ui.worldSelection.npcs = id;
    ui.worldNewFolder.npcs = folder;
    saveState();
    form.reset();
    render();
    return;
  }

  if (type === "quests") {
    const id = uid();
    const folder = normalizeWorldFolderName(fields.folder);
    if (folder) addWorldFolder("quests", folder);
    state.quests.unshift({
      id,
      title: str(fields.title),
      status: str(fields.status) || "open",
      objective: str(fields.objective),
      giver: str(fields.giver),
      stakes: str(fields.stakes),
      folder,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    ui.worldSelection.quests = id;
    ui.worldNewFolder.quests = folder;
    saveState();
    form.reset();
    render();
    return;
  }

  if (type === "locations") {
    const id = uid();
    const folder = normalizeWorldFolderName(fields.folder);
    if (folder) addWorldFolder("locations", folder);
    state.locations.unshift({
      id,
      name: str(fields.name),
      hex: str(fields.hex),
      whatChanged: str(fields.whatChanged),
      notes: str(fields.notes),
      folder,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
    ui.worldSelection.locations = id;
    ui.worldNewFolder.locations = folder;
    saveState();
    form.reset();
    render();
    return;
  }

  if (type === "session-close-wizard") {
    const sessionId = str(fields.sessionId) || ui.wizardDraft.sessionId;
    const session = state.sessions.find((s) => s.id === sessionId);
    if (!session) {
      ui.sessionMessage = "Wizard failed: target session not found.";
      ui.wizardOpen = false;
      render();
      return;
    }

    const wizardAnswers = {
      highlights: str(fields.highlights),
      cliffhanger: str(fields.cliffhanger),
      playerIntent: str(fields.playerIntent),
    };
    const sceneOpeners = generateSceneOpeners(session, wizardAnswers);
    const result = generateWrapUpForSession(sessionId, {
      wizardAnswers,
      sceneOpeners,
      silent: true,
    });

    ui.wizardOpen = false;
    ui.wizardDraft = {
      sessionId: "",
      highlights: "",
      cliffhanger: "",
      playerIntent: "",
    };

    if (result) {
      ui.sessionMessage = `Wizard complete for "${result.session.title}": wrap-up + ${sceneOpeners.length} scene openers added.`;
    } else {
      ui.sessionMessage = "Wizard ran, but no session was updated.";
    }
    render();
    return;
  }

  if (type === "pdf-search") {
    if (!desktopApi) return;
    const query = str(fields.query);
    const limit = Number.parseInt(str(fields.limit), 10) || 20;
    await runPdfSearch(query, limit);
  }
}

function getEntityCollectionRef(collection) {
  const clean = str(collection);
  if (!clean) return null;
  if (clean === "kingdomLeaders") return getKingdomState().leaders;
  if (clean === "kingdomSettlements") return getKingdomState().settlements;
  if (clean === "kingdomRegions") return getKingdomState().regions;
  if (clean === "kingdomTurns") return getKingdomState().turns;
  return Array.isArray(state[clean]) ? state[clean] : null;
}

function normalizeEntityPatch(collection, patch) {
  const cleanCollection = str(collection);
  const out = {};
  for (const [field, value] of Object.entries(patch || {})) {
    if (cleanCollection === "kingdomLeaders" && field === "leadershipBonus") {
      out[field] = Math.max(0, Math.min(4, Number.parseInt(String(value || "0"), 10) || 0));
      continue;
    }
    if (cleanCollection === "kingdomSettlements" && ["influence", "resourceDice", "consumption"].includes(field)) {
      out[field] = Math.max(0, Number.parseInt(String(value || "0"), 10) || 0);
      continue;
    }
    out[field] = value;
  }
  return out;
}

function deleteEntity(collection, id) {
  if (!confirm("Delete this entry?")) return;
  const group = getEntityCollectionRef(collection);
  if (!Array.isArray(group)) return;
  const index = group.findIndex((item) => item.id === id);
  if (index < 0) return;
  group.splice(index, 1);
  if (ui.worldSelection && collection in ui.worldSelection && Array.isArray(state[collection])) {
    if (ui.worldSelection[collection] === id) {
      ui.worldSelection[collection] = state[collection][0]?.id || "";
    }
  }
  saveState();
  render();
}

function patchEntity(collection, id, patch) {
  const group = getEntityCollectionRef(collection);
  if (!Array.isArray(group)) return;
  const item = group.find((entry) => entry.id === id);
  if (!item) return;
  Object.assign(item, normalizeEntityPatch(collection, patch), { updatedAt: new Date().toISOString() });
  saveState();
}

function getPdfSummaryMap() {
  if (!state.meta.pdfSummaries || typeof state.meta.pdfSummaries !== "object" || Array.isArray(state.meta.pdfSummaries)) {
    state.meta.pdfSummaries = {};
  }
  return state.meta.pdfSummaries;
}

function syncPdfSummarySelection() {
  const files = Array.isArray(state?.meta?.pdfIndexedFiles)
    ? state.meta.pdfIndexedFiles.map((name) => str(name)).filter(Boolean)
    : [];
  if (!files.length) {
    ui.pdfSummaryFile = "";
    if (!ui.pdfSummaryBusy) ui.pdfSummaryOutput = "";
    if (!ui.pdfSummaryBusy) resetPdfSummaryProgress();
    return;
  }
  if (!str(ui.pdfSummaryFile) || !files.includes(str(ui.pdfSummaryFile))) {
    ui.pdfSummaryFile = files[0];
  }
}

function getPdfSummaryByFileName(fileName) {
  const name = str(fileName);
  if (!name) return null;
  const summaries = getPdfSummaryMap();
  for (const value of Object.values(summaries)) {
    if (!value || typeof value !== "object") continue;
    if (str(value.fileName) === name) return value;
  }
  return null;
}

function upsertPdfSummary(summaryResult) {
  const fileName = str(summaryResult?.fileName);
  const filePath = str(summaryResult?.path);
  const summary = str(summaryResult?.summary).slice(0, 24000);
  if (!fileName || !summary) return;
  const key = filePath || fileName;
  const summaries = getPdfSummaryMap();
  summaries[key] = {
    fileName,
    path: filePath,
    summary,
    updatedAt: str(summaryResult?.summaryUpdatedAt) || new Date().toISOString(),
  };
  state.meta.pdfSummaries = summaries;
}

function resetPdfSummaryProgress() {
  ui.pdfSummaryProgressCurrent = 0;
  ui.pdfSummaryProgressTotal = 0;
  ui.pdfSummaryProgressLabel = "";
}

function applyPdfSummarizeProgress(payload) {
  const fileName = str(payload?.fileName);
  const selectedFile = str(ui.pdfSummaryFile);
  if (fileName && selectedFile && fileName !== selectedFile) return;

  const totalRaw = Number.parseInt(String(payload?.total || "0"), 10);
  const currentRaw = Number.parseInt(String(payload?.current || "0"), 10);
  const total = Number.isFinite(totalRaw) ? Math.max(1, totalRaw) : 1;
  const current = Number.isFinite(currentRaw) ? Math.max(0, Math.min(currentRaw, total)) : 0;

  ui.pdfSummaryProgressTotal = total;
  ui.pdfSummaryProgressCurrent = current;
  const msg = str(payload?.message);
  if (msg) {
    ui.pdfSummaryProgressLabel = msg;
    ui.pdfMessage = msg;
  }

  if (str(payload?.stage) === "done") {
    ui.pdfSummaryProgressCurrent = ui.pdfSummaryProgressTotal || 1;
  }

  if (activeTab === "pdf") render();
}

async function summarizeSelectedPdf(force = false) {
  if (!desktopApi?.summarizePdfFile) {
    ui.pdfMessage = "PDF summary requires the desktop runtime bridge.";
    render();
    return;
  }
  const fileName = str(ui.pdfSummaryFile);
  if (!fileName) {
    ui.pdfMessage = "Choose an indexed PDF first.";
    render();
    return;
  }

  syncPdfSummarySelection();
  resetPdfSummaryProgress();
  ui.pdfSummaryProgressTotal = 1;
  ui.pdfSummaryProgressCurrent = 0;
  ui.pdfSummaryProgressLabel = force
    ? `Refreshing summary for ${fileName}...`
    : `Summarizing ${fileName}...`;
  ui.pdfSummaryBusy = true;
  ui.pdfMessage = force
    ? `Refreshing summary for ${fileName}...`
    : `Summarizing ${fileName}...`;
  render();
  try {
    const config = ensureAiConfig();
    const result = await desktopApi.summarizePdfFile({
      fileName,
      force: force === true,
      config,
    });
    upsertPdfSummary(result);
    ui.pdfSummaryOutput = str(result?.summary);
    state.meta.pdfIndexedAt = str(state.meta.pdfIndexedAt) || new Date().toISOString();
    saveState();
    ui.pdfMessage = result?.reused
      ? `Loaded saved summary for ${fileName}.`
      : `Summary generated for ${fileName} (${Number.parseInt(String(result?.chunks || 0), 10) || 0} chunk(s)).`;
    if (!result?.reused && !ui.pdfSummaryProgressTotal) {
      const chunks = Math.max(1, Number.parseInt(String(result?.chunks || "1"), 10) || 1);
      ui.pdfSummaryProgressTotal = chunks + 1;
      ui.pdfSummaryProgressCurrent = chunks + 1;
    }
    if (result?.reused) {
      ui.pdfSummaryProgressTotal = 1;
      ui.pdfSummaryProgressCurrent = 1;
    }
    ui.pdfSummaryProgressLabel = ui.pdfMessage;
  } catch (err) {
    ui.pdfMessage = `Summary failed: ${readableError(err)}`;
    ui.pdfSummaryProgressLabel = ui.pdfMessage;
  } finally {
    ui.pdfSummaryBusy = false;
    render();
  }
}

async function choosePdfFolder() {
  if (!desktopApi) return;
  try {
    const selected = await desktopApi.pickPdfFolder();
    if (!selected) return;
    state.meta.pdfFolder = selected;
    saveState();
    ui.pdfMessage = "PDF folder updated.";
    render();
  } catch (err) {
    ui.pdfMessage = `Failed to choose folder: ${String(err)}`;
    render();
  }
}

async function indexPdfLibrary() {
  if (!desktopApi) return;
  const folderPath = str(state.meta.pdfFolder);
  if (!folderPath) {
    ui.pdfMessage = "Set a PDF folder first.";
    render();
    return;
  }

  ui.pdfBusy = true;
  ui.pdfMessage = "Indexing PDFs. This can take a bit on first run...";
  render();
  try {
    const summary = await desktopApi.indexPdfFolder(folderPath);
    state.meta.pdfIndexedAt = summary.indexedAt || new Date().toISOString();
    state.meta.pdfIndexedCount = summary.count || 0;
    state.meta.pdfIndexedFiles = Array.isArray(summary?.fileNames)
      ? summary.fileNames.map((name) => str(name)).filter(Boolean).sort((a, b) => a.localeCompare(b))
      : [];
    const files = Array.isArray(summary?.files) ? summary.files : [];
    if (files.length) {
      const summaries = getPdfSummaryMap();
      for (const file of files) {
        const fileName = str(file?.fileName);
        const filePath = str(file?.path);
        const key = filePath || fileName;
        if (!key) continue;
        const text = str(file?.summary);
        if (!text) continue;
        summaries[key] = {
          fileName: fileName || key,
          path: filePath,
          summary: text.slice(0, 24000),
          updatedAt: str(file?.summaryUpdatedAt) || str(summary?.indexedAt) || "",
        };
      }
      state.meta.pdfSummaries = summaries;
    }
    syncPdfSummarySelection();
    if (str(ui.pdfSummaryFile)) {
      const existing = getPdfSummaryByFileName(ui.pdfSummaryFile);
      ui.pdfSummaryOutput = str(existing?.summary);
    } else {
      ui.pdfSummaryOutput = "";
    }
    saveState();
    ui.pdfMessage = `Indexed ${summary.count || 0} file(s). Failed: ${summary.failed || 0}.`;
    ui.pdfSearchResults = [];
  } catch (err) {
    ui.pdfMessage = `Index failed: ${String(err)}`;
  } finally {
    ui.pdfBusy = false;
    render();
  }
}

async function runPdfSearch(query, limit = 20) {
  if (!desktopApi) return;
  const normalizedQuery = str(query);
  const normalizedLimit = Number.parseInt(String(limit || "20"), 10) || 20;
  const aiConfig = ensureAiConfig();
  ui.pdfSearchQuery = normalizedQuery;

  if (!normalizedQuery) {
    ui.pdfSearchResults = [];
    ui.pdfMessage = "Enter a search query.";
    render();
    return;
  }

  ui.pdfBusy = true;
  ui.pdfMessage = "Searching indexed PDFs...";
  render();
  try {
    const result = await desktopApi.searchPdf({ query: normalizedQuery, limit: normalizedLimit, config: aiConfig });
    ui.pdfSearchResults = Array.isArray(result.results) ? result.results : [];
    const retrievalMode = str(result?.retrieval?.mode || "");
    const embeddingModel = str(result?.retrieval?.embeddingModel || "");
    const retrievalNote = str(result?.retrieval?.note || "");
    if (ui.pdfSearchResults.length) {
      if (retrievalMode === "hybrid" && embeddingModel) {
        ui.pdfMessage = `Found ${ui.pdfSearchResults.length} result(s) using hybrid search with ${embeddingModel}.`;
      } else if (retrievalMode === "semantic" && embeddingModel) {
        ui.pdfMessage = `Found ${ui.pdfSearchResults.length} result(s) using semantic search with ${embeddingModel}.`;
      } else if (retrievalMode === "lexical" && retrievalNote) {
        ui.pdfMessage = `Found ${ui.pdfSearchResults.length} result(s). ${retrievalNote}`;
      } else {
        ui.pdfMessage = `Found ${ui.pdfSearchResults.length} result(s).`;
      }
    } else {
      ui.pdfMessage = retrievalNote || "No PDF matches found.";
    }
  } catch (err) {
    ui.pdfMessage = `Search failed: ${String(err)}`;
    ui.pdfSearchResults = [];
  } finally {
    ui.pdfBusy = false;
    render();
  }
}

function exportFoundry(kind) {
  const ts = dateStamp();
  if (kind === "npcs") {
    const actors = state.npcs.map(toFoundryActor);
    return downloadJson(actors, `dm-helper-npcs-foundry-${ts}.json`);
  }

  if (kind === "quests") {
    const quests = state.quests.map((q) => toFoundryJournal(q, "quest"));
    return downloadJson(quests, `dm-helper-quests-foundry-${ts}.json`);
  }

  if (kind === "locations") {
    const locations = state.locations.map((l) => toFoundryJournal(l, "location"));
    return downloadJson(locations, `dm-helper-locations-foundry-${ts}.json`);
  }

  const all = [
    ...state.npcs.map(toFoundryActor),
    ...state.quests.map((q) => toFoundryJournal(q, "quest")),
    ...state.locations.map((l) => toFoundryJournal(l, "location")),
  ];
  downloadJson(all, `dm-helper-full-foundry-pack-${ts}.json`);
}

function toFoundryActor(npc) {
  return {
    _id: foundryId(),
    name: npc.name,
    type: "npc",
    img: "icons/svg/mystery-man.svg",
    system: {},
    prototypeToken: {},
    flags: {
      dmHelper: {
        source: "dm-helper-desktop",
        exportType: "npc",
        exportDate: new Date().toISOString(),
      },
    },
    items: [],
    effects: [],
    folder: null,
    sort: 0,
    ownership: { default: 0 },
    _stats: { systemId: "pf2e", coreVersion: "11" },
    biography: {
      role: npc.role || "",
      agenda: npc.agenda || "",
      disposition: npc.disposition || "",
      notes: npc.notes || "",
    },
  };
}

function toFoundryJournal(entry, type) {
  const title = type === "quest" ? entry.title : entry.name;
  const body =
    type === "quest"
      ? `<h2>Objective</h2><p>${escapeHtml(entry.objective || "")}</p><h2>Stakes</h2><p>${escapeHtml(
          entry.stakes || ""
        )}</p>`
      : `<h2>What Changed</h2><p>${escapeHtml(entry.whatChanged || "")}</p><h2>Notes</h2><p>${escapeHtml(
          entry.notes || ""
        )}</p>`;

  return {
    _id: foundryId(),
    name: title,
    pages: [
      {
        _id: foundryId(),
        name: title,
        type: "text",
        text: { content: body, format: 1 },
      },
    ],
    flags: {
      dmHelper: {
        source: "dm-helper-desktop",
        exportType: type,
        exportDate: new Date().toISOString(),
      },
    },
    folder: null,
    sort: 0,
    ownership: { default: 0 },
    _stats: { systemId: "pf2e", coreVersion: "11" },
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return createStarterState();
    return normalizeState(JSON.parse(raw));
  } catch {
    return createStarterState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function normalizeState(input) {
  const base = createStarterState();
  const out = {
    ...base,
    ...input,
  };
  out.meta = { ...base.meta, ...(out.meta || {}) };
  if (
    !out.meta.checklistChecks ||
    typeof out.meta.checklistChecks !== "object" ||
    Array.isArray(out.meta.checklistChecks)
  ) {
    out.meta.checklistChecks = {};
  }
  if (
    !out.meta.prepQueueChecks ||
    typeof out.meta.prepQueueChecks !== "object" ||
    Array.isArray(out.meta.prepQueueChecks)
  ) {
    out.meta.prepQueueChecks = {};
  }
  if (!Array.isArray(out.meta.customChecklistItems)) {
    out.meta.customChecklistItems = [];
  }
  if (
    !out.meta.checklistOverrides ||
    typeof out.meta.checklistOverrides !== "object" ||
    Array.isArray(out.meta.checklistOverrides)
  ) {
    out.meta.checklistOverrides = {};
  }
  if (
    !out.meta.checklistArchived ||
    typeof out.meta.checklistArchived !== "object" ||
    Array.isArray(out.meta.checklistArchived)
  ) {
    out.meta.checklistArchived = {};
  }
  const mode = Number.parseInt(String(out.meta.prepQueueMode || "60"), 10);
  out.meta.prepQueueMode = mode === 30 || mode === 90 ? mode : 60;
  if (!Array.isArray(out.meta.pdfIndexedFiles)) {
    out.meta.pdfIndexedFiles = [];
  } else {
    out.meta.pdfIndexedFiles = out.meta.pdfIndexedFiles
      .map((name) => str(name))
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b))
      .slice(0, 400);
  }
  if (!out.meta.pdfSummaries || typeof out.meta.pdfSummaries !== "object" || Array.isArray(out.meta.pdfSummaries)) {
    out.meta.pdfSummaries = {};
  } else {
    const cleanSummaries = {};
    for (const [key, value] of Object.entries(out.meta.pdfSummaries)) {
      const fileKey = str(key);
      if (!fileKey) continue;
      if (value && typeof value === "object") {
        const fileName = str(value.fileName) || fileKey;
        const path = str(value.path);
        const summary = str(value.summary).slice(0, 24000);
        if (!summary) continue;
        cleanSummaries[fileKey] = {
          fileName,
          path,
          summary,
          updatedAt: str(value.updatedAt) || "",
        };
      } else {
        const summary = str(value).slice(0, 24000);
        if (!summary) continue;
        cleanSummaries[fileKey] = {
          fileName: fileKey,
          path: "",
          summary,
          updatedAt: "",
        };
      }
    }
    out.meta.pdfSummaries = cleanSummaries;
  }
  out.meta.aiConfig =
    out.meta.aiConfig && typeof out.meta.aiConfig === "object" && !Array.isArray(out.meta.aiConfig)
      ? {
          endpoint: str(out.meta.aiConfig.endpoint) || "http://127.0.0.1:11434",
          model: str(out.meta.aiConfig.model) || "llama3.1:8b",
          temperature: Number.isFinite(Number.parseFloat(String(out.meta.aiConfig.temperature)))
            ? Math.max(0, Math.min(Number.parseFloat(String(out.meta.aiConfig.temperature)), 2))
            : 0.2,
          maxOutputTokens: Number.isFinite(Number.parseInt(String(out.meta.aiConfig.maxOutputTokens), 10))
            ? Math.max(64, Math.min(Number.parseInt(String(out.meta.aiConfig.maxOutputTokens), 10), 2048))
            : 320,
          timeoutSec: Number.isFinite(Number.parseInt(String(out.meta.aiConfig.timeoutSec), 10))
            ? Math.max(15, Math.min(Number.parseInt(String(out.meta.aiConfig.timeoutSec), 10), 1200))
            : 120,
          compactContext: out.meta.aiConfig.compactContext === false ? false : true,
          autoRunTabs: out.meta.aiConfig.autoRunTabs === false ? false : true,
          usePdfContext: out.meta.aiConfig.usePdfContext === false ? false : true,
          aiProfile: ["fast", "deep", "custom"].includes(str(out.meta.aiConfig.aiProfile).toLowerCase())
            ? str(out.meta.aiConfig.aiProfile).toLowerCase()
            : "fast",
        }
      : {
          endpoint: "http://127.0.0.1:11434",
          model: "llama3.1:8b",
          temperature: 0.2,
          maxOutputTokens: 320,
          timeoutSec: 120,
          compactContext: true,
          autoRunTabs: true,
          usePdfContext: true,
          aiProfile: "fast",
        };
  if (!Array.isArray(out.meta.aiHistory)) {
    out.meta.aiHistory = [];
  } else {
    out.meta.aiHistory = out.meta.aiHistory
      .map((entry) => ({
        id: str(entry?.id) || `ai-turn-${uid()}`,
        tabId: str(entry?.tabId) || "dashboard",
        role: str(entry?.role).toLowerCase() === "assistant" ? "assistant" : "user",
        mode: str(entry?.mode) || "assistant",
        text: str(entry?.text).replace(/\s+/g, " ").slice(0, 1800),
        at: str(entry?.at) || new Date().toISOString(),
      }))
      .filter((entry) => entry.text)
      .slice(-AI_HISTORY_LIMIT);
  }
  if (!out.meta.worldFolders || typeof out.meta.worldFolders !== "object" || Array.isArray(out.meta.worldFolders)) {
    out.meta.worldFolders = { npcs: [], quests: [], locations: [] };
  } else {
    out.meta.worldFolders = {
      npcs: Array.isArray(out.meta.worldFolders.npcs) ? out.meta.worldFolders.npcs : [],
      quests: Array.isArray(out.meta.worldFolders.quests) ? out.meta.worldFolders.quests : [],
      locations: Array.isArray(out.meta.worldFolders.locations) ? out.meta.worldFolders.locations : [],
    };
  }
  out.kingdom = normalizeKingdomState(out.kingdom);
  out.sessions = Array.isArray(out.sessions) ? out.sessions : [];
  out.npcs = Array.isArray(out.npcs) ? out.npcs : [];
  out.quests = Array.isArray(out.quests) ? out.quests : [];
  out.locations = Array.isArray(out.locations) ? out.locations : [];
  out.npcs = out.npcs.map((item) => ({ ...item, folder: normalizeWorldFolderName(item?.folder) }));
  out.quests = out.quests.map((item) => ({ ...item, folder: normalizeWorldFolderName(item?.folder) }));
  out.locations = out.locations.map((item) => ({ ...item, folder: normalizeWorldFolderName(item?.folder) }));
  out.liveCapture = Array.isArray(out.liveCapture) ? out.liveCapture : [];
  return out;
}

function createStarterState() {
  return {
    meta: {
      campaignName: "My Campaign",
      createdAt: new Date().toISOString(),
      pdfFolder: "",
      pdfIndexedAt: "",
      pdfIndexedCount: 0,
      pdfIndexedFiles: [],
      pdfSummaries: {},
      checklistChecks: {},
      customChecklistItems: [],
      checklistOverrides: {},
      checklistArchived: {},
      prepQueueMode: 60,
      prepQueueChecks: {},
      aiConfig: {
        endpoint: "http://127.0.0.1:11434",
        model: "llama3.1:8b",
        temperature: 0.2,
        maxOutputTokens: 320,
        timeoutSec: 120,
        compactContext: true,
        autoRunTabs: true,
        usePdfContext: true,
        aiProfile: "fast",
      },
      aiHistory: [],
      worldFolders: {
        npcs: ["Capital", "Frontier"],
        quests: ["Main Arc", "Waystation"],
        locations: ["Capital", "Frontier"],
      },
    },
    sessions: [
      {
        id: uid(),
        title: "Session 00 - Campaign Kickoff",
        date: "",
        arc: "Frontier Arc",
        kingdomTurn: "",
        summary: "Patron briefing complete. Party goals aligned. Travel plan set for the border waystation.",
        nextPrep: "Prepare patron + quartermaster scenes, one road encounter, and first rumor threads.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    ],
    npcs: [
      {
        id: uid(),
        name: "Lady Ardyn Vale",
        role: "Noble patron",
        agenda: "Stabilize the frontier under a reliable chartered force.",
        disposition: "Allied",
        folder: "Capital",
        notes: "Direct, formal, and pragmatic.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
      {
        id: uid(),
        name: "Quartermaster Bren",
        role: "Waystation owner",
        agenda: "Remove bandit pressure and keep trade routes open.",
        disposition: "Cautiously allied",
        folder: "Frontier",
        notes: "Values practical help over promises.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    ],
    quests: [
      {
        id: uid(),
        title: "Establish a Foothold in the Frontier",
        status: "open",
        objective: "Secure local allies, map immediate threat zones, and establish a safe operational base.",
        giver: "Lady Ardyn Vale",
        folder: "Main Arc",
        stakes: "Without early momentum, rivals and bandits control the frontier narrative.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
      {
        id: uid(),
        title: "Bandit Pressure at the Waystation",
        status: "open",
        objective: "Identify and reduce active bandit operations near Blackbridge Waystation.",
        giver: "Quartermaster Bren",
        folder: "Waystation",
        stakes: "Trade and trust collapse if raids continue.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    ],
    locations: [
      {
        id: uid(),
        name: "Blackbridge Waystation",
        hex: "Frontier Route",
        folder: "Frontier",
        whatChanged: "Declared as primary safe base for early expedition stages.",
        notes: "Anchor location for rumors, supplies, and consequence callbacks.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
      {
        id: uid(),
        name: "Council Hall",
        hex: "Capital District",
        folder: "Capital",
        whatChanged: "Charter issued and authority delegated to party.",
        notes: "Use for political updates and mission reframing.",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      },
    ],
    kingdom: createStarterKingdomState(),
    liveCapture: [],
  };
}

function downloadJson(payload, filename) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  downloadBlob(blob, filename);
}

function downloadText(content, filename, mimeType = "text/plain") {
  const blob = new Blob([content], { type: mimeType });
  downloadBlob(blob, filename);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function uid() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function foundryId() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let id = "";
  for (let i = 0; i < 16; i += 1) id += chars[Math.floor(Math.random() * chars.length)];
  return id;
}

function str(value) {
  return String(value || "").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function safeDate(value) {
  const ms = Date.parse(value || "");
  return Number.isNaN(ms) ? 0 : ms;
}

function dateStamp() {
  return new Date().toISOString().slice(0, 10);
}

function readableError(err) {
  if (!err) return "Unknown error.";
  const raw =
    typeof err === "string"
      ? err
      : typeof err.message === "string" && err.message.trim()
        ? err.message.trim()
        : String(err);
  return raw
    .replace(/^Error invoking remote method '[^']+':\s*/i, "")
    .replace(/^Error:\s*/i, "")
    .trim();
}

function recordAiError(action, err) {
  const clean = readableError(err);
  ui.aiLastError = `${action}: ${clean}`;
  ui.aiLastErrorAt = new Date().toISOString();
  return clean;
}

function clearAiError() {
  ui.aiLastError = "";
  ui.aiLastErrorAt = "";
}

function formatAiErrorHint(message) {
  const text = String(message || "").toLowerCase();
  if (!text) return [];

  if (text.includes("timed out")) {
    return [
      "Increase Timeout or reduce Max Output Tokens.",
      "Turn on Compact context mode.",
      "Try a faster model (for example, a smaller local model).",
      "Disable Auto-run on tab switch to avoid surprise background requests.",
    ];
  }
  if (text.includes("could not connect") || text.includes("could not reach") || text.includes("econnrefused")) {
    return [
      "Make sure Ollama is running.",
      "Check Endpoint matches your local AI server (default: http://127.0.0.1:11434).",
      "Click Test AI to confirm connection before generating.",
    ];
  }
  if (text.includes("not found locally")) {
    return [
      "Model tag is not installed locally.",
      "Run: ollama pull <model-tag>",
      "Or choose an installed model from the dropdown and retry.",
    ];
  }
  if (text.includes("empty output")) {
    return [
      "The model returned no content. Retry once.",
      "Lower Temperature slightly and shorten the prompt.",
      "Try another local model if this repeats.",
    ];
  }
  return [
    "Click Test AI and confirm endpoint/model.",
    "Retry with Compact context mode on.",
    "If it repeats, use a shorter prompt or a faster model.",
  ];
}

function renderAiTroubleshootingPanel() {
  const errorText = str(ui.aiLastError);
  if (!errorText) return "";
  const tips = formatAiErrorHint(errorText);
  const when = ui.aiLastErrorAt ? new Date(ui.aiLastErrorAt).toLocaleString() : "";
  return `
    <details class="copilot-settings" style="margin-top:8px;">
      <summary>AI Troubleshooting</summary>
      <p class="small" style="margin-top:8px;"><strong>Last error:</strong> ${escapeHtml(errorText)}</p>
      ${when ? `<p class="small">When: ${escapeHtml(when)}</p>` : ""}
      ${
        tips.length
          ? `<ul class="list">${tips.map((tip) => `<li>${escapeHtml(tip)}</li>`).join("")}</ul>`
          : ""
      }
    </details>
  `;
}

function renderAiTestStatus() {
  const status = replaceAiModelLabelsInText(str(ui.aiTestStatus || ""));
  if (!status) return "";
  const lower = status.toLowerCase();
  if (lower.includes("not run")) return "";
  const isRunning = ui.aiBusy || lower.includes("running");
  const testAtMs = safeDate(ui.aiTestAt);
  const isRecent = testAtMs > 0 && Date.now() - testAtMs <= 5 * 60 * 1000;
  if (!isRunning && !isRecent) return "";
  const when = ui.aiTestAt ? new Date(ui.aiTestAt).toLocaleTimeString() : "";
  const summary = summarizeAiTestStatus(status);
  const titleAttr = summary !== status ? ` title="${escapeHtml(status)}"` : "";
  return `<p class="small copilot-status-line"${titleAttr}><strong>${escapeHtml(summary)}</strong>${when ? ` <span class="mono">(${escapeHtml(when)})</span>` : ""}</p>`;
}

function summarizeCopilotStatus(message) {
  const text = str(message).replace(/\s+/g, " ");
  if (!text) return "";
  const lower = text.toLowerCase();
  if (lower.includes("instruction-style output") || lower.includes("fallback generated")) {
    return "Fallback used for this response.";
  }
  if (lower.startsWith("auto-generated for")) {
    return "Auto-generated response.";
  }
  if (lower.startsWith("generated with")) {
    return "Generated.";
  }
  if (lower.startsWith("testing local ai connection")) {
    return "Testing AI connection...";
  }
  if (text.length > 96) return `${text.slice(0, 93)}...`;
  return text;
}

function summarizeAiTestStatus(status) {
  const text = str(status);
  const lower = text.toLowerCase();
  if (!text) return "";
  if (lower.startsWith("passed")) return "AI test: Passed";
  if (lower.startsWith("failed")) return "AI test: Failed";
  if (lower.includes("running")) return "AI test: Running";
  if (lower.includes("not run")) return "AI test: Not run";
  return text.length > 64 ? `${text.slice(0, 61)}...` : text;
}

function compactLine(text, max = 120) {
  const clean = str(text).replace(/\s+/g, " ");
  const limit = Number.isFinite(Number(max)) ? Math.max(24, Number(max)) : 120;
  if (clean.length <= limit) return clean;
  return `${clean.slice(0, limit - 3)}...`;
}
