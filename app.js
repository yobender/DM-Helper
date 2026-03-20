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

const RULE_STORE_KIND_LABELS = Object.freeze({
  official_note: "Official Note",
  accepted_ruling: "Accepted Ruling",
  house_rule: "House Rule",
  canon_memory: "Canon Memory",
});

const tabs = [
  { id: "dashboard", label: "Dashboard", group: "Run Game" },
  { id: "sessions", label: "Session Runner", group: "Run Game" },
  { id: "capture", label: "Live Capture HUD", group: "Run Game" },
  { id: "writing", label: "Writing Helper", group: "Run Game" },
  { id: "kingdom", label: "Kingdom", group: "World" },
  { id: "hexmap", label: "Hex Map", group: "World" },
  { id: "npcs", label: "NPCs", group: "World" },
  { id: "quests", label: "Quests", group: "World" },
  { id: "locations", label: "Locations", group: "World" },
  { id: "rules", label: "PF2e Rules", group: "Tools" },
  { id: "pdf", label: "PDF Intel", group: "Tools" },
  { id: "obsidian", label: "Obsidian Vault", group: "Tools" },
  { id: "foundry", label: "Foundry Export", group: "Tools" },
];

const tabGroups = ["Run Game", "World", "Tools"];

const desktopApi = window.kmDesktop || null;
const kingdomRulesData = await loadKingdomRulesData();

let activeTab = "dashboard";
let state = null;
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
  rulesBusy: false,
  rulesMessage: "",
  rulesSearchQuery: "",
  rulesSearchLimit: 5,
  rulesScope: "both",
  rulesResults: [],
  rulesSelectedUrl: "",
  rulesIndexedAt: "",
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
  copilotRetrievalPreview: null,
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
  obsidianMessage: "",
  obsidianBusy: false,
  hexMapMessage: "",
  hexMapSelectedHex: "",
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
  dashboardMessage: "",
  startupError: "",
};
let hexMapPointerState = null;
let hexMapSuppressClickUntil = 0;
let hexMapViewportSaveTimer = 0;

const tabsEl = document.getElementById("tabs");
const appEl = document.getElementById("app");
const seedBtn = document.getElementById("seed-btn");
const exportBtn = document.getElementById("export-btn");
const importBtn = document.getElementById("import-btn");
const importFile = document.getElementById("import-file");

window.addEventListener("error", (event) => {
  handleStartupFailure(event?.error || event?.message || "Unknown renderer error.");
});

window.addEventListener("unhandledrejection", (event) => {
  handleStartupFailure(event?.reason || "Unhandled promise rejection.");
});

queueMicrotask(() => {
  bootApp();
});
// Avoid startup lockups: only auto-run after an explicit tab change.

function bootApp() {
  try {
    state = loadState();
    state.meta.aiMemory = buildAiMemoryDigests(state);
    wireGlobalEvents();
    render();
    void initDesktopDefaults();
  } catch (err) {
    handleStartupFailure(err);
  }
}

function handleStartupFailure(err) {
  const message = readableError(err);
  if (ui.startupError === message) return;
  ui.startupError = message;
  console.error("DM Helper startup failed:", err);
  try {
    state = createStarterState();
    render();
    renderStartupRecoveryBanner(message);
  } catch (fallbackErr) {
    console.error("DM Helper startup fallback failed:", fallbackErr);
    tabsEl.innerHTML = "";
    appEl.innerHTML = `
      <section class="panel fatal-panel">
        <h2>Startup Error</h2>
        <p>DM Helper could not finish loading.</p>
        <pre>${escapeHtml(message)}</pre>
        <p>Close the app and send this error text back here.</p>
      </section>
    `;
  }
}

function renderStartupRecoveryBanner(message) {
  const summary = escapeHtml(message);
  const warning = `
    <section class="panel startup-warning">
      <h2>Recovery Mode</h2>
      <p>Saved app state failed to load. DM Helper started with safe starter data for this launch.</p>
      <details>
        <summary>Startup error details</summary>
        <pre>${summary}</pre>
      </details>
    </section>
  `;
  appEl.innerHTML = `${warning}${appEl.innerHTML}`;
}

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

  appEl.addEventListener("pointerdown", (event) => {
    handleHexMapPointerDown(event);
  });

  window.addEventListener("pointermove", (event) => {
    handleHexMapPointerMove(event);
  });

  window.addEventListener("pointerup", () => {
    finishHexMapPointerSession();
  });

  window.addEventListener("pointercancel", () => {
    finishHexMapPointerSession();
  });

  appEl.addEventListener(
    "wheel",
    (event) => {
      handleHexMapWheel(event);
    },
    { passive: false }
  );

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
    ui.obsidianMessage = "";
    ui.obsidianBusy = false;
    ui.hexMapMessage = "";
    ui.hexMapSelectedHex = "";
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
      ui.obsidianMessage = "";
      ui.obsidianBusy = false;
      ui.hexMapMessage = "";
      ui.hexMapSelectedHex = "";
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

    if (action === "hexmap-select-hex") {
      if (Date.now() < hexMapSuppressClickUntil) return;
      const targetHex = normalizeHexCoordinate(button.dataset.hex);
      if (!targetHex) return;
      const hexMap = getHexMapState();
      if (hexMap.partyMoveMode) {
        const party = moveHexMapPartyToHex(targetHex);
        setHexMapSelectedHex(targetHex);
        saveState();
        ui.hexMapMessage = `${party.label || "Party"} moved to ${targetHex}.`;
        render();
        return;
      }
      setHexMapSelectedHex(targetHex);
      render();
      return;
    }

    if (action === "hexmap-toggle-party-move") {
      const hexMap = getHexMapState();
      hexMap.partyMoveMode = !hexMap.partyMoveMode;
      state.hexMap = normalizeHexMapState(hexMap);
      saveState();
      ui.hexMapMessage = hexMap.partyMoveMode
        ? "Party move mode is on. Click any hex on the map to move the party there."
        : "Party move mode is off. Clicking the map selects hexes again.";
      render();
      return;
    }

    if (action === "hexmap-zoom-in" || action === "hexmap-zoom-out") {
      const hexMap = getHexMapState();
      const factor = action === "hexmap-zoom-in" ? 1.15 : 0.87;
      hexMap.zoom = Math.max(HEX_MAP_ZOOM_MIN, Math.min(HEX_MAP_ZOOM_MAX, Number((hexMap.zoom * factor).toFixed(2))));
      clampHexMapPan(hexMap);
      saveState();
      render();
      return;
    }

    if (action === "hexmap-reset-view") {
      const hexMap = getHexMapState();
      hexMap.zoom = 1;
      hexMap.panX = 0;
      hexMap.panY = 0;
      saveState();
      render();
      return;
    }

    if (action === "hexmap-pan") {
      const hexMap = getHexMapState();
      const view = getHexMapViewBox(hexMap);
      const stepX = Math.max(30, view.width * 0.18);
      const stepY = Math.max(30, view.height * 0.18);
      const direction = str(button.dataset.direction).toLowerCase();
      if (direction === "left") hexMap.panX -= stepX;
      if (direction === "right") hexMap.panX += stepX;
      if (direction === "up") hexMap.panY -= stepY;
      if (direction === "down") hexMap.panY += stepY;
      clampHexMapPan(hexMap);
      saveState();
      render();
      return;
    }

    if (action === "hexmap-clear-region") {
      const selectedHex = getHexMapSelectedHex();
      if (!selectedHex) return;
      if (!confirm(`Clear the region record for ${selectedHex}?`)) return;
      const removed = clearHexMapRegion(selectedHex);
      saveState();
      ui.hexMapMessage = removed ? `Cleared region record for ${selectedHex}.` : `No region record existed for ${selectedHex}.`;
      render();
      return;
    }

    if (action === "hexmap-choose-background") {
      void chooseHexMapBackground();
      return;
    }

    if (action === "hexmap-clear-background") {
      const hexMap = getHexMapState();
      hexMap.backgroundPath = "";
      hexMap.backgroundUrl = "";
      hexMap.backgroundName = "";
      hexMap.backgroundNaturalWidth = 0;
      hexMap.backgroundNaturalHeight = 0;
      hexMap.backgroundScale = 1;
      hexMap.backgroundOffsetX = 0;
      hexMap.backgroundOffsetY = 0;
      saveState();
      ui.hexMapMessage = "Map background cleared.";
      render();
      return;
    }

    if (action === "hexmap-fit-background") {
      const hexMap = getHexMapState();
      if (!hexMap.backgroundUrl) return;
      hexMap.backgroundScale = 1;
      hexMap.backgroundOffsetX = 0;
      hexMap.backgroundOffsetY = 0;
      saveState();
      ui.hexMapMessage = "Background recentered. Fine-tune with scale and offset if needed.";
      render();
      return;
    }

    if (action === "hexmap-clear-party") {
      const hexMap = getHexMapState();
      hexMap.party = {
        hex: "",
        label: str(hexMap.party?.label) || "Party",
        notes: "",
        updatedAt: new Date().toISOString(),
        trail: [],
      };
      state.hexMap = normalizeHexMapState(hexMap);
      saveState();
      ui.hexMapMessage = "Party marker cleared.";
      render();
      return;
    }

    if (action === "hexmap-center-party") {
      const hexMap = getHexMapState();
      const party = getHexMapParty(hexMap);
      if (!party.hex) return;
      setHexMapSelectedHex(party.hex);
      centerHexMapOnHex(party.hex, hexMap);
      saveState();
      render();
      return;
    }

    if (action === "hexmap-clear-party-trail") {
      const hexMap = getHexMapState();
      const party = getHexMapParty(hexMap);
      hexMap.party = {
        ...party,
        trail: party.hex
          ? [
              {
                hex: party.hex,
                at: new Date().toISOString(),
              },
            ]
          : [],
      };
      state.hexMap = normalizeHexMapState(hexMap);
      saveState();
      ui.hexMapMessage = "Party travel history cleared.";
      render();
      return;
    }

    if (action === "obsidian-choose-vault") {
      void chooseObsidianVault();
      return;
    }

    if (action === "obsidian-open-vault") {
      const settings = ensureObsidianSettings();
      if (!settings.vaultPath || !desktopApi?.openPath) return;
      void desktopApi.openPath(settings.vaultPath);
      return;
    }

    if (action === "obsidian-sync") {
      void syncObsidianVault();
      return;
    }

    if (action === "ai-memory-refresh") {
      refreshAiMemoryDigests({ persist: true });
      return;
    }

    if (action === "obsidian-write-current-ai" || action === "ai-copilot-write-vault") {
      void writeCurrentAiOutputToObsidian();
      return;
    }

    if (action === "rules-refresh-search") {
      void runRulesSearch(ui.rulesSearchQuery, ui.rulesSearchLimit, true);
      return;
    }

    if (action === "rules-select-result") {
      const url = decodeURIComponent(str(button.dataset.url));
      ui.rulesSelectedUrl = url;
      render();
      return;
    }

    if (action === "rules-open-result") {
      const url = decodeURIComponent(str(button.dataset.url));
      if (url && desktopApi?.openExternal) {
        void desktopApi.openExternal(url);
      }
      return;
    }

    if (action === "rules-save-result") {
      const url = decodeURIComponent(str(button.dataset.url));
      if (url) ui.rulesSelectedUrl = url;
      saveSelectedOfficialRuleToStore();
      return;
    }

    if (action === "rules-open-store-source") {
      const url = decodeURIComponent(str(button.dataset.url));
      if (url && desktopApi?.openExternal) {
        void desktopApi.openExternal(url);
      }
      return;
    }

    if (action === "rules-delete-store-entry") {
      const id = decodeURIComponent(str(button.dataset.id));
      if (deleteRulesStoreEntry(id)) {
        state.meta.aiMemory = buildAiMemoryDigests(state);
        saveState();
        ui.rulesMessage = "Local rules/canon entry removed.";
      }
      render();
      return;
    }

    if (action === "rules-use-result" || action === "rules-compare-result") {
      const url = decodeURIComponent(str(button.dataset.url));
      if (url) ui.rulesSelectedUrl = url;
      const selected = getSelectedRulesResult(ui.rulesResults);
      ui.copilotOpen = true;
      ui.copilotDraft.input = selected
        ? buildRulesPromptFromResult(selected, action === "rules-compare-result" ? "compare" : "explain")
        : buildRulesQuestionPrompt(ui.rulesSearchQuery || "");
      ui.copilotDraft.output = "";
      ui.copilotRetrievalPreview = null;
      ui.copilotMessage = action === "rules-compare-result"
        ? "Loaded official-vs-local PF2e rules prompt."
        : "Loaded official PF2e rules prompt.";
      render();
      return;
    }

    if (action === "rules-use-query") {
      ui.copilotOpen = true;
      ui.copilotDraft.input = buildRulesQuestionPrompt(ui.rulesSearchQuery || "");
      ui.copilotDraft.output = "";
      ui.copilotRetrievalPreview = null;
      ui.copilotMessage = "Loaded the current PF2e rules query into Loremaster.";
      render();
      return;
    }

    if (action === "ai-copilot-save-ruling") {
      saveCopilotOutputToRulesStore("accepted_ruling");
      return;
    }

    if (action === "ai-copilot-save-canon") {
      saveCopilotOutputToRulesStore("canon_memory");
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

    if (action === "kingdom-recalculate-creation") {
      const form = button.closest("form");
      if (!(form instanceof HTMLFormElement)) return;
      const fields = getFormFields(form);
      const nextLevel = Math.max(1, Number.parseInt(String(fields.level || getKingdomState().level || "1"), 10) || 1);
      if (
        nextLevel > 1 &&
        !confirm("This resets the kingdom's ability modifiers and trained-skill baseline to the creation package. Continue?")
      ) {
        return;
      }
      const plan = recalculateKingdomFromCreationChoices(fields);
      saveState();
      ui.kingdomMessage = `Recalculated creation package: ${Object.entries(plan.abilityAdjustments)
        .map(([key, value]) => `${KINGDOM_ABILITY_LABELS[key]} ${formatSignedNumber(value)}`)
        .join(", ")}.`;
      render();
      return;
    }

    if (action === "kingdom-use-creation-choice") {
      const form = appEl.querySelector('form[data-form="kingdom-overview"]');
      if (!(form instanceof HTMLFormElement)) return;
      const section = str(button.dataset.section);
      const value = str(button.dataset.value);
      const fieldName =
        section === "charters" ? "charter" : section === "governments" ? "government" : section === "heartlands" ? "heartland" : "";
      if (!fieldName) return;
      const field = form.elements.namedItem(fieldName);
      if (field instanceof HTMLSelectElement || field instanceof HTMLInputElement) {
        field.value = value;
      }
      syncKingdomCreationPreview(form);
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
      ui.copilotRetrievalPreview = null;
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
      ui.copilotRetrievalPreview = null;
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
      ui.copilotRetrievalPreview = null;
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

    const rulesForm = input.closest('form[data-form="rules-search"]');
    if (rulesForm instanceof HTMLFormElement) {
      if (input.getAttribute("name") === "query") ui.rulesSearchQuery = normalizeRulesSearchQuery(input.value);
      if (input.getAttribute("name") === "limit") ui.rulesSearchLimit = Math.max(1, Math.min(6, Number.parseInt(String(input.value || "5"), 10) || 5));
      if (input.getAttribute("name") === "scope") {
        ui.rulesScope = ["both", "official", "local"].includes(str(input.value)) ? str(input.value) : "both";
        render();
      }
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
      if (copilotField === "input") {
        ui.copilotRetrievalPreview = null;
      }
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

    const kingdomForm = input.closest('form[data-form="kingdom-overview"]');
    if (kingdomForm instanceof HTMLFormElement) {
      syncKingdomCreationPreview(kingdomForm);
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

  appEl.addEventListener("input", (event) => {
    const input = event.target;
    if (
      !(
        input instanceof HTMLInputElement ||
        input instanceof HTMLTextAreaElement ||
        input instanceof HTMLSelectElement
      )
    )
      return;

    const kingdomForm = input.closest('form[data-form="kingdom-overview"]');
    if (kingdomForm instanceof HTMLFormElement) {
      syncKingdomCreationPreview(kingdomForm);
    }

    const rulesForm = input.closest('form[data-form="rules-search"]');
    if (rulesForm instanceof HTMLFormElement) {
      if (input.getAttribute("name") === "query") {
        ui.rulesSearchQuery = normalizeRulesSearchQuery(input.value);
      }
    }
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
  if (activeTab === "hexmap") content = renderHexMap();
  if (activeTab === "npcs") content = renderNpcs();
  if (activeTab === "quests") content = renderQuests();
  if (activeTab === "locations") content = renderLocations();
  if (activeTab === "rules") content = renderRules();
  if (activeTab === "pdf") content = renderPdfIntel();
  if (activeTab === "obsidian") content = renderObsidian();
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
  const obsidianSettings = ensureObsidianSettings();
  const tabLabel = getTabLabel(activeTab);
  const routePreview = buildGlobalCopilotRequest(activeTab, ui.copilotDraft.input, false);
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
        <p class="small">Route: <strong>${escapeHtml(routePreview.taskLabel || "General Prep")}</strong> • Mode ${escapeHtml(
          routePreview.mode || getGlobalCopilotMode(activeTab)
        )} • Save target ${escapeHtml(routePreview.saveTarget || "latest session prep")}${routePreview.routeReason ? ` • ${escapeHtml(routePreview.routeReason)}` : ""}</p>
        <div class="toolbar">
          <button class="btn btn-primary" data-action="ai-copilot-generate">Generate</button>
          <button class="btn btn-secondary" data-action="ai-copilot-seed">Smart Prompt</button>
          <button class="btn btn-secondary" data-action="ai-copilot-test" ${aiBusy ? "disabled" : ""}>${testLabel}</button>
          <button class="btn btn-secondary" data-action="ai-copilot-unlock">Unlock</button>
          <button class="btn btn-secondary" data-action="ai-copilot-copy" ${hasOutput ? "" : "disabled"}>Copy</button>
          <button class="btn btn-secondary" data-action="ai-copilot-write-vault" ${(hasOutput && obsidianSettings.vaultPath && !ui.obsidianBusy) ? "" : "disabled"}>Write To Vault</button>
          <button class="btn btn-secondary" data-action="ai-copilot-save-ruling" ${hasOutput ? "" : "disabled"}>Save As Ruling</button>
          <button class="btn btn-secondary" data-action="ai-copilot-save-canon" ${hasOutput ? "" : "disabled"}>Save As Canon</button>
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
          <label style="margin-top:8px;">
            <input type="checkbox" data-ai-field="useAonRules" ${aiConfig.useAonRules ? "checked" : ""} />
            Use Archives of Nethys live lookup for PF2e rules questions
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
        ${renderCopilotRetrievalPreview(ui.copilotRetrievalPreview)}
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

function renderCopilotRetrievalPreview(preview) {
  if (!preview || typeof preview !== "object") {
    return `
      <details class="copilot-settings">
        <summary>Retrieved Context (0)</summary>
        <p class="small">No retrieval run yet for the current prompt.</p>
      </details>
    `;
  }

  const chunks = Array.isArray(preview.chunks) ? preview.chunks : [];
  const sourceCounts = preview.sourceCounts && typeof preview.sourceCounts === "object" ? preview.sourceCounts : {};
  const sourceSummary = Object.entries(sourceCounts)
    .filter((entry) => Number(entry[1] || 0) > 0)
    .sort((a, b) => Number(b[1] || 0) - Number(a[1] || 0) || String(a[0]).localeCompare(String(b[0])))
    .map(([key, value]) => `${key} ${value}`)
    .join(" • ");

  return `
    <details class="copilot-settings">
      <summary>Retrieved Context (${chunks.length})</summary>
      <p class="small">
        Query: ${escapeHtml(str(preview.query || "none"))}
        ${preview.taskLabel ? ` • Task ${escapeHtml(str(preview.taskLabel))}` : ""}
        ${preview.generatedAt ? ` • Built ${escapeHtml(str(preview.generatedAt))}` : ""}
      </p>
      <p class="small">
        ${escapeHtml(sourceSummary || "No source chunks selected.")}
      </p>
      ${
        chunks.length
          ? `<div class="ai-history-list">${chunks
              .map(
                (chunk) => `
                  <article class="panel" style="margin-top:8px;">
                    <strong>${escapeHtml(str(chunk.title || "Untitled chunk"))}</strong>
                    <p class="small">${escapeHtml(
                      [str(chunk.sourceLabel || chunk.sourceType), chunk.reason, chunk.score ? `score ${chunk.score}` : ""]
                        .filter(Boolean)
                        .join(" • ")
                    )}</p>
                    <div class="memory-block">${renderMultilineText(str(chunk.text || "No excerpt."))}</div>
                  </article>
                `
              )
              .join("")}</div>`
          : `<p class="small">No retrieval chunks were selected for the last run.</p>`
      }
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

function renderMultilineText(text) {
  return escapeHtml(str(text)).replace(/\n/g, "<br />");
}

function renderAiMemoryCard(title, text) {
  return `
    <article class="memory-card">
      <h3>${escapeHtml(title)}</h3>
      <div class="memory-block">${renderMultilineText(text || "None yet.")}</div>
    </article>
  `;
}

function renderDashboard() {
  const aiMemory = buildAiMemoryDigests(state);
  state.meta.aiMemory = aiMemory;
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
          <div class="panel-head">
            <h2>AI Memory Layer</h2>
            <div class="toolbar">
              <button class="btn btn-secondary" data-action="ai-memory-refresh">Refresh Memory</button>
            </div>
          </div>
          <p class="small">Last updated ${escapeHtml(aiMemory.updatedAt || "never")} • sources: ${escapeHtml(
            `sessions ${aiMemory.sourceCounts.sessions}, open quests ${aiMemory.sourceCounts.openQuests}, NPCs ${aiMemory.sourceCounts.npcs}, locations ${aiMemory.sourceCounts.locations}, rule entries ${aiMemory.sourceCounts.ruleEntries}, canon ${aiMemory.sourceCounts.canonEntries}`
          )}</p>
          ${ui.dashboardMessage ? `<p class="small">${escapeHtml(ui.dashboardMessage)}</p>` : ""}
          <div class="memory-grid">
            ${renderAiMemoryCard("Campaign Summary", aiMemory.campaignSummary)}
            ${renderAiMemoryCard("Recent Session Summary", aiMemory.recentSessionSummary)}
            ${renderAiMemoryCard("Active Quests", aiMemory.activeQuestsSummary)}
            ${renderAiMemoryCard("Active Entities", aiMemory.activeEntitiesSummary)}
            ${renderAiMemoryCard("Canon Memory", aiMemory.canonSummary)}
          </div>
        </article>

        <article class="panel">
          <h2>Rulings / House Rules Digest</h2>
          <p class="small">This is the authoritative manual digest Loremaster should prefer when a PF2e ruling or house-rule question comes up. If left blank, DM Helper falls back to recent Rule / Retcon capture entries.</p>
          <form data-form="ai-memory-rulings">
            <label>Manual Rulings Digest
              <textarea name="manualRulings" placeholder="Example: Hero Point rerolls must be declared before any new info is revealed. Persistent damage from house-rule fire traps ticks at end of round, not end of turn.">${escapeHtml(
                aiMemory.manualRulings || ""
              )}</textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Save Rulings Digest</button>
            </div>
          </form>
          <div class="memory-card">
            <h3>Effective Rulings Digest</h3>
            <div class="memory-block">${renderMultilineText(aiMemory.rulingsDigest || "No rulings digest captured yet.")}</div>
          </div>
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
            <label style="margin-top:8px;">
              <input type="checkbox" data-ai-field="useAonRules" ${aiConfig.useAonRules ? "checked" : ""} />
              Use Archives of Nethys live lookup for PF2e rules questions
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

const KINGDOM_SKILL_DEFINITIONS = [
  { name: "Agriculture", ability: "economy" },
  { name: "Arts", ability: "culture" },
  { name: "Boating", ability: "economy" },
  { name: "Defense", ability: "stability" },
  { name: "Engineering", ability: "stability" },
  { name: "Exploration", ability: "stability" },
  { name: "Folklore", ability: "culture" },
  { name: "Industry", ability: "economy" },
  { name: "Intrigue", ability: "loyalty" },
  { name: "Magic", ability: "culture" },
  { name: "Politics", ability: "loyalty" },
  { name: "Scholarship", ability: "culture" },
  { name: "Statecraft", ability: "loyalty" },
  { name: "Trade", ability: "economy" },
  { name: "Warfare", ability: "loyalty" },
  { name: "Wilderness", ability: "stability" }
];

const KINGDOM_ABILITY_KEYS = ["culture", "economy", "loyalty", "stability"];
const KINGDOM_ABILITY_LABELS = {
  culture: "Culture",
  economy: "Economy",
  loyalty: "Loyalty",
  stability: "Stability"
};
const KINGDOM_SKILL_RANKS = ["untrained", "trained", "expert", "master", "legendary"];
const KINGDOM_SKILL_RANK_LABELS = {
  untrained: "Untrained",
  trained: "Trained",
  expert: "Expert",
  master: "Master",
  legendary: "Legendary"
};
const KINGDOM_SETTLEMENT_ACTIONS = {
  "Town Hall": 1,
  Castle: 2,
  Palace: 2
};
const HEX_MAP_MARKER_TYPES = ["Encounter", "Building", "Event", "Settlement", "Resource", "Danger", "Note"];
const HEX_MAP_FORCE_TYPES = ["Allied Force", "Enemy Force", "Caravan"];
const HEX_MAP_STATUS_OPTIONS = ["Unclaimed", "Reconnoitered", "Claimed", "Work Site", "Settlement", "Contested"];
const HEX_MAP_TERRAIN_OPTIONS = ["Plains", "Forest", "Hills", "Mountains", "Marsh", "River", "Lake", "Ruins", "Road", "Settlement"];
const HEX_MAP_HEX_SIZE_MIN = 30;
const HEX_MAP_HEX_SIZE_MAX = 110;
const HEX_MAP_COLUMNS_MIN = 6;
const HEX_MAP_COLUMNS_MAX = 20;
const HEX_MAP_ROWS_MIN = 4;
const HEX_MAP_ROWS_MAX = 16;
const HEX_MAP_ZOOM_MIN = 0.7;
const HEX_MAP_ZOOM_MAX = 6;
const HEX_MAP_BACKGROUND_SCALE_MIN = 0.4;
const HEX_MAP_BACKGROUND_SCALE_MAX = 3.5;

function canonicalizeKingdomSkillName(value) {
  const clean = str(value).trim().toLowerCase();
  if (!clean) return "";
  return KINGDOM_SKILL_DEFINITIONS.find((entry) => entry.name.toLowerCase() === clean)?.name || "";
}

function canonicalizeKingdomAbilityName(value) {
  const clean = str(value).trim().toLowerCase();
  return KINGDOM_ABILITY_KEYS.find((entry) => entry === clean) || "";
}

function getKingdomSkillDefinition(skillName) {
  const clean = canonicalizeKingdomSkillName(skillName);
  return KINGDOM_SKILL_DEFINITIONS.find((entry) => entry.name === clean) || null;
}

function getKingdomSkillRank(rank) {
  const clean = str(rank).trim().toLowerCase();
  return KINGDOM_SKILL_RANKS.includes(clean) ? clean : "untrained";
}

function getKingdomSkillRankWeight(rank) {
  return KINGDOM_SKILL_RANKS.indexOf(getKingdomSkillRank(rank));
}

function parseSkillList(value, canonical = false) {
  const parts = str(value)
    .split(",")
    .map((entry) => str(entry).trim())
    .filter(Boolean);
  if (!canonical) return [...new Set(parts)];
  const normalized = parts.map((entry) => canonicalizeKingdomSkillName(entry)).filter(Boolean);
  return [...new Set(normalized)];
}

function parseCommaList(value) {
  return str(value)
    .split(",")
    .map((entry) => str(entry).trim())
    .filter(Boolean);
}

function parseAbilityList(value) {
  return parseCommaList(value).map((entry) => canonicalizeKingdomAbilityName(entry)).filter(Boolean);
}

function createStarterKingdomCreationState() {
  return {
    freeAbilityBoosts: [],
    charterSkill: "",
    heartlandSkill: "",
    bonusSkills: []
  };
}

function normalizeKingdomCreationState(input) {
  const base = createStarterKingdomCreationState();
  const out = {
    ...base,
    ...(input && typeof input === "object" && !Array.isArray(input) ? input : {})
  };
  out.freeAbilityBoosts = parseAbilityList(Array.isArray(out.freeAbilityBoosts) ? out.freeAbilityBoosts.join(", ") : out.freeAbilityBoosts);
  out.charterSkill = canonicalizeKingdomSkillName(out.charterSkill);
  out.heartlandSkill = canonicalizeKingdomSkillName(out.heartlandSkill);
  out.bonusSkills = parseSkillList(Array.isArray(out.bonusSkills) ? out.bonusSkills.join(", ") : out.bonusSkills, true).slice(0, 8);
  return out;
}

function createStarterHexMapState() {
  return {
    mapName: "Stolen Lands Hex Planner",
    columns: 10,
    rows: 8,
    hexSize: 54,
    zoom: 1,
    panX: 0,
    panY: 0,
    showLabels: true,
    backgroundPath: "",
    backgroundUrl: "",
    backgroundName: "",
    backgroundNaturalWidth: 0,
    backgroundNaturalHeight: 0,
    backgroundOpacity: 0.5,
    backgroundScale: 1,
    backgroundOffsetX: 0,
    backgroundOffsetY: 0,
    gridFillOpacity: 0.16,
    gridLineOpacity: 0.65,
    partyMoveMode: false,
    party: {
      hex: "",
      label: "Party",
      notes: "",
      updatedAt: "",
      trail: [],
    },
    forces: [],
    markers: [],
  };
}

function normalizeHexMapState(input) {
  const base = createStarterHexMapState();
  const out = {
    ...base,
    ...(input && typeof input === "object" && !Array.isArray(input) ? input : {}),
  };
  out.mapName = str(out.mapName) || base.mapName;
  out.columns = Math.max(HEX_MAP_COLUMNS_MIN, Math.min(HEX_MAP_COLUMNS_MAX, Number.parseInt(String(out.columns || base.columns), 10) || base.columns));
  out.rows = Math.max(HEX_MAP_ROWS_MIN, Math.min(HEX_MAP_ROWS_MAX, Number.parseInt(String(out.rows || base.rows), 10) || base.rows));
  out.hexSize = Math.max(HEX_MAP_HEX_SIZE_MIN, Math.min(HEX_MAP_HEX_SIZE_MAX, Number.parseInt(String(out.hexSize || base.hexSize), 10) || base.hexSize));
  out.zoom = Math.max(HEX_MAP_ZOOM_MIN, Math.min(HEX_MAP_ZOOM_MAX, Number.parseFloat(String(out.zoom || base.zoom)) || base.zoom));
  out.panX = Number.isFinite(Number.parseFloat(String(out.panX))) ? Number.parseFloat(String(out.panX)) : 0;
  out.panY = Number.isFinite(Number.parseFloat(String(out.panY))) ? Number.parseFloat(String(out.panY)) : 0;
  out.showLabels = out.showLabels !== false;
  out.backgroundPath = str(out.backgroundPath);
  out.backgroundUrl = str(out.backgroundUrl);
  out.backgroundName = str(out.backgroundName);
  out.backgroundNaturalWidth = Math.max(0, Number.parseFloat(String(out.backgroundNaturalWidth || "0")) || 0);
  out.backgroundNaturalHeight = Math.max(0, Number.parseFloat(String(out.backgroundNaturalHeight || "0")) || 0);
  out.backgroundOpacity = Math.max(0, Math.min(0.95, Number.parseFloat(String(out.backgroundOpacity ?? base.backgroundOpacity)) || base.backgroundOpacity));
  out.backgroundScale = Math.max(
    HEX_MAP_BACKGROUND_SCALE_MIN,
    Math.min(HEX_MAP_BACKGROUND_SCALE_MAX, Number.parseFloat(String(out.backgroundScale ?? base.backgroundScale)) || base.backgroundScale)
  );
  out.backgroundOffsetX = Number.isFinite(Number.parseFloat(String(out.backgroundOffsetX)))
    ? Number.parseFloat(String(out.backgroundOffsetX))
    : 0;
  out.backgroundOffsetY = Number.isFinite(Number.parseFloat(String(out.backgroundOffsetY)))
    ? Number.parseFloat(String(out.backgroundOffsetY))
    : 0;
  out.gridFillOpacity = Math.max(0, Math.min(0.65, Number.parseFloat(String(out.gridFillOpacity ?? base.gridFillOpacity)) || base.gridFillOpacity));
  out.gridLineOpacity = Math.max(0.15, Math.min(1, Number.parseFloat(String(out.gridLineOpacity ?? base.gridLineOpacity)) || base.gridLineOpacity));
  out.partyMoveMode = out.partyMoveMode === true;
  const rawParty = out.party && typeof out.party === "object" && !Array.isArray(out.party) ? out.party : {};
  out.party = {
    hex: normalizeHexCoordinate(rawParty.hex, out.columns, out.rows) || "",
    label: str(rawParty.label) || "Party",
    notes: str(rawParty.notes),
    updatedAt: str(rawParty.updatedAt) || "",
    trail: Array.isArray(rawParty.trail)
      ? rawParty.trail
          .map((entry) => ({
            hex: normalizeHexCoordinate(entry?.hex, out.columns, out.rows),
            at: str(entry?.at) || "",
          }))
          .filter((entry) => entry.hex)
          .slice(0, 30)
      : [],
  };
  out.forces = Array.isArray(out.forces)
    ? out.forces
        .map((force) => ({
          id: str(force?.id) || uid(),
          hex: normalizeHexCoordinate(force?.hex, out.columns, out.rows),
          type: HEX_MAP_FORCE_TYPES.includes(str(force?.type)) ? str(force?.type) : "Allied Force",
          name: str(force?.name) || "Unnamed force",
          notes: str(force?.notes),
          updatedAt: str(force?.updatedAt) || "",
        }))
        .filter((force) => force.hex)
    : [];
  out.markers = Array.isArray(out.markers)
    ? out.markers
        .map((marker) => ({
          id: str(marker?.id) || uid(),
          hex: normalizeHexCoordinate(marker?.hex),
          type: HEX_MAP_MARKER_TYPES.includes(str(marker?.type)) ? str(marker?.type) : "Note",
          title: str(marker?.title),
          notes: str(marker?.notes),
          updatedAt: str(marker?.updatedAt) || "",
        }))
        .filter((marker) => marker.hex)
    : [];
  return out;
}

function getHexMapState() {
  if (!state.hexMap || typeof state.hexMap !== "object" || Array.isArray(state.hexMap)) {
    state.hexMap = createStarterHexMapState();
  } else {
    state.hexMap = normalizeHexMapState(state.hexMap);
  }
  return state.hexMap;
}

function getHexColumnLabel(index) {
  let value = Math.max(0, Number.parseInt(String(index || "0"), 10) || 0) + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label || "A";
}

function getHexColumnIndex(label) {
  const clean = String(label || "").trim().toUpperCase();
  if (!/^[A-Z]+$/.test(clean)) return -1;
  let total = 0;
  for (const char of clean) {
    total = total * 26 + (char.charCodeAt(0) - 64);
  }
  return total - 1;
}

function parseHexCoordinate(value) {
  const clean = String(value || "").trim().toUpperCase().replace(/\s+/g, "");
  const match = clean.match(/^([A-Z]+)[-:]?(\d{1,2})$/);
  if (!match) return null;
  return {
    columnLabel: match[1],
    columnIndex: getHexColumnIndex(match[1]),
    rowIndex: Math.max(0, Number.parseInt(match[2], 10) - 1),
  };
}

function normalizeHexCoordinate(value, columns = 0, rows = 0) {
  const parsed = parseHexCoordinate(value);
  if (!parsed || parsed.columnIndex < 0 || parsed.rowIndex < 0) return "";
  if (columns && parsed.columnIndex >= columns) return "";
  if (rows && parsed.rowIndex >= rows) return "";
  return `${getHexColumnLabel(parsed.columnIndex)}${parsed.rowIndex + 1}`;
}

function getHexMapMetrics(hexMap) {
  const size = Math.max(HEX_MAP_HEX_SIZE_MIN, Math.min(HEX_MAP_HEX_SIZE_MAX, Number.parseInt(String(hexMap?.hexSize || 54), 10) || 54));
  const hexWidth = size * 2;
  const hexHeight = Math.sqrt(3) * size;
  const stepX = size * 1.5;
  const margin = size * 1.6;
  const columns = Math.max(HEX_MAP_COLUMNS_MIN, Number.parseInt(String(hexMap?.columns || "10"), 10) || 10);
  const rows = Math.max(HEX_MAP_ROWS_MIN, Number.parseInt(String(hexMap?.rows || "8"), 10) || 8);
  const boardWidth = margin * 2 + hexWidth + Math.max(0, columns - 1) * stepX;
  const boardHeight = margin * 2 + rows * hexHeight + hexHeight / 2;
  return {
    size,
    hexWidth,
    hexHeight,
    stepX,
    margin,
    columns,
    rows,
    boardWidth,
    boardHeight,
  };
}

function getHexCenter(columnIndex, rowIndex, hexMap) {
  const metrics = getHexMapMetrics(hexMap);
  return {
    cx: metrics.margin + metrics.size + columnIndex * metrics.stepX,
    cy: metrics.margin + metrics.hexHeight / 2 + rowIndex * metrics.hexHeight + (columnIndex % 2 ? metrics.hexHeight / 2 : 0),
  };
}

function buildHexPolygonPoints(cx, cy, size) {
  const points = [];
  for (let i = 0; i < 6; i += 1) {
    const angle = (Math.PI / 180) * (60 * i);
    const x = cx + size * Math.cos(angle);
    const y = cy + size * Math.sin(angle);
    points.push(`${x.toFixed(2)},${y.toFixed(2)}`);
  }
  return points.join(" ");
}

function clampHexMapPan(hexMap) {
  const metrics = getHexMapMetrics(hexMap);
  const zoom = Math.max(HEX_MAP_ZOOM_MIN, Math.min(HEX_MAP_ZOOM_MAX, Number(hexMap.zoom || 1)));
  const viewWidth = metrics.boardWidth / zoom;
  const viewHeight = metrics.boardHeight / zoom;
  const maxX = Math.max(0, metrics.boardWidth - viewWidth);
  const maxY = Math.max(0, metrics.boardHeight - viewHeight);
  hexMap.panX = Math.max(0, Math.min(maxX, Number(hexMap.panX || 0)));
  hexMap.panY = Math.max(0, Math.min(maxY, Number(hexMap.panY || 0)));
}

function getHexMapViewBox(hexMap) {
  clampHexMapPan(hexMap);
  const metrics = getHexMapMetrics(hexMap);
  const zoom = Math.max(HEX_MAP_ZOOM_MIN, Math.min(HEX_MAP_ZOOM_MAX, Number(hexMap.zoom || 1)));
  return {
    x: Number(hexMap.panX || 0),
    y: Number(hexMap.panY || 0),
    width: metrics.boardWidth / zoom,
    height: metrics.boardHeight / zoom,
    ...metrics,
  };
}

function syncHexMapViewportDom() {
  const stage = document.querySelector(".hexmap-stage");
  if (!(stage instanceof SVGElement)) return;
  const hexMap = getHexMapState();
  const view = getHexMapViewBox(hexMap);
  stage.setAttribute(
    "viewBox",
    `${view.x.toFixed(2)} ${view.y.toFixed(2)} ${view.width.toFixed(2)} ${view.height.toFixed(2)}`
  );
  const zoomChip = document.querySelector("[data-hexmap-zoom-chip]");
  if (zoomChip) zoomChip.textContent = `Zoom ${Math.round(hexMap.zoom * 100)}%`;
}

function scheduleHexMapViewportSave() {
  if (hexMapViewportSaveTimer) window.clearTimeout(hexMapViewportSaveTimer);
  hexMapViewportSaveTimer = window.setTimeout(() => {
    hexMapViewportSaveTimer = 0;
    saveState();
  }, 180);
}

function handleHexMapPointerDown(event) {
  if (!(event.target instanceof Element)) return;
  const stageShell = event.target.closest("[data-hexmap-stage-shell]");
  if (!(stageShell instanceof HTMLElement)) return;
  if (activeTab !== "hexmap") return;
  if (event.button !== 0) return;
  const hexMap = getHexMapState();
  const view = getHexMapViewBox(hexMap);
  const rect = stageShell.getBoundingClientRect();
  hexMapPointerState = {
    shell: stageShell,
    startClientX: event.clientX,
    startClientY: event.clientY,
    startPanX: Number(hexMap.panX || 0),
    startPanY: Number(hexMap.panY || 0),
    viewWidth: view.width,
    viewHeight: view.height,
    rectWidth: rect.width,
    rectHeight: rect.height,
    didDrag: false,
  };
  stageShell.classList.add("is-dragging");
  event.preventDefault();
}

function handleHexMapPointerMove(event) {
  if (!hexMapPointerState) return;
  const hexMap = getHexMapState();
  const deltaClientX = event.clientX - hexMapPointerState.startClientX;
  const deltaClientY = event.clientY - hexMapPointerState.startClientY;
  if (Math.abs(deltaClientX) > 4 || Math.abs(deltaClientY) > 4) {
    hexMapPointerState.didDrag = true;
  }
  const boardDeltaX = deltaClientX * (hexMapPointerState.viewWidth / Math.max(1, hexMapPointerState.rectWidth));
  const boardDeltaY = deltaClientY * (hexMapPointerState.viewHeight / Math.max(1, hexMapPointerState.rectHeight));
  hexMap.panX = hexMapPointerState.startPanX - boardDeltaX;
  hexMap.panY = hexMapPointerState.startPanY - boardDeltaY;
  clampHexMapPan(hexMap);
  syncHexMapViewportDom();
}

function finishHexMapPointerSession() {
  if (!hexMapPointerState) return;
  const didDrag = hexMapPointerState.didDrag;
  hexMapPointerState.shell?.classList.remove("is-dragging");
  hexMapPointerState = null;
  if (didDrag) {
    hexMapSuppressClickUntil = Date.now() + 180;
    saveState();
  }
}

function handleHexMapWheel(event) {
  if (!(event.target instanceof Element)) return;
  const stageShell = event.target.closest("[data-hexmap-stage-shell]");
  if (!(stageShell instanceof HTMLElement)) return;
  if (activeTab !== "hexmap") return;
  event.preventDefault();
  const hexMap = getHexMapState();
  const metrics = getHexMapMetrics(hexMap);
  const rect = stageShell.getBoundingClientRect();
  const oldView = getHexMapViewBox(hexMap);
  const offsetX = Math.max(0, Math.min(rect.width, event.clientX - rect.left));
  const offsetY = Math.max(0, Math.min(rect.height, event.clientY - rect.top));
  const anchorX = oldView.x + (offsetX / Math.max(1, rect.width)) * oldView.width;
  const anchorY = oldView.y + (offsetY / Math.max(1, rect.height)) * oldView.height;
  const factor = event.deltaY < 0 ? 1.12 : 0.89;
  const nextZoom = Math.max(HEX_MAP_ZOOM_MIN, Math.min(HEX_MAP_ZOOM_MAX, Number((hexMap.zoom * factor).toFixed(2))));
  if (nextZoom === hexMap.zoom) return;
  hexMap.zoom = nextZoom;
  const nextViewWidth = metrics.boardWidth / nextZoom;
  const nextViewHeight = metrics.boardHeight / nextZoom;
  hexMap.panX = anchorX - (offsetX / Math.max(1, rect.width)) * nextViewWidth;
  hexMap.panY = anchorY - (offsetY / Math.max(1, rect.height)) * nextViewHeight;
  clampHexMapPan(hexMap);
  syncHexMapViewportDom();
  scheduleHexMapViewportSave();
}

function getHexMapBackgroundPlacement(hexMap) {
  const metrics = getHexMapMetrics(hexMap);
  const naturalWidth = Math.max(1, Number.parseFloat(String(hexMap?.backgroundNaturalWidth || "0")) || metrics.boardWidth);
  const naturalHeight = Math.max(1, Number.parseFloat(String(hexMap?.backgroundNaturalHeight || "0")) || metrics.boardHeight);
  const fitScale = Math.min(metrics.boardWidth / naturalWidth, metrics.boardHeight / naturalHeight);
  const manualScale = Math.max(
    HEX_MAP_BACKGROUND_SCALE_MIN,
    Math.min(HEX_MAP_BACKGROUND_SCALE_MAX, Number.parseFloat(String(hexMap?.backgroundScale || "1")) || 1)
  );
  const width = naturalWidth * fitScale * manualScale;
  const height = naturalHeight * fitScale * manualScale;
  const offsetX = Number.parseFloat(String(hexMap?.backgroundOffsetX || "0")) || 0;
  const offsetY = Number.parseFloat(String(hexMap?.backgroundOffsetY || "0")) || 0;
  return {
    x: (metrics.boardWidth - width) / 2 + offsetX,
    y: (metrics.boardHeight - height) / 2 + offsetY,
    width,
    height,
    naturalWidth,
    naturalHeight,
    fitScale,
  };
}

function getHexStatusColor(status) {
  const clean = str(status).toLowerCase();
  if (clean === "claimed") return "#4d8f74";
  if (clean === "reconnoitered") return "#d5c187";
  if (clean === "work site") return "#9b6f45";
  if (clean === "settlement") return "#2f7a63";
  if (clean === "contested") return "#b25b47";
  return "#f2ead8";
}

function getHexMarkerColor(type) {
  const clean = str(type).toLowerCase();
  if (clean === "encounter") return "#8a3c2a";
  if (clean === "building") return "#8f6a3d";
  if (clean === "event") return "#9c7b25";
  if (clean === "settlement") return "#2f7a63";
  if (clean === "resource") return "#336d91";
  if (clean === "danger") return "#6d2432";
  return "#6d5a42";
}

function getHexForceColor(type) {
  const clean = str(type).toLowerCase();
  if (clean === "enemy force") return "#8a2f2f";
  if (clean === "caravan") return "#a06a22";
  return "#2f7a63";
}

function getHexForceGlyph(type) {
  const clean = str(type).toLowerCase();
  if (clean === "enemy force") return "E";
  if (clean === "caravan") return "C";
  return "A";
}

function getHexMapSelectedHex(hexMap = getHexMapState()) {
  const kingdom = getKingdomState();
  const candidates = [
    ui.hexMapSelectedHex,
    ...(kingdom.regions || []).map((region) => region.hex),
    ...(hexMap.markers || []).map((marker) => marker.hex),
    "A1",
  ];
  const chosen = candidates
    .map((value) => normalizeHexCoordinate(value, hexMap.columns, hexMap.rows))
    .find(Boolean);
  ui.hexMapSelectedHex = chosen || "A1";
  return ui.hexMapSelectedHex;
}

function setHexMapSelectedHex(value) {
  const hexMap = getHexMapState();
  ui.hexMapSelectedHex = normalizeHexCoordinate(value, hexMap.columns, hexMap.rows) || getHexMapSelectedHex(hexMap);
}

function getKingdomRegionByHex(hex) {
  const clean = normalizeHexCoordinate(hex);
  if (!clean) return null;
  return getKingdomState().regions.find((region) => normalizeHexCoordinate(region.hex) === clean) || null;
}

function getHexMapMarkersForHex(hex) {
  const clean = normalizeHexCoordinate(hex);
  if (!clean) return [];
  return getHexMapState().markers.filter((marker) => normalizeHexCoordinate(marker.hex) === clean);
}

function getHexMapParty(hexMap = getHexMapState()) {
  const cleanHex = normalizeHexCoordinate(hexMap?.party?.hex, hexMap.columns, hexMap.rows) || "";
  return {
    hex: cleanHex,
    label: str(hexMap?.party?.label) || "Party",
    notes: str(hexMap?.party?.notes),
    updatedAt: str(hexMap?.party?.updatedAt) || "",
    trail: Array.isArray(hexMap?.party?.trail)
      ? hexMap.party.trail
          .map((entry) => ({
            hex: normalizeHexCoordinate(entry?.hex, hexMap.columns, hexMap.rows),
            at: str(entry?.at) || "",
          }))
          .filter((entry) => entry.hex)
      : [],
  };
}

function getHexMapForcesForHex(hex) {
  const clean = normalizeHexCoordinate(hex);
  if (!clean) return [];
  return getHexMapState().forces.filter((force) => normalizeHexCoordinate(force.hex) === clean);
}

function centerHexMapOnHex(hex, hexMap = getHexMapState()) {
  const parsed = parseHexCoordinate(hex);
  if (!parsed) return;
  const center = getHexCenter(parsed.columnIndex, parsed.rowIndex, hexMap);
  const metrics = getHexMapMetrics(hexMap);
  const zoom = Math.max(HEX_MAP_ZOOM_MIN, Math.min(HEX_MAP_ZOOM_MAX, Number(hexMap.zoom || 1)));
  hexMap.panX = center.cx - metrics.boardWidth / (2 * zoom);
  hexMap.panY = center.cy - metrics.boardHeight / (2 * zoom);
  clampHexMapPan(hexMap);
}

function moveHexMapPartyToHex(hex, options = {}) {
  const hexMap = getHexMapState();
  const cleanHex = normalizeHexCoordinate(hex, hexMap.columns, hexMap.rows);
  if (!cleanHex) throw new Error("Select a valid hex first.");
  const current = getHexMapParty(hexMap);
  const nextTrail = Array.isArray(current.trail) ? [...current.trail] : [];
  if (!nextTrail.length || nextTrail[0]?.hex !== cleanHex) {
    nextTrail.unshift({
      hex: cleanHex,
      at: new Date().toISOString(),
    });
  } else {
    nextTrail[0] = {
      hex: cleanHex,
      at: new Date().toISOString(),
    };
  }
  hexMap.party = {
    hex: cleanHex,
    label: str(options.label) || current.label || "Party",
    notes: options.notes !== undefined ? str(options.notes) : current.notes,
    updatedAt: new Date().toISOString(),
    trail: nextTrail.slice(0, 24),
  };
  state.hexMap = normalizeHexMapState(hexMap);
  return getHexMapParty(hexMap);
}

function getHexLinkedLocations(hex) {
  const clean = normalizeHexCoordinate(hex);
  if (!clean) return [];
  return (state.locations || []).filter((location) => normalizeHexCoordinate(location.hex) === clean);
}

function normalizeKingdomCreationChoice(section, value) {
  const clean = str(value).trim().toLowerCase().replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
  if (!clean) return "";
  if (section === "heartlands") {
    if (/(grassland|plain|plains|hill|hills)/.test(clean)) return "hill or plain";
    if (/(forest|swamp)/.test(clean)) return "forest or swamp";
    if (/(lake|river)/.test(clean)) return "lake or river";
    if (/(mountain|mountains|ruins|ruin)/.test(clean)) return "mountain or ruins";
  }
  if (section === "governments") {
    if (/(council|council rule|ruling council)/.test(clean)) return "oligarchy";
  }
  return clean;
}

function buildKingdomSkillRanks(trainedSkills = [], existing = {}) {
  const out = Object.fromEntries(KINGDOM_SKILL_DEFINITIONS.map((entry) => [entry.name, "untrained"]));
  if (existing && typeof existing === "object" && !Array.isArray(existing)) {
    for (const [skillName, rank] of Object.entries(existing)) {
      const cleanSkill = canonicalizeKingdomSkillName(skillName);
      if (!cleanSkill) continue;
      out[cleanSkill] = getKingdomSkillRank(rank);
    }
  }
  for (const skillName of Array.isArray(trainedSkills) ? trainedSkills : parseSkillList(trainedSkills, true)) {
    const cleanSkill = canonicalizeKingdomSkillName(skillName);
    if (!cleanSkill) continue;
    if (getKingdomSkillRankWeight(out[cleanSkill]) < getKingdomSkillRankWeight("trained")) {
      out[cleanSkill] = "trained";
    }
  }
  return out;
}

function applyQuickEditToSkillRanks(trainedSkills = [], existing = {}) {
  const out = buildKingdomSkillRanks([], existing);
  const wanted = new Set(Array.isArray(trainedSkills) ? trainedSkills.map((entry) => canonicalizeKingdomSkillName(entry)).filter(Boolean) : []);
  for (const definition of KINGDOM_SKILL_DEFINITIONS) {
    if (out[definition.name] === "trained" && !wanted.has(definition.name)) {
      out[definition.name] = "untrained";
    }
  }
  for (const skillName of wanted) {
    if (getKingdomSkillRankWeight(out[skillName]) < getKingdomSkillRankWeight("trained")) {
      out[skillName] = "trained";
    }
  }
  return out;
}

function getKingdomUntrainedImprovisationBonus(level) {
  const normalizedLevel = Math.max(1, Number.parseInt(String(level || "1"), 10) || 1);
  if (normalizedLevel >= 7) {
    return {
      bonus: normalizedLevel,
      label: `full level (+${normalizedLevel}) from Untrained Improvisation`
    };
  }
  if (normalizedLevel >= 2) {
    const bonus = Math.floor(normalizedLevel / 2);
    return {
      bonus,
      label: `half level (+${bonus}) from Untrained Improvisation`
    };
  }
  return {
    bonus: 0,
    label: "no proficiency bonus while untrained"
  };
}

function getKingdomSkillProficiency(level, rank) {
  const cleanRank = getKingdomSkillRank(rank);
  const normalizedLevel = Math.max(1, Number.parseInt(String(level || "1"), 10) || 1);
  if (cleanRank === "untrained") {
    const untrained = getKingdomUntrainedImprovisationBonus(normalizedLevel);
    return {
      bonus: untrained.bonus,
      label: untrained.label
    };
  }
  const base = {
    trained: 2,
    expert: 4,
    master: 6,
    legendary: 8
  }[cleanRank];
  const bonus = normalizedLevel + base;
  return {
    bonus,
    label: `${KINGDOM_SKILL_RANK_LABELS[cleanRank]} proficiency (+${base}) + level ${normalizedLevel}`
  };
}

function getLeaderName(leader) {
  const cleanName = str(leader?.name).trim();
  if (cleanName) return cleanName;
  return str(leader?.role) || "Unnamed leader";
}

function isActiveKingdomLeader(leader) {
  const cleanName = str(leader?.name).trim();
  return Boolean(cleanName && !/^unassigned$/i.test(cleanName));
}

function getLeaderSpecializedKingdomSkills(leader) {
  return parseSkillList(leader?.specializedSkills, true);
}

function getLeaderLeadershipBonus(leader) {
  return Math.max(0, Math.min(4, Number.parseInt(String(leader?.leadershipBonus || "0"), 10) || 0));
}

function formatSignedNumber(value) {
  const number = Number.parseInt(String(value || "0"), 10) || 0;
  return number >= 0 ? `+${number}` : String(number);
}

function getBestLeaderForKingdomSkill(kingdom, skillName) {
  const leaders = (kingdom?.leaders || []).filter(isActiveKingdomLeader);
  if (!leaders.length) {
    return {
      leader: null,
      bonus: 0,
      mode: "none",
      label: "No active leader assigned"
    };
  }
  const specialized = leaders
    .filter((leader) => getLeaderSpecializedKingdomSkills(leader).includes(skillName))
    .sort((a, b) => getLeaderLeadershipBonus(b) - getLeaderLeadershipBonus(a));
  if (specialized.length) {
    const leader = specialized[0];
    const bonus = getLeaderLeadershipBonus(leader);
    return {
      leader,
      bonus,
      mode: "specialized",
      label: `${getLeaderName(leader)} (${leader.role || "Leader"}) full Leadership Bonus`
    };
  }
  const fallback = [...leaders].sort((a, b) => getLeaderLeadershipBonus(b) - getLeaderLeadershipBonus(a))[0];
  const fallbackBonus = Math.floor(getLeaderLeadershipBonus(fallback) / 2);
  return {
    leader: fallback,
    bonus: fallbackBonus,
    mode: "fallback",
    label: `${getLeaderName(fallback)} (${fallback.role || "Leader"}) half Leadership Bonus`
  };
}

function getSettlementActionCount(settlement) {
  return KINGDOM_SETTLEMENT_ACTIONS[str(settlement?.civicStructure)] || 0;
}

function getKingdomCreationReference(profile) {
  return profile?.creationReference && typeof profile.creationReference === "object" ? profile.creationReference : {};
}

function getKingdomCreationOption(profile, section, selectedValue) {
  const clean = normalizeKingdomCreationChoice(section, selectedValue);
  if (!clean) return null;
  return getKingdomCreationOptions(profile, section).find((entry) => normalizeKingdomCreationChoice(section, entry?.name) === clean) || null;
}

function countFreeAbilityBoostSlots(entry) {
  return (entry?.abilityBoosts || []).filter((boost) => /free ability boost/i.test(str(boost))).length;
}

function calculateKingdomCreationPlan(kingdom, profile) {
  const creation = normalizeKingdomCreationState(kingdom?.creation);
  const charter = getKingdomCreationOption(profile, "charters", kingdom?.charter);
  const government = getKingdomCreationOption(profile, "governments", kingdom?.government);
  const heartland = getKingdomCreationOption(profile, "heartlands", kingdom?.heartland);
  const reference = getKingdomCreationReference(profile);
  const abilityAdjustments = Object.fromEntries(KINGDOM_ABILITY_KEYS.map((key) => [key, 0]));
  const trainedSkills = [];
  const entries = [charter, government, heartland].filter(Boolean);
  const addAbilityAdjustment = (value, delta) => {
    const clean = canonicalizeKingdomAbilityName(value);
    if (!clean) return;
    abilityAdjustments[clean] += delta;
  };
  const addTrainedSkill = (value) => {
    const clean = canonicalizeKingdomSkillName(value);
    if (!clean) return;
    trainedSkills.push(clean);
  };
  for (const entry of entries) {
    for (const boost of entry?.abilityBoosts || []) {
      if (/free ability boost/i.test(str(boost))) continue;
      addAbilityAdjustment(boost, 1);
    }
    if (entry?.abilityFlaw && str(entry.abilityFlaw).toLowerCase() !== "none") {
      addAbilityAdjustment(entry.abilityFlaw, -1);
    }
    for (const skill of entry?.trainedSkills || []) {
      addTrainedSkill(skill);
    }
  }
  for (const boost of creation.freeAbilityBoosts || []) {
    addAbilityAdjustment(boost, 1);
  }
  addTrainedSkill(creation.charterSkill);
  addTrainedSkill(creation.heartlandSkill);
  for (const skill of creation.bonusSkills || []) addTrainedSkill(skill);
  const expectedFreeBoosts =
    Number.parseInt(String(reference?.finalizeAbilityBoosts || "0"), 10) +
    countFreeAbilityBoostSlots(charter) +
    countFreeAbilityBoostSlots(government) +
    countFreeAbilityBoostSlots(heartland);
  const expectedFreeSkills =
    Number.parseInt(String(reference?.charterFreeSkillChoices || "0"), 10) +
    Number.parseInt(String(reference?.heartlandFreeSkillChoices || "0"), 10) +
    Number.parseInt(String(reference?.additionalTrainedSkills || "0"), 10);
  return {
    charter,
    government,
    heartland,
    creation,
    abilityAdjustments,
    trainedSkills: [...new Set(trainedSkills)],
    expectedFreeBoosts,
    expectedFreeSkills
  };
}

function applyKingdomCreationChoicesToState(kingdom, profile) {
  const plan = calculateKingdomCreationPlan(kingdom, profile);
  kingdom.abilities = { ...plan.abilityAdjustments };
  kingdom.skillRanks = buildKingdomSkillRanks(plan.trainedSkills);
  kingdom.trainedSkills = [...plan.trainedSkills];
  return plan;
}

function calculateKingdomDerivedState(kingdom, profile) {
  const activeProfile = profile || getActiveKingdomProfile();
  const normalizedKingdom = normalizeKingdomState(kingdom);
  const activeLeaders = normalizedKingdom.leaders.filter(isActiveKingdomLeader);
  const pcLeaders = activeLeaders.filter((leader) => str(leader?.type).toUpperCase() === "PC");
  const npcLeaders = activeLeaders.filter((leader) => str(leader?.type).toUpperCase() !== "PC");
  const settlementActionDetails = normalizedKingdom.settlements
    .map((settlement) => ({
      settlement,
      actions: getSettlementActionCount(settlement)
    }))
    .filter((entry) => entry.actions > 0);
  const settlementConsumption = normalizedKingdom.settlements.reduce(
    (sum, settlement) => sum + Math.max(0, Number.parseInt(String(settlement?.consumption || "0"), 10) || 0),
    0
  );
  const totalConsumption = normalizedKingdom.consumption + settlementConsumption;
  const recommendedControlDC = getControlDcForLevel(activeProfile, normalizedKingdom.level);
  const cityResourceDice = normalizedKingdom.settlements.filter((settlement) => ["City", "Metropolis"].includes(str(settlement?.size))).length;
  const settlementResourceDice = normalizedKingdom.settlements.reduce(
    (sum, settlement) => sum + Math.max(0, Number.parseInt(String(settlement?.resourceDice || "0"), 10) || 0),
    0
  );
  const skillRows = KINGDOM_SKILL_DEFINITIONS.map((definition) => {
    const rank = getKingdomSkillRank(normalizedKingdom?.skillRanks?.[definition.name]);
    const proficiency = getKingdomSkillProficiency(normalizedKingdom.level, rank);
    const leader = getBestLeaderForKingdomSkill(normalizedKingdom, definition.name);
    const abilityScore = Number.parseInt(String(normalizedKingdom?.abilities?.[definition.ability] || "0"), 10) || 0;
    return {
      skill: definition.name,
      abilityKey: definition.ability,
      abilityLabel: `${definition.ability[0].toUpperCase()}${definition.ability.slice(1)}`,
      abilityScore,
      rank,
      rankLabel: KINGDOM_SKILL_RANK_LABELS[rank],
      proficiencyBonus: proficiency.bonus,
      proficiencyLabel: proficiency.label,
      leaderBonus: leader.bonus,
      leaderLabel: leader.label,
      leaderMode: leader.mode,
      leaderName: leader.leader ? getLeaderName(leader.leader) : "",
      leaderRole: str(leader?.leader?.role),
      totalModifier: abilityScore + proficiency.bonus + leader.bonus
    };
  }).sort((a, b) => a.abilityLabel.localeCompare(b.abilityLabel) || a.skill.localeCompare(b.skill));

  return {
    recommendedControlDC,
    controlDcOverride: normalizedKingdom.controlDC - recommendedControlDC,
    pcLeaderActions: pcLeaders.length * 3,
    npcLeaderActions: npcLeaders.length * 2,
    settlementActions: settlementActionDetails.reduce((sum, entry) => sum + entry.actions, 0),
    settlementActionDetails,
    totalActions: pcLeaders.length * 3 + npcLeaders.length * 2 + settlementActionDetails.reduce((sum, entry) => sum + entry.actions, 0),
    settlementConsumption,
    totalConsumption,
    foodAfterUpkeep: normalizedKingdom.commodities.food - totalConsumption,
    settlementResourceDice,
    cityResourceDice,
    highestRuin: Math.max(
      normalizedKingdom.ruin.corruption,
      normalizedKingdom.ruin.crime,
      normalizedKingdom.ruin.decay,
      normalizedKingdom.ruin.strife
    ),
    peacefulNegotiationShift: -Math.floor(normalizedKingdom.fame / 10) + Math.floor(normalizedKingdom.infamy / 10),
    hostileNegotiationShift: -Math.floor(normalizedKingdom.infamy / 10),
    trainedCount: skillRows.filter((row) => row.rank !== "untrained").length,
    advancedCount: skillRows.filter((row) => getKingdomSkillRankWeight(row.rank) > getKingdomSkillRankWeight("trained")).length,
    activeLeaderCount: activeLeaders.length,
    expectedRoleCount: Array.isArray(activeProfile?.leadershipRoles) ? activeProfile.leadershipRoles.length : 8,
    skillRows
  };
}

function getControlDcForLevel(profile, level) {
  const normalizedLevel = Math.max(1, Number.parseInt(String(level || "1"), 10) || 1);
  const table = Array.isArray(profile?.advancement) ? profile.advancement : [];
  return table.find((entry) => Number.parseInt(String(entry?.level || "0"), 10) === normalizedLevel)?.controlDC || 14;
}

function createStarterKingdomState() {
  const profile = getKingdomProfileById(getDefaultKingdomProfileId());
  const kingdom = {
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
    skillRanks: buildKingdomSkillRanks(["Agriculture", "Politics", "Trade", "Wilderness"]),
    creation: createStarterKingdomCreationState(),
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
  applyKingdomCreationChoicesToState(kingdom, profile);
  return kingdom;
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
  const normalizedTrainedSkills = Array.isArray(out.trainedSkills)
    ? out.trainedSkills.map((skill) => str(skill)).filter(Boolean)
    : str(out.trainedSkills)
        .split(",")
        .map((skill) => str(skill).trim())
        .filter(Boolean);
  out.creation = normalizeKingdomCreationState(out.creation);
  out.skillRanks = buildKingdomSkillRanks(normalizedTrainedSkills, out.skillRanks);
  out.trainedSkills = KINGDOM_SKILL_DEFINITIONS.map((entry) => entry.name).filter(
    (skillName) => getKingdomSkillRankWeight(out.skillRanks?.[skillName]) >= getKingdomSkillRankWeight("trained")
  );
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
  out.regions = Array.isArray(out.regions)
    ? out.regions.map((region) => ({
        ...region,
        id: str(region?.id) || uid(),
        hex: normalizeHexCoordinate(region?.hex) || str(region?.hex),
        updatedAt: str(region?.updatedAt) || "",
      }))
    : [];
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
  const derived = calculateKingdomDerivedState(data, rulesProfile);
  return {
    name: data.name,
    currentTurnLabel: data.currentTurnLabel,
    currentDate: data.currentDate,
    level: data.level,
    size: data.size,
    controlDC: data.controlDC,
    recommendedControlDC: derived.recommendedControlDC,
    resourceDie: data.resourceDie,
    resourcePoints: data.resourcePoints,
    trainedSkills: [...(data.trainedSkills || [])].slice(0, 16),
    skillRanks: { ...(data.skillRanks || {}) },
    skillModifiers: derived.skillRows.map((entry) => ({
      skill: entry.skill,
      linkedAbility: entry.abilityLabel,
      rank: entry.rankLabel,
      modifier: entry.totalModifier,
      leader: entry.leaderName || "No active leader"
    })),
    abilities: { ...(data.abilities || {}) },
    commodities: { ...(data.commodities || {}) },
    consumption: data.consumption,
    totalConsumption: derived.totalConsumption,
    foodAfterUpkeep: derived.foodAfterUpkeep,
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
  kingdom.creation = normalizeKingdomCreationState({
    ...(kingdom.creation || createStarterKingdomCreationState()),
    freeAbilityBoosts: fields.creationFreeAbilityBoosts,
    charterSkill: fields.creationCharterSkill,
    heartlandSkill: fields.creationHeartlandSkill,
    bonusSkills: fields.creationBonusSkills
  });
  kingdom.controlDC = Math.max(
    10,
    Number.parseInt(String(fields.controlDC || getControlDcForLevel(profile, kingdom.level)), 10) || getControlDcForLevel(profile, kingdom.level)
  );
  kingdom.resourceDie = ["d4", "d6", "d8", "d10", "d12"].includes(str(fields.resourceDie)) ? str(fields.resourceDie) : kingdom.resourceDie;
  kingdom.resourcePoints = Number.parseInt(String(fields.resourcePoints || kingdom.resourcePoints || "0"), 10) || 0;
  kingdom.xp = Number.parseInt(String(fields.xp || kingdom.xp || "0"), 10) || 0;
  const nextSkillRanks = applyQuickEditToSkillRanks(parseSkillList(fields.trainedSkills, true), kingdom.skillRanks);
  for (const definition of KINGDOM_SKILL_DEFINITIONS) {
    const fieldName = `skillRank__${slugify(definition.name)}`;
    if (fieldName in fields) {
      nextSkillRanks[definition.name] = getKingdomSkillRank(fields[fieldName]);
    }
  }
  kingdom.skillRanks = nextSkillRanks;
  kingdom.trainedSkills = KINGDOM_SKILL_DEFINITIONS.map((entry) => entry.name).filter(
    (skillName) => getKingdomSkillRankWeight(nextSkillRanks?.[skillName]) >= getKingdomSkillRankWeight("trained")
  );
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

function recalculateKingdomFromCreationChoices(fields) {
  applyKingdomOverviewForm(fields);
  const kingdom = getKingdomState();
  const profile = getKingdomProfileById(kingdom.profileId);
  return applyKingdomCreationChoicesToState(kingdom, profile);
}

function createKingdomLeader(fields) {
  const kingdom = getKingdomState();
  const role = str(fields.role) || "Leader";
  const roleProfile = (getActiveKingdomProfile()?.leadershipRoles || []).find((entry) => str(entry?.role) === role);
  kingdom.leaders.unshift({
    id: uid(),
    role,
    name: str(fields.name) || "Unnamed leader",
    type: str(fields.type) || "NPC",
    leadershipBonus: Number.parseInt(String(fields.leadershipBonus || "0"), 10) || 0,
    relevantSkills: str(fields.relevantSkills) || (roleProfile?.relevantSkills || []).join(", "),
    specializedSkills: str(fields.specializedSkills) || (roleProfile?.specializedSkills || []).join(", "),
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
  const hexMap = getHexMapState();
  kingdom.regions.unshift({
    id: uid(),
    hex: normalizeHexCoordinate(fields.hex, hexMap.columns, hexMap.rows) || str(fields.hex) || "Unknown hex",
    status: str(fields.status) || "Claimed",
    terrain: str(fields.terrain),
    workSite: str(fields.workSite),
    notes: str(fields.notes),
    updatedAt: new Date().toISOString()
  });
}

function applyHexMapSettings(fields) {
  const hexMap = getHexMapState();
  hexMap.mapName = str(fields.mapName) || hexMap.mapName;
  hexMap.columns = Math.max(
    HEX_MAP_COLUMNS_MIN,
    Math.min(HEX_MAP_COLUMNS_MAX, Number.parseInt(String(fields.columns || hexMap.columns), 10) || hexMap.columns)
  );
  hexMap.rows = Math.max(
    HEX_MAP_ROWS_MIN,
    Math.min(HEX_MAP_ROWS_MAX, Number.parseInt(String(fields.rows || hexMap.rows), 10) || hexMap.rows)
  );
  hexMap.hexSize = Math.max(
    HEX_MAP_HEX_SIZE_MIN,
    Math.min(HEX_MAP_HEX_SIZE_MAX, Number.parseInt(String(fields.hexSize || hexMap.hexSize), 10) || hexMap.hexSize)
  );
  hexMap.backgroundOpacity = Math.max(
    0,
    Math.min(0.95, Number.parseFloat(String(fields.backgroundOpacity ?? hexMap.backgroundOpacity)) || hexMap.backgroundOpacity)
  );
  hexMap.backgroundScale = Math.max(
    HEX_MAP_BACKGROUND_SCALE_MIN,
    Math.min(
      HEX_MAP_BACKGROUND_SCALE_MAX,
      Number.parseFloat(String(fields.backgroundScale ?? hexMap.backgroundScale)) || hexMap.backgroundScale
    )
  );
  hexMap.backgroundOffsetX = Number.isFinite(Number.parseFloat(String(fields.backgroundOffsetX)))
    ? Number.parseFloat(String(fields.backgroundOffsetX))
    : hexMap.backgroundOffsetX;
  hexMap.backgroundOffsetY = Number.isFinite(Number.parseFloat(String(fields.backgroundOffsetY)))
    ? Number.parseFloat(String(fields.backgroundOffsetY))
    : hexMap.backgroundOffsetY;
  hexMap.gridFillOpacity = Math.max(
    0,
    Math.min(0.65, Number.parseFloat(String(fields.gridFillOpacity ?? hexMap.gridFillOpacity)) || hexMap.gridFillOpacity)
  );
  hexMap.gridLineOpacity = Math.max(
    0.15,
    Math.min(1, Number.parseFloat(String(fields.gridLineOpacity ?? hexMap.gridLineOpacity)) || hexMap.gridLineOpacity)
  );
  hexMap.showLabels = String(fields.showLabels || "true") !== "false";
  clampHexMapPan(hexMap);
  state.hexMap = normalizeHexMapState(hexMap);
  setHexMapSelectedHex(ui.hexMapSelectedHex || "A1");
}

function upsertHexMapRegion(fields) {
  const kingdom = getKingdomState();
  const hexMap = getHexMapState();
  const hex = normalizeHexCoordinate(fields.hex, hexMap.columns, hexMap.rows);
  if (!hex) throw new Error("Select a valid hex first.");
  const existing = kingdom.regions.find((region) => normalizeHexCoordinate(region.hex) === hex);
  const patch = {
    hex,
    status: str(fields.status) || "Claimed",
    terrain: str(fields.terrain),
    workSite: str(fields.workSite),
    notes: str(fields.notes),
    updatedAt: new Date().toISOString(),
  };
  if (existing) {
    Object.assign(existing, patch);
    return existing;
  }
  const created = {
    id: uid(),
    ...patch,
  };
  kingdom.regions.unshift(created);
  return created;
}

function clearHexMapRegion(hex) {
  const kingdom = getKingdomState();
  const clean = normalizeHexCoordinate(hex);
  if (!clean) return false;
  const before = kingdom.regions.length;
  kingdom.regions = kingdom.regions.filter((region) => normalizeHexCoordinate(region.hex) !== clean);
  return kingdom.regions.length !== before;
}

function createHexMapMarker(fields) {
  const hexMap = getHexMapState();
  const hex = normalizeHexCoordinate(fields.hex, hexMap.columns, hexMap.rows);
  if (!hex) throw new Error("Select a valid hex first.");
  const marker = {
    id: uid(),
    hex,
    type: HEX_MAP_MARKER_TYPES.includes(str(fields.type)) ? str(fields.type) : "Note",
    title: str(fields.title) || "Untitled marker",
    notes: str(fields.notes),
    updatedAt: new Date().toISOString(),
  };
  hexMap.markers.unshift(marker);
  state.hexMap = normalizeHexMapState(hexMap);
  return marker;
}

function upsertHexMapParty(fields) {
  return moveHexMapPartyToHex(fields.hex, {
    label: fields.label,
    notes: fields.notes,
  });
}

function createHexMapForce(fields) {
  const hexMap = getHexMapState();
  const hex = normalizeHexCoordinate(fields.hex, hexMap.columns, hexMap.rows);
  if (!hex) throw new Error("Select a valid hex first.");
  const force = {
    id: uid(),
    hex,
    type: HEX_MAP_FORCE_TYPES.includes(str(fields.type)) ? str(fields.type) : "Allied Force",
    name: str(fields.name) || "Unnamed force",
    notes: str(fields.notes),
    updatedAt: new Date().toISOString(),
  };
  hexMap.forces.unshift(force);
  state.hexMap = normalizeHexMapState(hexMap);
  return force;
}

function appendHexMapAiNote(text) {
  const hex = getHexMapSelectedHex();
  const region = getKingdomRegionByHex(hex);
  const stamp = new Date().toLocaleString();
  const block = `[AI ${stamp}]\n${str(text)}`;
  if (region) {
    region.notes = region.notes ? `${region.notes}\n\n${block}`.trim() : block;
    region.updatedAt = new Date().toISOString();
  } else {
    createKingdomRegion({
      hex,
      status: "Reconnoitered",
      terrain: "",
      workSite: "",
      notes: block,
    });
  }
  saveState();
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
  const derived = calculateKingdomDerivedState(kingdom, profile);
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

      ${renderKingdomDerivedPanel(kingdom, profile, derived)}

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
              ${renderKingdomCreationSelect("charter", profile, "charters", kingdom.charter, "Choose a charter")}
            </label>
            <label>Government
              ${renderKingdomCreationSelect("government", profile, "governments", kingdom.government, "Choose a government")}
            </label>
            <label>Heartland
              ${renderKingdomCreationSelect("heartland", profile, "heartlands", kingdom.heartland, "Choose a heartland")}
            </label>
          </div>
          <div data-kingdom-choice-summary>
            ${renderKingdomCreationChoiceSummary(kingdom, profile)}
          </div>
          ${renderKingdomCreationPlanner(kingdom, profile)}
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
          <label>Trained Skills (quick edit, comma separated)
            <input name="trainedSkills" value="${escapeHtml((kingdom.trainedSkills || []).join(", "))}" placeholder="Agriculture, Politics, Trade, Wilderness" />
          </label>
          <p class="small">This quick-edit list auto-syncs with the rank table below. Use the table for Expert, Master, and Legendary skills and to see the actual modifier math.</p>
          ${renderKingdomSkillMatrix(kingdom, derived)}
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

function renderKingdomDerivedPanel(kingdom, profile, derived) {
  const controlDcNote =
    derived.controlDcOverride === 0
      ? `Matches the level ${kingdom.level} profile value.`
      : `${derived.controlDcOverride > 0 ? "Override above" : "Override below"} profile default by ${Math.abs(derived.controlDcOverride)}.`;
  const settlementActionNote = derived.settlementActionDetails.length
    ? derived.settlementActionDetails.map((entry) => `${entry.settlement.name || "Settlement"} +${entry.actions}`).join(" • ")
    : "No civic structures are adding settlement actions yet.";
  return `
    <section class="kingdom-overview-grid kingdom-derived-grid">
      <article class="panel">
        <h2>Computed Summary</h2>
        <div class="card-list kingdom-derived-cards">
          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Control DC</span>
              <span class="entry-meta">${escapeHtml(String(kingdom.controlDC || 14))}</span>
            </div>
            <p>Recommended from ${escapeHtml(profile?.shortLabel || profile?.label || "profile")} at level ${escapeHtml(String(kingdom.level || 1))}: <strong>${escapeHtml(String(derived.recommendedControlDC))}</strong>.</p>
            <p class="small">${escapeHtml(controlDcNote)}</p>
          </article>
          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Action Economy</span>
              <span class="entry-meta">${escapeHtml(String(derived.totalActions))} total</span>
            </div>
            <p><strong>${escapeHtml(String(derived.pcLeaderActions))}</strong> from PC leaders, <strong>${escapeHtml(String(derived.npcLeaderActions))}</strong> from NPC leaders, and <strong>${escapeHtml(String(derived.settlementActions))}</strong> from settlements.</p>
            <p class="small">${escapeHtml(settlementActionNote)}</p>
          </article>
          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Upkeep</span>
              <span class="entry-meta">${escapeHtml(String(derived.totalConsumption))} total consumption</span>
            </div>
            <p>Kingdom base consumption <strong>${escapeHtml(String(kingdom.consumption || 0))}</strong> + settlement consumption <strong>${escapeHtml(String(derived.settlementConsumption))}</strong>.</p>
            <p class="small">Food after upkeep preview: <strong>${escapeHtml(String(derived.foodAfterUpkeep))}</strong>.</p>
          </article>
          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Skill Coverage</span>
              <span class="entry-meta">${escapeHtml(String(derived.trainedCount))} trained+</span>
            </div>
            <p><strong>${escapeHtml(String(derived.advancedCount))}</strong> skills are above Trained. Active leaders assigned: <strong>${escapeHtml(String(derived.activeLeaderCount))}</strong> / ${escapeHtml(String(derived.expectedRoleCount))} expected roles.</p>
            <p class="small">If a skill has no specialized leader, the sheet falls back to half of the best Leadership Bonus.</p>
          </article>
          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Negotiation Pressure</span>
              <span class="entry-meta">Fame ${escapeHtml(String(kingdom.fame || 0))} • Infamy ${escapeHtml(String(kingdom.infamy || 0))}</span>
            </div>
            <p>Peaceful negotiation DC shift: <strong>${escapeHtml(formatSignedNumber(derived.peacefulNegotiationShift))}</strong>. Hostile negotiation DC shift: <strong>${escapeHtml(formatSignedNumber(derived.hostileNegotiationShift))}</strong>.</p>
            <p class="small">Every 10 Fame helps peaceful negotiations. Every 10 Infamy helps hostile negotiations but hurts peaceful ones.</p>
          </article>
          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Ruin Watch</span>
              <span class="entry-meta">Threshold ${escapeHtml(String(kingdom.ruin.threshold || 5))}</span>
            </div>
            <p>Highest ruin track: <strong>${escapeHtml(String(derived.highestRuin))}</strong>. You have <strong>${escapeHtml(String(Math.max(0, (kingdom.ruin.threshold || 5) - derived.highestRuin)))}</strong> before the threshold is met on that track.</p>
            <p class="small">Settlement resource dice tracked: <strong>${escapeHtml(String(derived.settlementResourceDice))}</strong>. City-size bonus settlements: <strong>${escapeHtml(String(derived.cityResourceDice))}</strong>.</p>
          </article>
        </div>
      </article>
    </section>
  `;
}

function renderKingdomSkillMatrix(kingdom, derived) {
  const rows = (derived?.skillRows || [])
    .map((entry) => {
      const selectName = `skillRank__${slugify(entry.skill)}`;
      return `
        <tr>
          <td><strong>${escapeHtml(entry.skill)}</strong></td>
          <td>${escapeHtml(entry.abilityLabel)} ${escapeHtml(formatSignedNumber(entry.abilityScore))}</td>
          <td>
            <select name="${escapeHtml(selectName)}">
              ${KINGDOM_SKILL_RANKS.map(
                (rank) =>
                  `<option value="${rank}" ${entry.rank === rank ? "selected" : ""}>${escapeHtml(KINGDOM_SKILL_RANK_LABELS[rank])}</option>`
              ).join("")}
            </select>
          </td>
          <td>${escapeHtml(formatSignedNumber(entry.proficiencyBonus))}</td>
          <td>${escapeHtml(formatSignedNumber(entry.leaderBonus))}</td>
          <td><strong>${escapeHtml(formatSignedNumber(entry.totalModifier))}</strong></td>
          <td>${escapeHtml(entry.leaderLabel)}</td>
        </tr>
      `;
    })
    .join("");
  return `
    <details class="session-edit-panel kingdom-skill-panel" open>
      <summary>Computed Kingdom Skill Modifiers</summary>
      <p class="small">Final modifier = linked ability + proficiency from rank/level + best Leadership Bonus for that skill.</p>
      <div class="readable-table-wrap kingdom-skill-table-wrap">
        <table class="readable-table kingdom-skill-table">
          <thead>
            <tr>
              <th>Skill</th>
              <th>Linked Ability</th>
              <th>Rank</th>
              <th>Prof</th>
              <th>Leader</th>
              <th>Total</th>
              <th>Why</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
          </tbody>
        </table>
      </div>
    </details>
  `;
}

function getKingdomCreationOptions(profile, section) {
  return Array.isArray(profile?.creationReference?.[section]) ? profile.creationReference[section] : [];
}

function getKingdomCreationDisplayValue(profile, section, selectedValue) {
  return getKingdomCreationOption(profile, section, selectedValue)?.name || str(selectedValue);
}

function renderKingdomCreationSelect(name, profile, section, selectedValue, placeholder = "Choose...") {
  const options = getKingdomCreationOptions(profile, section);
  const displayValue = getKingdomCreationDisplayValue(profile, section, selectedValue);
  const hasExactOption = options.some((entry) => str(entry?.name) === displayValue);
  const customOption = displayValue && !hasExactOption
    ? `<option value="${escapeHtml(displayValue)}" selected>${escapeHtml(displayValue)} (custom/manual)</option>`
    : "";
  return `
    <select name="${escapeHtml(name)}" data-kingdom-preview-field="${escapeHtml(name)}">
      <option value="">${escapeHtml(placeholder)}</option>
      ${customOption}
      ${options
        .map((entry) => {
          const optionName = str(entry?.name);
          return `<option value="${escapeHtml(optionName)}" ${optionName === displayValue ? "selected" : ""}>${escapeHtml(optionName)}</option>`;
        })
        .join("")}
    </select>
  `;
}

function renderKingdomChoiceSkillSelect(name, selectedValue, placeholder = "Choose a skill") {
  const cleanValue = canonicalizeKingdomSkillName(selectedValue);
  return `
    <select name="${escapeHtml(name)}" data-kingdom-preview-field="${escapeHtml(name)}">
      <option value="">${escapeHtml(placeholder)}</option>
      ${KINGDOM_SKILL_DEFINITIONS.map((entry) => {
        const optionName = str(entry?.name);
        return `<option value="${escapeHtml(optionName)}" ${optionName === cleanValue ? "selected" : ""}>${escapeHtml(optionName)}</option>`;
      }).join("")}
    </select>
  `;
}

function renderKingdomChoiceMultiSelect(name, selectedValues, items, mapLabel, size = 4) {
  const selectedSet = new Set(Array.isArray(selectedValues) ? selectedValues.map((entry) => str(entry)) : []);
  return `
    <select name="${escapeHtml(name)}" multiple size="${escapeHtml(String(size))}" class="kingdom-multi-select" data-kingdom-preview-field="${escapeHtml(name)}">
      ${items
        .map((value) => {
          const itemValue = str(value);
          const label = mapLabel ? mapLabel(itemValue) : itemValue;
          return `<option value="${escapeHtml(itemValue)}" ${selectedSet.has(itemValue) ? "selected" : ""}>${escapeHtml(label)}</option>`;
        })
        .join("")}
    </select>
  `;
}

function getFormFields(form) {
  const formData = new FormData(form);
  const out = {};
  for (const [key, value] of formData.entries()) {
    if (key in out) {
      if (Array.isArray(out[key])) {
        out[key].push(value);
      } else {
        out[key] = [out[key], value];
      }
    } else {
      out[key] = value;
    }
  }
  return out;
}

function buildKingdomCreationPreviewDraft(fields) {
  const current = getKingdomState();
  const profile = getKingdomProfileById(str(fields.profileId) || current.profileId || getDefaultKingdomProfileId());
  const draft = normalizeKingdomState({
    ...current,
    profileId: profile?.id || current.profileId,
    charter: str(fields.charter),
    government: str(fields.government),
    heartland: str(fields.heartland),
    creation: normalizeKingdomCreationState({
      ...(current.creation || createStarterKingdomCreationState()),
      freeAbilityBoosts: fields.creationFreeAbilityBoosts,
      charterSkill: fields.creationCharterSkill,
      heartlandSkill: fields.creationHeartlandSkill,
      bonusSkills: fields.creationBonusSkills
    })
  });
  return { kingdom: draft, profile };
}

function renderKingdomCreationPlannerChipsFromPlan(plan) {
  const freeBoostTargets = plan.expectedFreeBoosts || 0;
  const freeSkillTargets = plan.expectedFreeSkills || 0;
  return `
    <span class="chip">Expected free ability boosts ${escapeHtml(String(freeBoostTargets))}</span>
    <span class="chip">Expected free skill picks ${escapeHtml(String(freeSkillTargets))}</span>
    <span class="chip">Current free boosts entered ${escapeHtml(String((plan?.creation?.freeAbilityBoosts || []).length))}</span>
    <span class="chip">Current free skills entered ${escapeHtml(
      String([plan?.creation?.charterSkill, plan?.creation?.heartlandSkill, ...((plan?.creation?.bonusSkills || []))].filter(Boolean).length)
    )}</span>
  `;
}

function renderKingdomCreationPlannerBaselineFromPlan(plan) {
  return `Current recalculated baseline if applied now: ${escapeHtml(
    KINGDOM_ABILITY_KEYS.map((key) => `${KINGDOM_ABILITY_LABELS[key]} ${formatSignedNumber(plan.abilityAdjustments[key])}`).join(" • ")
  )}. Skills: ${escapeHtml(plan.trainedSkills.join(", ") || "None")}`;
}

function syncKingdomCreationPreview(form) {
  if (!(form instanceof HTMLFormElement) || form.dataset.form !== "kingdom-overview") return;
  const fields = getFormFields(form);
  const { kingdom, profile } = buildKingdomCreationPreviewDraft(fields);
  const plan = calculateKingdomCreationPlan(kingdom, profile);

  const summaryTarget = form.querySelector("[data-kingdom-choice-summary]");
  if (summaryTarget) {
    summaryTarget.innerHTML = renderKingdomCreationChoiceSummary(kingdom, profile);
  }

  const chipTarget = form.querySelector("[data-kingdom-creation-chips]");
  if (chipTarget) {
    chipTarget.innerHTML = renderKingdomCreationPlannerChipsFromPlan(plan);
  }

  const baselineTarget = form.querySelector("[data-kingdom-creation-baseline]");
  if (baselineTarget) {
    baselineTarget.textContent = `Current recalculated baseline if applied now: ${KINGDOM_ABILITY_KEYS.map(
      (key) => `${KINGDOM_ABILITY_LABELS[key]} ${formatSignedNumber(plan.abilityAdjustments[key])}`
    ).join(" • ")}. Skills: ${plan.trainedSkills.join(", ") || "None"}`;
  }

  const referenceTarget = appEl.querySelector("[data-kingdom-creation-reference]");
  if (referenceTarget) {
    referenceTarget.innerHTML = renderKingdomCreationReference(profile, kingdom);
  }
}

function renderKingdomCreationChoiceSummary(kingdom, profile) {
  const charter = getKingdomCreationOption(profile, "charters", kingdom?.charter);
  const government = getKingdomCreationOption(profile, "governments", kingdom?.government);
  const heartland = getKingdomCreationOption(profile, "heartlands", kingdom?.heartland);
  const cards = [
    {
      title: "Current Charter",
      selected: getKingdomCreationDisplayValue(profile, "charters", kingdom?.charter) || "Not set",
      option: charter
    },
    {
      title: "Current Government",
      selected: getKingdomCreationDisplayValue(profile, "governments", kingdom?.government) || "Not set",
      option: government
    },
    {
      title: "Current Heartland",
      selected: getKingdomCreationDisplayValue(profile, "heartlands", kingdom?.heartland) || "Not set",
      option: heartland
    }
  ];
  return `
    <div class="card-list kingdom-choice-summary-grid">
      ${cards
        .map((card) => {
          const option = card.option;
          return `
            <article class="entry">
              <div class="entry-head">
                <span class="entry-title">${escapeHtml(card.title)}</span>
                <span class="entry-meta">${escapeHtml(card.selected)}</span>
              </div>
              <p>${escapeHtml(option?.summary || "No structured rule note found for the current entry.")}</p>
              ${
                option
                  ? `<p class="small"><strong>Effects:</strong> ${escapeHtml(
                      [
                        ...(option.abilityBoosts || []).map((entry) => `Boost ${entry}`),
                        option.abilityFlaw && option.abilityFlaw !== "None" ? `Flaw ${option.abilityFlaw}` : option?.abilityFlaw === "None" ? "No ability flaw" : "",
                        ...(option.trainedSkills || []).map((entry) => `Train ${entry}`),
                        option?.bonusFeat ? `Bonus feat ${option.bonusFeat}` : ""
                      ]
                        .filter(Boolean)
                        .join(" • ")
                    )}</p>`
                  : `<p class="small">Use one of the listed profile options below if you want the sheet to explain it automatically.</p>`
              }
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderKingdomCreationPlanner(kingdom, profile) {
  const plan = calculateKingdomCreationPlan(kingdom, profile);
  return `
    <details class="session-edit-panel kingdom-creation-panel" open>
      <summary>Creation Planner</summary>
      <p class="small">Use this to apply the level-1 creation package from your selected charter, government, and heartland. The recalc button overwrites the current kingdom ability modifiers and trained-skill baseline.</p>
      <div class="kingdom-chip-row" data-kingdom-creation-chips>
        ${renderKingdomCreationPlannerChipsFromPlan(plan)}
      </div>
      <div class="row">
        <label>Free Ability Boosts
          ${renderKingdomChoiceMultiSelect(
            "creationFreeAbilityBoosts",
            kingdom?.creation?.freeAbilityBoosts || [],
            KINGDOM_ABILITY_KEYS,
            (value) => KINGDOM_ABILITY_LABELS[value] || value,
            4
          )}
        </label>
        <label>Charter Free Skill
          ${renderKingdomChoiceSkillSelect("creationCharterSkill", kingdom?.creation?.charterSkill || "", "Choose charter free skill")}
        </label>
        <label>Heartland Free Skill
          ${renderKingdomChoiceSkillSelect("creationHeartlandSkill", kingdom?.creation?.heartlandSkill || "", "Choose heartland free skill")}
        </label>
      </div>
      <label>Additional Starting Skills
        ${renderKingdomChoiceMultiSelect(
          "creationBonusSkills",
          kingdom?.creation?.bonusSkills || [],
          KINGDOM_SKILL_DEFINITIONS.map((entry) => entry.name),
          null,
          8
        )}
      </label>
      <p class="small">Tip: hold <strong>Ctrl</strong> while clicking to select multiple boosts or multiple starting skills.</p>
      <p class="small">Selections here update the preview immediately. The kingdom sheet itself changes only when you click the recalculate button.</p>
      <p class="small" data-kingdom-creation-baseline>${renderKingdomCreationPlannerBaselineFromPlan(plan)}</p>
      <div class="toolbar">
        <button class="btn btn-secondary" type="button" data-action="kingdom-recalculate-creation">Recalculate from Creation Choices</button>
      </div>
    </details>
  `;
}

function renderKingdomCreationReference(profile, kingdom) {
  const sections = [
    {
      title: "Charters",
      key: "charters",
      selected: getKingdomCreationDisplayValue(profile, "charters", kingdom?.charter)
    },
    {
      title: "Governments",
      key: "governments",
      selected: getKingdomCreationDisplayValue(profile, "governments", kingdom?.government)
    },
    {
      title: "Heartlands",
      key: "heartlands",
      selected: getKingdomCreationDisplayValue(profile, "heartlands", kingdom?.heartland)
    }
  ];
  const note = str(profile?.creationReference?.sourceNote);
  return `
    ${note ? `<p class="small">${escapeHtml(note)}</p>` : ""}
    <div class="kingdom-guide-grid">
      ${sections
        .map((section) => {
          const entries = getKingdomCreationOptions(profile, section.key);
          if (!entries.length) return "";
          return `
            <article class="entry">
              <div class="entry-head">
                <span class="entry-title">${escapeHtml(section.title)}</span>
                <span class="entry-meta">${escapeHtml(section.selected || "Not set")}</span>
              </div>
              <div class="card-list kingdom-choice-card-grid">
                ${entries
                  .map((entry) => {
                    const active =
                      normalizeKingdomCreationChoice(section.key, entry?.name) ===
                      normalizeKingdomCreationChoice(section.key, section.selected);
                    const details = [
                      ...(entry?.abilityBoosts || []).map((item) => `Boost ${item}`),
                      entry?.abilityFlaw && entry.abilityFlaw !== "None" ? `Flaw ${entry.abilityFlaw}` : entry?.abilityFlaw === "None" ? "No ability flaw" : "",
                      ...(entry?.trainedSkills || []).map((item) => `Train ${item}`),
                      entry?.bonusFeat ? `Bonus feat ${entry.bonusFeat}` : ""
                    ].filter(Boolean);
                    return `
                      <article class="entry kingdom-choice-card ${active ? "is-selected" : ""}">
                        <div class="entry-head">
                          <span class="entry-title">${escapeHtml(entry?.name || "Option")}</span>
                          ${active ? `<span class="entry-meta">Selected</span>` : ""}
                        </div>
                        <p>${escapeHtml(entry?.summary || "")}</p>
                        ${details.length ? `<p class="small">${escapeHtml(details.join(" • "))}</p>` : ""}
                        <div class="toolbar kingdom-choice-card-toolbar">
                          <button
                            class="btn btn-secondary"
                            type="button"
                            data-action="kingdom-use-creation-choice"
                            data-section="${escapeHtml(section.key)}"
                            data-value="${escapeHtml(entry?.name || "")}"
                          >
                            ${active ? "Using This Option" : "Use This"}
                          </button>
                        </div>
                      </article>
                    `;
                  })
                  .join("")}
              </div>
            </article>
          `;
        })
        .join("")}
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
    renderKingdomGuideSection(
      "Charter / Government / Heartland Reference",
      `<div data-kingdom-creation-reference>${renderKingdomCreationReference(profile, kingdom)}</div>`
    ),
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
  const active = isActiveKingdomLeader(leader);
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(leader.name || "Unnamed leader")}</span>
        <span class="entry-meta">${escapeHtml(leader.role || "Role")} • ${escapeHtml(leader.type || "NPC")} • +${escapeHtml(String(leader.leadershipBonus || 0))} • ${
          active ? "active" : "vacant"
        }</span>
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
  const settlementActions = getSettlementActionCount(settlement);
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(settlement.name || "Unnamed settlement")}</span>
        <span class="entry-meta">${escapeHtml(settlement.size || "Settlement")} • influence ${escapeHtml(String(settlement.influence || 0))} • ${
          settlementActions ? `${settlementActions} settlement action${settlementActions === 1 ? "" : "s"}` : "no civic actions"
        }</span>
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

function renderHexMap() {
  const hexMap = getHexMapState();
  const selectedHex = getHexMapSelectedHex(hexMap);
  const region = getKingdomRegionByHex(selectedHex);
  const markers = getHexMapMarkersForHex(selectedHex);
  const forces = getHexMapForcesForHex(selectedHex);
  const party = getHexMapParty(hexMap);
  const linkedLocations = getHexLinkedLocations(selectedHex);
  const view = getHexMapViewBox(hexMap);
  const metrics = getHexMapMetrics(hexMap);
  const backgroundPlacement = hexMap.backgroundUrl ? getHexMapBackgroundPlacement(hexMap) : null;

  return `
    <div class="page-stack hexmap-page">
      ${renderPageIntro("Hex Map", "Lay out claimed hexes, building sites, encounters, and future kingdom events in one visual board.")}
      <section class="panel flow-panel">
        <h2>How To Use This Map</h2>
        <ol class="flow-list">
          <li><strong>Step 1:</strong> optionally load your own Stolen Lands image as a background.</li>
          <li><strong>Step 2:</strong> click any hex to inspect it and save kingdom-region details.</li>
          <li><strong>Step 3:</strong> pin encounters, buildings, settlements, or events as map markers.</li>
        </ol>
        ${ui.hexMapMessage ? `<p class="small">${escapeHtml(ui.hexMapMessage)}</p>` : ""}
      </section>

      <section class="hexmap-layout">
        <article class="panel hexmap-board-panel">
          <div class="entry-head">
            <div>
              <h2 style="margin:0;">${escapeHtml(hexMap.mapName || "Hex Map")}</h2>
              <p class="small" style="margin:6px 0 0;">${escapeHtml(
                `${metrics.columns} columns • ${metrics.rows} rows • hex size ${metrics.size}px • selected ${selectedHex}`
              )}</p>
            </div>
            <div class="hexmap-chip-row">
              <span class="chip">${escapeHtml(String(getKingdomState().regions.length))} tracked regions</span>
              <span class="chip">${escapeHtml(String(hexMap.markers.length))} markers</span>
              <span class="chip">${escapeHtml(String(hexMap.forces.length || 0))} force markers</span>
              <span class="chip">${escapeHtml(party.hex ? `${party.label} at ${party.hex}` : "Party marker not placed")}</span>
              <span class="chip ${hexMap.partyMoveMode ? "chip-accent" : ""}">${escapeHtml(hexMap.partyMoveMode ? "Party Move Mode: On" : "Party Move Mode: Off")}</span>
              <span class="chip" data-hexmap-zoom-chip>Zoom ${escapeHtml(`${Math.round(hexMap.zoom * 100)}%`)}</span>
            </div>
          </div>

          <form class="hexmap-settings-grid" data-form="hexmap-settings">
            <label>Map Name
              <input name="mapName" value="${escapeHtml(hexMap.mapName || "")}" placeholder="Stolen Lands Hex Planner" />
            </label>
            <label>Columns
              <input name="columns" type="number" min="${HEX_MAP_COLUMNS_MIN}" max="${HEX_MAP_COLUMNS_MAX}" value="${escapeHtml(String(hexMap.columns))}" />
            </label>
            <label>Rows
              <input name="rows" type="number" min="${HEX_MAP_ROWS_MIN}" max="${HEX_MAP_ROWS_MAX}" value="${escapeHtml(String(hexMap.rows))}" />
            </label>
            <label>Hex Size
              <input name="hexSize" type="number" min="${HEX_MAP_HEX_SIZE_MIN}" max="${HEX_MAP_HEX_SIZE_MAX}" value="${escapeHtml(String(hexMap.hexSize))}" />
            </label>
            <label>Background Opacity
              <input name="backgroundOpacity" type="number" min="0" max="0.95" step="0.05" value="${escapeHtml(String(hexMap.backgroundOpacity))}" />
            </label>
            <label>Background Scale
              <input name="backgroundScale" type="number" min="${HEX_MAP_BACKGROUND_SCALE_MIN}" max="${HEX_MAP_BACKGROUND_SCALE_MAX}" step="0.05" value="${escapeHtml(String(hexMap.backgroundScale))}" />
            </label>
            <label>Background Offset X
              <input name="backgroundOffsetX" type="number" step="10" value="${escapeHtml(String(hexMap.backgroundOffsetX))}" />
            </label>
            <label>Background Offset Y
              <input name="backgroundOffsetY" type="number" step="10" value="${escapeHtml(String(hexMap.backgroundOffsetY))}" />
            </label>
            <label>Grid Fill Opacity
              <input name="gridFillOpacity" type="number" min="0" max="0.65" step="0.05" value="${escapeHtml(String(hexMap.gridFillOpacity))}" />
            </label>
            <label>Grid Line Opacity
              <input name="gridLineOpacity" type="number" min="0.15" max="1" step="0.05" value="${escapeHtml(String(hexMap.gridLineOpacity))}" />
            </label>
            <label>Show Labels
              <select name="showLabels">
                <option value="true" ${hexMap.showLabels ? "selected" : ""}>On</option>
                <option value="false" ${hexMap.showLabels ? "" : "selected"}>Off</option>
              </select>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Save Map Settings</button>
            </div>
          </form>

          <div class="toolbar hexmap-toolbar">
            <button class="btn btn-secondary" type="button" data-action="hexmap-choose-background">Choose Background</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-clear-background" ${hexMap.backgroundUrl ? "" : "disabled"}>Clear Background</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-fit-background" ${hexMap.backgroundUrl ? "" : "disabled"}>Recenter Background</button>
            <button class="btn ${hexMap.partyMoveMode ? "btn-primary" : "btn-secondary"}" type="button" data-action="hexmap-toggle-party-move">${hexMap.partyMoveMode ? "Party Move Mode On" : "Party Move Mode Off"}</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-zoom-out">Zoom Out</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-zoom-in">Zoom In</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-reset-view">Reset View</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-pan" data-direction="left">Pan Left</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-pan" data-direction="right">Pan Right</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-pan" data-direction="up">Pan Up</button>
            <button class="btn btn-secondary" type="button" data-action="hexmap-pan" data-direction="down">Pan Down</button>
          </div>
          <p class="small">Background: ${escapeHtml(
            hexMap.backgroundName || "None loaded. Bring your own local Stolen Lands image when ready."
          )}</p>
          ${
            backgroundPlacement
              ? `<p class="small">Map image ${escapeHtml(
                  `${Math.round(backgroundPlacement.naturalWidth)}x${Math.round(backgroundPlacement.naturalHeight)} px`
                )}. Use scale and offsets until the overlay hexes sit on the printed hex centers.</p>`
              : `<p class="small">Tip: save the Stolen Lands map to disk, click <strong>Choose Background</strong>, then line it up with the overlay using the scale and offset controls.</p>`
          }

          <div class="hexmap-stage-shell" data-hexmap-stage-shell="true">
            <svg
              class="hexmap-stage"
              data-hexmap-stage="true"
              viewBox="${escapeHtml(`${view.x.toFixed(2)} ${view.y.toFixed(2)} ${view.width.toFixed(2)} ${view.height.toFixed(2)}`)}"
              xmlns="http://www.w3.org/2000/svg"
              role="img"
              aria-label="Interactive kingdom hex map"
            >
              <rect x="0" y="0" width="${escapeHtml(String(view.boardWidth.toFixed(2)))}" height="${escapeHtml(String(view.boardHeight.toFixed(2)))}" fill="#f7f1e5" />
              ${
                hexMap.backgroundUrl && backgroundPlacement
                  ? `<image
                      href="${escapeHtml(hexMap.backgroundUrl)}"
                      x="${escapeHtml(String(backgroundPlacement.x.toFixed(2)))}"
                      y="${escapeHtml(String(backgroundPlacement.y.toFixed(2)))}"
                      width="${escapeHtml(String(backgroundPlacement.width.toFixed(2)))}"
                      height="${escapeHtml(String(backgroundPlacement.height.toFixed(2)))}"
                      preserveAspectRatio="none"
                      opacity="${escapeHtml(String(hexMap.backgroundOpacity))}"
                    />`
                  : ""
              }
              ${renderHexMapPartyTrail(hexMap)}
              ${renderHexMapSvgCells(hexMap, selectedHex)}
            </svg>
          </div>

          <div class="hexmap-legend-grid">
            <article class="entry">
              <div class="entry-head">
                <span class="entry-title">Hex Status Colors</span>
              </div>
              <div class="hexmap-legend-row">
                ${HEX_MAP_STATUS_OPTIONS.map(
                  (status) => `
                    <span class="hexmap-legend-chip">
                      <span class="hexmap-legend-swatch" style="background:${escapeHtml(getHexStatusColor(status))};"></span>
                      ${escapeHtml(status)}
                    </span>
                  `
                ).join("")}
              </div>
            </article>
            <article class="entry">
              <div class="entry-head">
                <span class="entry-title">Marker Types</span>
              </div>
              <div class="hexmap-legend-row">
                <span class="hexmap-legend-chip">
                  <span class="hexmap-party-dot">P</span>
                  Party
                </span>
                ${HEX_MAP_MARKER_TYPES.map(
                  (type) => `
                    <span class="hexmap-legend-chip">
                      <span class="hexmap-marker-dot" style="background:${escapeHtml(getHexMarkerColor(type))};"></span>
                      ${escapeHtml(type)}
                    </span>
                  `
                ).join("")}
              </div>
            </article>
            <article class="entry">
              <div class="entry-head">
                <span class="entry-title">Force Types</span>
              </div>
              <div class="hexmap-legend-row">
                ${HEX_MAP_FORCE_TYPES.map(
                  (type) => `
                    <span class="hexmap-legend-chip">
                      <span class="hexmap-party-dot" style="background:${escapeHtml(getHexForceColor(type))};">${escapeHtml(getHexForceGlyph(type))}</span>
                      ${escapeHtml(type)}
                    </span>
                  `
                ).join("")}
              </div>
            </article>
          </div>
        </article>

        <aside class="panel hexmap-inspector">
          <div class="entry-head">
            <span class="entry-title">Selected Hex ${escapeHtml(selectedHex)}</span>
            <span class="entry-meta">${escapeHtml(region?.status || "No region record yet")}</span>
          </div>

          <div class="hexmap-summary-grid">
            <article class="entry">
              <div class="entry-head"><span class="entry-title">Terrain</span></div>
              <p>${escapeHtml(region?.terrain || "Not set")}</p>
            </article>
            <article class="entry">
              <div class="entry-head"><span class="entry-title">Work Site</span></div>
              <p>${escapeHtml(region?.workSite || "None")}</p>
            </article>
            <article class="entry">
              <div class="entry-head"><span class="entry-title">Markers</span></div>
              <p>${escapeHtml(String(markers.length))}</p>
            </article>
            <article class="entry">
              <div class="entry-head"><span class="entry-title">Forces</span></div>
              <p>${escapeHtml(String(forces.length))}</p>
            </article>
            <article class="entry">
              <div class="entry-head"><span class="entry-title">Party</span></div>
              <p>${escapeHtml(party.hex ? `${party.label} at ${party.hex}` : "Not placed")}</p>
            </article>
            <article class="entry">
              <div class="entry-head"><span class="entry-title">Locations</span></div>
              <p>${escapeHtml(String(linkedLocations.length))}</p>
            </article>
          </div>

          <form data-form="hexmap-region">
            <input type="hidden" name="hex" value="${escapeHtml(selectedHex)}" />
            <div class="row">
              <label>Hex
                <input name="hexReadOnly" value="${escapeHtml(selectedHex)}" readonly />
              </label>
              <label>Status
                <select name="status">
                  ${HEX_MAP_STATUS_OPTIONS.map((status) => `<option value="${escapeHtml(status)}" ${str(region?.status) === status ? "selected" : ""}>${escapeHtml(status)}</option>`).join("")}
                </select>
              </label>
            </div>
            <div class="row">
              <label>Terrain
                <input name="terrain" list="hexmap-terrain-options" value="${escapeHtml(region?.terrain || "")}" placeholder="Forest" />
              </label>
              <label>Work Site
                <input name="workSite" value="${escapeHtml(region?.workSite || "")}" placeholder="Lumber Camp" />
              </label>
            </div>
            <label>Hex Notes
              <textarea name="notes" placeholder="Why this hex matters, what lives here, and what changes next.">${escapeHtml(region?.notes || "")}</textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Save Hex Record</button>
              <button class="btn btn-danger" type="button" data-action="hexmap-clear-region" ${region ? "" : "disabled"}>Clear Hex Record</button>
            </div>
          </form>

          <form data-form="hexmap-marker">
            <input type="hidden" name="hex" value="${escapeHtml(selectedHex)}" />
            <div class="row">
              <label>Marker Type
                <select name="type">
                  ${HEX_MAP_MARKER_TYPES.map((type) => `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`).join("")}
                </select>
              </label>
              <label>Marker Title
                <input name="title" placeholder="Bandit scouting camp" />
              </label>
            </div>
            <label>Marker Notes
              <textarea name="notes" placeholder="What happens here, what clue is present, what escalates if ignored."></textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Add Marker</button>
            </div>
          </form>

          <form data-form="hexmap-party">
            <input type="hidden" name="hex" value="${escapeHtml(selectedHex)}" />
            <div class="entry-head">
              <span class="entry-title">Party Tracker</span>
              <span class="entry-meta">${escapeHtml(party.hex || "Not placed")}</span>
            </div>
            <div class="row">
              <label>Party Label
                <input name="label" value="${escapeHtml(party.label || "Party")}" placeholder="Party" />
              </label>
              <label>Current Party Hex
                <input value="${escapeHtml(party.hex || "Not placed")}" readonly />
              </label>
            </div>
            <label>Party Notes
              <textarea name="notes" placeholder="What they are doing here, current objective, who they are traveling with.">${escapeHtml(party.notes || "")}</textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Move Party To ${escapeHtml(selectedHex)}</button>
              <button class="btn btn-secondary" type="button" data-action="hexmap-center-party" ${party.hex ? "" : "disabled"}>Center On Party</button>
              <button class="btn btn-secondary" type="button" data-action="hexmap-toggle-party-move">${hexMap.partyMoveMode ? "Stop Click-Move" : "Enable Click-Move"}</button>
              <button class="btn btn-secondary" type="button" data-action="hexmap-clear-party-trail" ${party.trail.length > 1 ? "" : "disabled"}>Clear Trail</button>
              <button class="btn btn-danger" type="button" data-action="hexmap-clear-party" ${party.hex ? "" : "disabled"}>Clear Party Marker</button>
            </div>
          </form>

          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Recent Party Route</span>
              <span class="entry-meta">${escapeHtml(String(party.trail.length))} points</span>
            </div>
            ${
              party.trail.length
                ? `<ul class="flow-list">${party.trail
                    .slice(0, 10)
                    .map((entry, index) => `<li><strong>${escapeHtml(entry.hex)}</strong>${index === 0 ? " (current/latest)" : ""}${entry.at ? ` • ${escapeHtml(formatTimestamp(entry.at))}` : ""}</li>`)
                    .join("")}</ul>`
                : `<p class="small">No party travel history yet.</p>`
            }
          </article>

          <form data-form="hexmap-force">
            <input type="hidden" name="hex" value="${escapeHtml(selectedHex)}" />
            <div class="entry-head">
              <span class="entry-title">Force Marker</span>
              <span class="entry-meta">${escapeHtml(selectedHex)}</span>
            </div>
            <div class="row">
              <label>Force Type
                <select name="type">
                  ${HEX_MAP_FORCE_TYPES.map((type) => `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`).join("")}
                </select>
              </label>
              <label>Force Name
                <input name="name" placeholder="South Road Caravan" />
              </label>
            </div>
            <label>Force Notes
              <textarea name="notes" placeholder="Strength, destination, commander, threat level, current goal."></textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Add Force</button>
            </div>
          </form>

          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Markers In ${escapeHtml(selectedHex)}</span>
            </div>
            <div class="card-list">
              ${markers.length ? markers.map((marker) => renderHexMapMarkerEntry(marker)).join("") : `<p class="empty">No markers in this hex yet.</p>`}
            </div>
          </article>

          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Forces In ${escapeHtml(selectedHex)}</span>
            </div>
            <div class="card-list">
              ${forces.length ? forces.map((force) => renderHexMapForceEntry(force)).join("") : `<p class="empty">No force markers in this hex yet.</p>`}
            </div>
          </article>

          <article class="entry">
            <div class="entry-head">
              <span class="entry-title">Linked Locations</span>
            </div>
            ${
              linkedLocations.length
                ? `<ul class="flow-list">${linkedLocations
                    .map((location) => `<li><strong>${escapeHtml(location.name || "Unnamed Location")}</strong>${location.whatChanged ? `: ${escapeHtml(location.whatChanged)}` : ""}</li>`)
                    .join("")}</ul>`
                : `<p class="small">No location records currently use ${escapeHtml(selectedHex)}.</p>`
            }
          </article>
        </aside>
      </section>
      <datalist id="hexmap-terrain-options">
        ${HEX_MAP_TERRAIN_OPTIONS.map((terrain) => `<option value="${escapeHtml(terrain)}"></option>`).join("")}
      </datalist>
    </div>
  `;
}

function renderHexMapSvgCells(hexMap, selectedHex) {
  const metrics = getHexMapMetrics(hexMap);
  const selected = normalizeHexCoordinate(selectedHex, hexMap.columns, hexMap.rows);
  const party = getHexMapParty(hexMap);
  const cells = [];
  for (let col = 0; col < metrics.columns; col += 1) {
    for (let row = 0; row < metrics.rows; row += 1) {
      const hex = `${getHexColumnLabel(col)}${row + 1}`;
      const center = getHexCenter(col, row, hexMap);
      const region = getKingdomRegionByHex(hex);
      const markers = getHexMapMarkersForHex(hex);
      const forces = getHexMapForcesForHex(hex);
      const isPartyHex = party.hex === hex;
      const markerOffsetBase = center.cx - (Math.max(0, markers.length - 1) * 9) / 2;
      const forceOffsetBase = center.cx - (Math.max(0, forces.length - 1) * 14) / 2;
      cells.push(`
        <g class="hexmap-cell-group ${hex === selected ? "is-selected" : ""}">
          <polygon
            class="hexmap-cell ${hex === selected ? "is-selected" : ""}"
            data-action="hexmap-select-hex"
            data-hex="${escapeHtml(hex)}"
            points="${escapeHtml(buildHexPolygonPoints(center.cx, center.cy, metrics.size))}"
            fill="${escapeHtml(getHexStatusColor(region?.status || ""))}"
            style="--hex-fill-opacity:${escapeHtml(String(hexMap.gridFillOpacity))}; --hex-stroke:rgba(77, 58, 29, ${escapeHtml(String(hexMap.gridLineOpacity))});"
          >
            <title>${escapeHtml(`${hex} • ${region?.status || "Unclaimed"}${region?.terrain ? ` • ${region.terrain}` : ""}`)}</title>
          </polygon>
          ${
            hexMap.showLabels
              ? `<text x="${escapeHtml(String(center.cx.toFixed(2)))}" y="${escapeHtml(String((center.cy - 4).toFixed(2)))}" class="hexmap-label">${escapeHtml(hex)}</text>`
              : ""
          }
          ${
            region?.workSite
              ? `<text x="${escapeHtml(String(center.cx.toFixed(2)))}" y="${escapeHtml(String((center.cy + 15).toFixed(2)))}" class="hexmap-sub-label">${escapeHtml(region.workSite)}</text>`
              : ""
          }
          ${
            isPartyHex
              ? `
                <g data-action="hexmap-select-hex" data-hex="${escapeHtml(hex)}" class="hexmap-party-group">
                  <circle
                    cx="${escapeHtml(String(center.cx.toFixed(2)))}"
                    cy="${escapeHtml(String((center.cy - metrics.size * 0.5).toFixed(2)))}"
                    r="13"
                    class="hexmap-party-marker"
                  >
                    <title>${escapeHtml(`${party.label || "Party"}: ${hex}`)}</title>
                  </circle>
                  <text
                    x="${escapeHtml(String(center.cx.toFixed(2)))}"
                    y="${escapeHtml(String((center.cy - metrics.size * 0.5 + 4).toFixed(2)))}"
                    class="hexmap-party-label"
                  >P</text>
                </g>
              `
              : ""
          }
          ${forces
            .slice(0, 3)
            .map(
              (force, index) => `
                <g data-action="hexmap-select-hex" data-hex="${escapeHtml(hex)}" class="hexmap-force-group">
                  <rect
                    x="${escapeHtml(String((forceOffsetBase + index * 28 - 11).toFixed(2)))}"
                    y="${escapeHtml(String((center.cy - metrics.size * 0.1 - 11).toFixed(2)))}"
                    width="22"
                    height="22"
                    rx="7"
                    fill="${escapeHtml(getHexForceColor(force.type))}"
                    class="hexmap-force-marker"
                  >
                    <title>${escapeHtml(`${force.type}: ${force.name || "Unnamed force"}`)}</title>
                  </rect>
                  <text
                    x="${escapeHtml(String((forceOffsetBase + index * 28).toFixed(2)))}"
                    y="${escapeHtml(String((center.cy - metrics.size * 0.1 + 4).toFixed(2)))}"
                    class="hexmap-force-label"
                  >${escapeHtml(getHexForceGlyph(force.type))}</text>
                </g>
              `
            )
            .join("")}
          ${markers
            .slice(0, 4)
            .map(
              (marker, index) => `
                <g data-action="hexmap-select-hex" data-hex="${escapeHtml(hex)}" class="hexmap-marker-group">
                  <circle
                    cx="${escapeHtml(String((markerOffsetBase + index * 18).toFixed(2)))}"
                    cy="${escapeHtml(String((center.cy + metrics.size * 0.55).toFixed(2)))}"
                    r="7"
                    fill="${escapeHtml(getHexMarkerColor(marker.type))}"
                    class="hexmap-marker-svg"
                  >
                    <title>${escapeHtml(`${marker.type}: ${marker.title || "Untitled marker"}`)}</title>
                  </circle>
                </g>
              `
            )
            .join("")}
        </g>
      `);
    }
  }
  return cells.join("");
}

function renderHexMapPartyTrail(hexMap) {
  const party = getHexMapParty(hexMap);
  const trail = Array.isArray(party.trail) ? party.trail.slice(0, 12) : [];
  if (!trail.length) return "";
  const ordered = [...trail].reverse();
  const points = ordered
    .map((entry) => {
      const parsed = parseHexCoordinate(entry.hex);
      if (!parsed) return null;
      const center = getHexCenter(parsed.columnIndex, parsed.rowIndex, hexMap);
      return center;
    })
    .filter(Boolean);
  if (!points.length) return "";
  const polyline = points.length > 1
    ? `<polyline class="hexmap-party-trail-line" points="${escapeHtml(points.map((point) => `${point.cx.toFixed(2)},${point.cy.toFixed(2)}`).join(" "))}" />`
    : "";
  const dots = points
    .map(
      (point, index) => `
        <circle
          class="hexmap-party-trail-dot ${index === points.length - 1 ? "is-current" : ""}"
          cx="${escapeHtml(String(point.cx.toFixed(2)))}"
          cy="${escapeHtml(String(point.cy.toFixed(2)))}"
          r="${index === points.length - 1 ? "6" : "4"}"
        />
      `
    )
    .join("");
  return `<g class="hexmap-party-trail-layer">${polyline}${dots}</g>`;
}

function renderHexMapMarkerEntry(marker) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(marker.title || "Untitled marker")}</span>
        <span class="entry-meta">${escapeHtml(marker.type || "Note")} • ${escapeHtml(marker.hex || "")}</span>
      </div>
      <div class="row">
        <label>Type
          <select data-collection="hexMapMarkers" data-id="${marker.id}" data-field="type">
            ${HEX_MAP_MARKER_TYPES.map((type) => `<option value="${escapeHtml(type)}" ${marker.type === type ? "selected" : ""}>${escapeHtml(type)}</option>`).join("")}
          </select>
        </label>
        <label>Title
          <input data-collection="hexMapMarkers" data-id="${marker.id}" data-field="title" value="${escapeHtml(marker.title || "")}" />
        </label>
      </div>
      <label>Notes
        <textarea data-collection="hexMapMarkers" data-id="${marker.id}" data-field="notes">${escapeHtml(marker.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="hexMapMarkers" data-id="${marker.id}">Delete Marker</button>
      </div>
    </article>
  `;
}

function renderHexMapForceEntry(force) {
  return `
    <article class="entry">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(force.name || "Unnamed force")}</span>
        <span class="entry-meta">${escapeHtml(force.type || "Allied Force")} • ${escapeHtml(force.hex || "")}</span>
      </div>
      <div class="row">
        <label>Type
          <select data-collection="hexMapForces" data-id="${force.id}" data-field="type">
            ${HEX_MAP_FORCE_TYPES.map((type) => `<option value="${escapeHtml(type)}" ${force.type === type ? "selected" : ""}>${escapeHtml(type)}</option>`).join("")}
          </select>
        </label>
        <label>Name
          <input data-collection="hexMapForces" data-id="${force.id}" data-field="name" value="${escapeHtml(force.name || "")}" />
        </label>
      </div>
      <label>Notes
        <textarea data-collection="hexMapForces" data-id="${force.id}" data-field="notes">${escapeHtml(force.notes || "")}</textarea>
      </label>
      <div class="toolbar">
        <button class="btn btn-danger" data-action="delete" data-collection="hexMapForces" data-id="${force.id}">Delete Force</button>
      </div>
    </article>
  `;
}

function normalizeRulesSearchQuery(value) {
  return str(value).replace(/\s+/g, " ").trim();
}

function getRulesSearchTerms(value) {
  return extractRetrievalTerms(normalizeRulesSearchQuery(value)).slice(0, 12);
}

function scoreRulesTitleMatch(query, result) {
  const cleanQuery = normalizeRulesSearchQuery(query).toLowerCase();
  if (!cleanQuery) return 0;
  const title = str(result?.title).toLowerCase();
  const pathValue = str(result?.path).toLowerCase();
  const titleWords = cleanQuery.split(/\s+/).filter(Boolean);
  let score = 0;
  if (title === cleanQuery) score += 12;
  if (title.startsWith(cleanQuery)) score += 8;
  if (title.includes(cleanQuery)) score += 5;
  if (pathValue.includes(cleanQuery)) score += 3;
  if (titleWords.length && titleWords.every((word) => title.includes(word))) score += 4;
  return score;
}

function sortRulesResults(query, results) {
  return [...(Array.isArray(results) ? results : [])].sort((a, b) => {
    const scoreA = Number(a?.score || 0) + scoreRulesTitleMatch(query, a);
    const scoreB = Number(b?.score || 0) + scoreRulesTitleMatch(query, b);
    return scoreB - scoreA || str(a?.title).localeCompare(str(b?.title));
  });
}

function getRecentRuleCaptureEntries(limit = 8) {
  return [...(state.liveCapture || [])]
    .filter((entry) => {
      const kind = str(entry?.kind).toLowerCase();
      return kind === "rule" || kind === "retcon" || kind === "house rule";
    })
    .sort((a, b) => safeDate(b.timestamp) - safeDate(a.timestamp))
    .slice(0, limit);
}

function extractRelevantRuleExcerpt(text, query, limit = 440) {
  const source = str(text);
  if (!source) return "";
  const terms = getRulesSearchTerms(query);
  if (!terms.length) return compactLine(source, limit);
  const paragraphs = source
    .replace(/\r/g, "")
    .split(/\n{2,}/)
    .map((item) => item.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  if (!paragraphs.length) return compactLine(source, limit);
  const ranked = paragraphs
    .map((paragraph, index) => ({
      paragraph,
      index,
      score: countRetrievalTokenHits(paragraph.toLowerCase(), terms),
    }))
    .sort((a, b) => b.score - a.score || a.index - b.index);
  const chosen = ranked[0]?.score > 0 ? ranked[0].paragraph : paragraphs[0];
  return compactLine(chosen, limit);
}

function buildRulesLocalMatches(query) {
  const aiMemory = buildAiMemoryDigests(state);
  state.meta.aiMemory = aiMemory;
  const effectiveDigest = str(aiMemory.rulingsDigest);
  const manualDigest = str(aiMemory.manualRulings);
  const ruleEntries = getRecentRuleCaptureEntries(10);
  const rulesStore = ensureRulesStore();
  const terms = getRulesSearchTerms(query);

  const digestMatches = [];
  if (effectiveDigest) {
    const digestText = extractRelevantRuleExcerpt(effectiveDigest, query, 520);
    const digestScore = terms.length ? countRetrievalTokenHits(digestText.toLowerCase(), terms) : 1;
    if (digestScore > 0 || !terms.length) {
      digestMatches.push({
        kind: manualDigest ? "Manual digest" : "Effective digest",
        title: manualDigest ? "Manual House Rules Digest" : "Derived Rulings Digest",
        text: digestText,
        score: digestScore,
      });
    }
  }

  const entryMatches = ruleEntries
    .map((entry) => {
      const note = str(entry.note);
      const score = terms.length ? countRetrievalTokenHits(note.toLowerCase(), terms) : 1;
      return {
        id: entry.id,
        kind: str(entry.kind) || "Rule",
        title: `${entry.kind || "Rule"} • ${formatAiHistoryTimestamp(entry.timestamp) || "No timestamp"}`,
        text: note,
        score,
      };
    })
    .filter((entry) => (terms.length ? entry.score > 0 : true))
    .sort((a, b) => b.score - a.score || a.title.localeCompare(b.title))
    .slice(0, 5);

  const savedMatches = rulesStore
    .map((entry) => {
      const haystack = `${entry.title} ${entry.text} ${entry.tags.join(" ")}`.toLowerCase();
      const score = terms.length ? countRetrievalTokenHits(haystack, terms) : 1;
      return {
        id: entry.id,
        title: entry.title,
        kind: entry.kind,
        kindLabel: getRuleStoreKindLabel(entry.kind),
        text: extractRelevantRuleExcerpt(entry.text, query, 520),
        sourceTitle: entry.sourceTitle,
        sourceUrl: entry.sourceUrl,
        tags: Array.isArray(entry.tags) ? entry.tags : [],
        updatedAt: entry.updatedAt || entry.createdAt || "",
        score,
      };
    })
    .filter((entry) => (terms.length ? entry.score > 0 : true))
    .sort((a, b) => b.score - a.score || safeDate(b.updatedAt) - safeDate(a.updatedAt))
    .slice(0, 8);

  return {
    aiMemory,
    digestMatches,
    entryMatches,
    savedMatches,
  };
}

function getSelectedRulesResult(results = ui.rulesResults) {
  const selectedUrl = str(ui.rulesSelectedUrl);
  const list = Array.isArray(results) ? results : [];
  if (selectedUrl) {
    const matched = list.find((result) => str(result?.url) === selectedUrl);
    if (matched) return matched;
  }
  return list[0] || null;
}

function saveSelectedOfficialRuleToStore() {
  const selected = getSelectedRulesResult(ui.rulesResults);
  if (!selected) {
    ui.rulesMessage = "Select an official rule first.";
    render();
    return;
  }
  const entry = upsertRulesStoreEntry({
    title: str(selected.title || ui.rulesSearchQuery || "Official PF2e Rule"),
    kind: "official_note",
    text: str(selected.snippet || ""),
    sourceTitle: str(selected.title || ""),
    sourceUrl: str(selected.url || ""),
    sourceOrigin: "official",
    tags: buildRuleStoreTags({ title: str(selected.title || ""), text: str(selected.snippet || ""), kind: "official_note" }),
  });
  state.meta.aiMemory = buildAiMemoryDigests(state);
  saveState();
  ui.rulesMessage = `Saved official note: ${entry.title}.`;
  render();
}

function saveCopilotOutputToRulesStore(kind = "accepted_ruling") {
  const output = str(ui.copilotDraft.output);
  if (!output) {
    ui.copilotMessage = "No Loremaster output to save.";
    render();
    return;
  }
  const request = buildGlobalCopilotRequest(activeTab, ui.copilotDraft.input, false);
  const selectedRule = activeTab === "rules" ? getSelectedRulesResult(ui.rulesResults) : null;
  const titleBase = selectedRule?.title || ui.copilotDraft.input || request?.taskLabel || "Loremaster Note";
  const entry = upsertRulesStoreEntry({
    title: compactLine(titleBase, 140),
    kind,
    text: output,
    sourceTitle: str(selectedRule?.title || request?.taskLabel || ""),
    sourceUrl: str(selectedRule?.url || ""),
    sourceOrigin: "loremaster",
    tags: buildRuleStoreTags({ title: titleBase, text: output, kind }),
  });
  state.meta.aiMemory = buildAiMemoryDigests(state);
  saveState();
  ui.copilotMessage = `${getRuleStoreKindLabel(kind)} saved to the local rules/canon store.`;
  render();
}

function buildRulesPromptFromResult(result, mode = "explain") {
  const title = str(result?.title || "PF2e rule");
  const url = str(result?.url || "");
  const snippet = str(result?.snippet || "");
  if (mode === "compare") {
    return [
      `Using the official Pathfinder 2e rule page "${title}", compare the official rule against my local rulings digest.`,
      "Return:",
      "Official Rule:",
      "- 2 to 5 bullets",
      "Local Override / House Rule:",
      "- bullet",
      "GM Quick Ruling:",
      "- 2 to 4 bullets",
      "Source Trail:",
      `- ${title}${url ? ` (${url})` : ""}`,
      "",
      "Official excerpt:",
      snippet,
    ].join("\n");
  }
  return [
    `Using the official Pathfinder 2e rule page "${title}", explain how this works at the table.`,
    "Return:",
    "Rules Answer:",
    "- 3 to 6 concise bullets",
    "Official vs Local Notes:",
    "- Confirmed official rule",
    "- Local override if one exists",
    "Source Trail:",
    `- ${title}${url ? ` (${url})` : ""}`,
    "",
    "Official excerpt:",
    snippet,
  ].join("\n");
}

function renderRulesResultCard(result, selected) {
  if (!result) return "";
  const title = str(result.title || "PF2e rule");
  const url = str(result.url || "");
  const pathValue = str(result.path || "");
  return `
    <article class="entry rules-result-card ${selected ? "selected" : ""}">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(title)}</span>
        <span class="entry-meta">${escapeHtml(result.source || "Archives of Nethys")} • Score ${escapeHtml(String(result.score || 0))}</span>
      </div>
      <p>${escapeHtml(str(result.snippet || "No excerpt available."))}</p>
      <p class="small mono">${escapeHtml(pathValue || url)}</p>
      <div class="toolbar">
        <button class="btn ${selected ? "btn-primary" : "btn-secondary"}" data-action="rules-select-result" data-url="${encodeURIComponent(url)}">Select</button>
        <button class="btn btn-secondary" data-action="rules-open-result" data-url="${encodeURIComponent(url)}" ${url ? "" : "disabled"}>Open AoN</button>
        <button class="btn btn-secondary" data-action="rules-use-result" data-url="${encodeURIComponent(url)}">Ask Loremaster</button>
        <button class="btn btn-secondary" data-action="rules-compare-result" data-url="${encodeURIComponent(url)}">Compare Official vs Local</button>
        <button class="btn btn-secondary" data-action="rules-save-result" data-url="${encodeURIComponent(url)}">Save Official Note</button>
      </div>
    </article>
  `;
}

function renderRulesStoreEntryCard(entry) {
  if (!entry) return "";
  const sourceUrl = str(entry.sourceUrl || "");
  return `
    <article class="entry rules-result-card">
      <div class="entry-head">
        <span class="entry-title">${escapeHtml(entry.title || "Saved entry")}</span>
        <span class="entry-meta">${escapeHtml(entry.kindLabel || getRuleStoreKindLabel(entry.kind))}${entry.updatedAt ? ` • ${escapeHtml(formatAiHistoryTimestamp(entry.updatedAt) || "")}` : ""}</span>
      </div>
      <p>${escapeHtml(str(entry.text || ""))}</p>
      ${entry.tags?.length ? `<p class="small">Tags: ${escapeHtml(entry.tags.join(", "))}</p>` : ""}
      ${sourceUrl ? `<p class="small mono">${escapeHtml(sourceUrl)}</p>` : ""}
      <div class="toolbar">
        <button class="btn btn-secondary" data-action="rules-delete-store-entry" data-id="${encodeURIComponent(entry.id)}">Delete</button>
        <button class="btn btn-secondary" data-action="rules-open-store-source" data-url="${encodeURIComponent(sourceUrl)}" ${sourceUrl ? "" : "disabled"}>Open Source</button>
      </div>
    </article>
  `;
}

function renderRules() {
  const aiMemory = buildAiMemoryDigests(state);
  state.meta.aiMemory = aiMemory;
  const query = normalizeRulesSearchQuery(ui.rulesSearchQuery);
  const officialResults = sortRulesResults(query, ui.rulesResults);
  const selectedResult = getSelectedRulesResult(officialResults);
  const localMatches = buildRulesLocalMatches(query);
  const savedRulesCount = ensureRulesStore().length;
  const scope = str(ui.rulesScope || "both") || "both";
  const showOfficial = scope !== "local";
  const showLocal = scope !== "official";
  const desktopSearchReady = !!desktopApi?.searchAonRules;
  const searchLabel = ui.rulesBusy ? "Searching..." : "Search Official Rules";

  return `
    <div class="page-stack">
      ${renderPageIntro("PF2e Rules", "Look up official Pathfinder 2e rules from Archives of Nethys, compare them against your local rulings digest, and send a grounded rules question into Loremaster.")}

      <section class="panel flow-panel">
        <h2>Run Order</h2>
        <ol class="flow-list">
          <li><strong>Step 1:</strong> search the official rules index for the exact PF2e term you need.</li>
          <li><strong>Step 2:</strong> review the local rulings / house-rules side of the split view.</li>
          <li><strong>Step 3:</strong> use Loremaster to produce a table-ready ruling grounded in both.</li>
        </ol>
        <p class="small">AoN live lookup: ${desktopSearchReady ? "available" : "desktop bridge not available"}${ui.rulesIndexedAt ? ` • Last cached index ${escapeHtml(ui.rulesIndexedAt)}` : ""}</p>
        ${ui.rulesMessage ? `<p class="small">${escapeHtml(ui.rulesMessage)}</p>` : ""}
      </section>

      <section class="grid grid-2">
        <article class="panel">
          <h2>Official Rules Search</h2>
          <form data-form="rules-search">
            <div class="row">
              <label>Search Query
                <input name="query" value="${escapeHtml(query)}" placeholder="bleed, concealed, persistent damage, aid, command an animal" />
              </label>
              <label>Result Limit
                <select name="limit">
                  ${[3, 5, 6].map((value) => `<option value="${value}" ${Number(ui.rulesSearchLimit) === value ? "selected" : ""}>${value}</option>`).join("")}
                </select>
              </label>
              <label>View Scope
                <select name="scope">
                  <option value="both" ${scope === "both" ? "selected" : ""}>Official + Local</option>
                  <option value="official" ${scope === "official" ? "selected" : ""}>Official Only</option>
                  <option value="local" ${scope === "local" ? "selected" : ""}>Local Only</option>
                </select>
              </label>
            </div>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit" ${desktopSearchReady ? "" : "disabled"}>${searchLabel}</button>
              <button class="btn btn-secondary" type="button" data-action="rules-refresh-search" ${desktopSearchReady && query ? "" : "disabled"}>Force Refresh</button>
              <button class="btn btn-secondary" type="button" data-action="rules-use-query" ${query ? "" : "disabled"}>Send Query To Loremaster</button>
            </div>
          </form>
          <p class="small">Searches use exact-title bias first, then broader title/path scoring. This keeps PF2e terms like conditions, actions, and subsystems from drifting into fuzzy fantasy matches.</p>
        </article>

        <article class="panel">
          <h2>House Rules / Rulings Digest</h2>
          <form data-form="ai-memory-rulings">
            <label>Manual Rulings Digest
              <textarea name="manualRulings" placeholder="Example: Hero Point rerolls must be declared before new information is revealed. We use remaster terminology only.">${escapeHtml(
                aiMemory.manualRulings || ""
              )}</textarea>
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Save Rulings Digest</button>
            </div>
          </form>
          <article class="memory-card" style="margin-top:12px;">
            <h3>Effective Local Rules Context</h3>
            <div class="memory-block">${renderMultilineText(aiMemory.rulingsDigest || "No local rulings digest yet.")}</div>
          </article>
        </article>
      </section>

      <section class="grid grid-2 rules-split">
        <article class="panel">
          <h2>Official Rules Matches</h2>
          ${
            showOfficial
              ? officialResults.length
                ? `<div class="rules-result-list">${officialResults
                    .map((result) => renderRulesResultCard(result, str(result.url) === str(selectedResult?.url)))
                    .join("")}</div>`
                : `<p class="empty">No official rule matches yet. Search for a condition, action, subsystem, or rule term above.</p>`
              : `<p class="small">Official matches hidden by the current scope filter.</p>`
          }
        </article>

        <article class="panel">
          <h2>Official vs Local Split</h2>
          ${
            showOfficial && selectedResult
              ? `
                <article class="memory-card">
                  <h3>Selected Official Rule</h3>
                  <p><strong>${escapeHtml(str(selectedResult.title))}</strong></p>
                  <p class="small mono">${escapeHtml(str(selectedResult.url || selectedResult.path || ""))}</p>
                  <div class="memory-block">${renderMultilineText(str(selectedResult.snippet || "No excerpt available."))}</div>
                </article>
              `
              : `<article class="memory-card"><h3>Selected Official Rule</h3><div class="memory-block">${renderMultilineText("No official rule selected.")}</div></article>`
          }

          ${
            showLocal
              ? `
                <article class="memory-card" style="margin-top:12px;">
                  <h3>Local Matches</h3>
                  ${
                    localMatches.savedMatches.length || localMatches.digestMatches.length || localMatches.entryMatches.length
                      ? `
                        ${
                          localMatches.savedMatches.length
                            ? `
                              <div class="rules-local-match">
                                <p><strong>Saved Rules / Canon Store</strong></p>
                                ${localMatches.savedMatches.map((entry) => renderRulesStoreEntryCard(entry)).join("")}
                              </div>
                            `
                            : ""
                        }
                        ${localMatches.digestMatches
                          .map(
                            (match) => `
                              <div class="rules-local-match">
                                <p><strong>${escapeHtml(match.title)}</strong></p>
                                <div class="memory-block">${renderMultilineText(match.text)}</div>
                              </div>
                            `
                          )
                          .join("")}
                        ${localMatches.entryMatches
                          .map(
                            (entry) => `
                              <div class="rules-local-match">
                                <p><strong>${escapeHtml(entry.title)}</strong></p>
                                <div class="memory-block">${renderMultilineText(entry.text)}</div>
                              </div>
                            `
                          )
                          .join("")}
                      `
                      : `<div class="memory-block">${renderMultilineText("No matching local rulings were found for the current query.")}</div>`
                  }
                </article>
              `
              : `<article class="memory-card" style="margin-top:12px;"><h3>Local Matches</h3><div class="memory-block">${renderMultilineText("Local rulings hidden by the current scope filter.")}</div></article>`
          }
        </article>
      </section>

      <section class="grid grid-2">
        <article class="panel">
          <h2>Manual Local Rule / Canon Entry</h2>
          <form data-form="rules-store-entry">
            <div class="row">
              <label>Title
                <input name="title" placeholder="Aid clarification, hero point policy, canon fact" value="${escapeHtml(query ? compactLine(query, 120) : "")}" />
              </label>
              <label>Type
                <select name="kind">
                  <option value="accepted_ruling">Accepted Ruling</option>
                  <option value="house_rule">House Rule</option>
                  <option value="canon_memory">Canon Memory</option>
                  <option value="official_note">Official Note</option>
                </select>
              </label>
            </div>
            <label>Text
              <textarea name="text" placeholder="Save a short canonical ruling, house rule, or campaign truth here."></textarea>
            </label>
            <label>Source URL (optional)
              <input name="sourceUrl" placeholder="https://2e.aonprd.com/Rules.aspx?ID=..." />
            </label>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Save Local Entry</button>
            </div>
          </form>
        </article>

        <article class="panel">
          <h2>Persistent Rules & Canon Store</h2>
          <p class="small">${escapeHtml(String(savedRulesCount))} saved entry${savedRulesCount === 1 ? "" : "ies"} that Loremaster can retrieve later.</p>
          ${
            savedRulesCount
              ? `<div class="rules-result-list">${ensureRulesStore()
                  .slice(0, 10)
                  .map((entry) =>
                    renderRulesStoreEntryCard({
                      ...entry,
                      kindLabel: getRuleStoreKindLabel(entry.kind),
                    })
                  )
                  .join("")}</div>`
              : `<p class="empty">No persistent rules or canon entries saved yet.</p>`
          }
        </article>
      </section>
    </div>
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
    useAonRules: true,
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
    useAonRules: current.useAonRules === false ? false : true,
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
    next.useAonRules = true;
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
    next.useAonRules = true;
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
  if (tabId === "rules") return "Ask a PF2e rules question like: how does bleed work, what does concealed do, or compare official vs local rulings.";
  if (tabId === "kingdom") return "Ask for kingdom-turn help, action order, leader assignments, or settlement advice.";
  if (tabId === "hexmap") return "Ask for hex encounters, work-site ideas, event seeds, or consequences for the selected hex.";
  if (tabId === "npcs") return "Describe an NPC concept and ask for table-ready details.";
  if (tabId === "quests") return "Describe a quest idea and ask for objective/stakes.";
  if (tabId === "locations") return "Describe a hex/location and ask for a usable scene brief.";
  if (tabId === "pdf") return "Ask for best PDF search queries for your next session.";
  if (tabId === "obsidian") return "Ask for vault folder structures, note templates, or cleaner markdown organization.";
  if (tabId === "foundry") return "Ask for Foundry handoff checklist for this session.";
  if (tabId === "writing") return "Ask for rewrite help, stronger wording, and clean structure.";
  return "Ask a GM question or chat naturally (example: hello, help me prep tonight).";
}

function getGlobalCopilotApplyLabel(tabId) {
  if (tabId === "sessions") return "Apply to Latest Session";
  if (tabId === "capture") return "Add as Live Capture";
  if (tabId === "rules") return "Attach to Latest Prep";
  if (tabId === "kingdom") return "Append Kingdom Notes";
  if (tabId === "hexmap") return "Append to Selected Hex";
  if (tabId === "npcs") return "Create NPC(s)";
  if (tabId === "quests") return "Create Quest";
  if (tabId === "locations") return "Create Location";
  if (tabId === "pdf") return "Use as PDF Query";
  if (tabId === "obsidian") return "No Direct Apply";
  if (tabId === "foundry") return "Attach Foundry Notes";
  if (tabId === "writing") return "Send to Writing Helper";
  return "Attach to Latest Prep";
}

function getGlobalCopilotMode(tabId) {
  if (tabId === "rules") return "assistant";
  if (tabId === "npcs") return "npc";
  if (tabId === "quests") return "quest";
  if (tabId === "locations") return "location";
  if (tabId === "hexmap") return "location";
  if (tabId === "sessions" || tabId === "capture" || tabId === "writing" || tabId === "kingdom") return "session";
  return "prep";
}

const COPILOT_TASK_META = Object.freeze({
  general_prep: { label: "General Prep", saveTarget: "latest session prep", mode: "prep" },
  rules_question: { label: "Rules Question", saveTarget: "answer only", mode: "assistant" },
  campaign_lookup: { label: "Campaign Lookup", saveTarget: "answer only", mode: "assistant" },
  session_summary: { label: "Session Summary", saveTarget: "latest session", mode: "session" },
  note_update: { label: "Note Update", saveTarget: "review before apply", mode: "session" },
  kingdom_helper: { label: "Kingdom Helper", saveTarget: "kingdom notes", mode: "session" },
  map_helper: { label: "Hex Map Helper", saveTarget: "selected hex or map notes", mode: "location" },
  pdf_lookup: { label: "PDF-Grounded Lookup", saveTarget: "pdf query or prep", mode: "prep" },
  vault_workflow: { label: "Vault Workflow", saveTarget: "vault note or folder plan", mode: "prep" },
  foundry_handoff: { label: "Foundry Handoff", saveTarget: "latest session prep", mode: "prep" },
  writing_cleanup: { label: "Writing Cleanup", saveTarget: "writing helper output", mode: "session" },
  small_talk: { label: "Small Talk", saveTarget: "none", mode: "assistant" },
});

function renderObsidian() {
  const settings = ensureObsidianSettings();
  const looksLikeVault = settings.looksLikeVault === true;
  const syncFolderPreview = settings.vaultPath
    ? `${settings.vaultPath}${settings.vaultPath.endsWith("\\") ? "" : "\\"}${settings.baseFolder || "DM Helper"}`
    : "(choose a vault folder first)";
  const syncLabel = ui.obsidianBusy ? "Syncing..." : "Sync To Vault";
  const hasAiOutput = str(ui.copilotDraft.output).length > 0;
  const writeLabel = ui.obsidianBusy ? "Writing..." : "Write Current AI Output To Vault";
  const readScope = settings.readWholeVault ? "Whole vault" : "DM Helper folder only";

  return `
    <div class="page-stack">
      ${renderPageIntro("Obsidian Vault", "Connect Loremaster to your local Obsidian vault. DM Helper can export campaign notes, pull compact vault context into AI prompts, and write current AI output back into a markdown note.")}

      <section class="panel flow-panel">
        <h2>How This Works</h2>
        <ol class="flow-list">
          <li><strong>Step 1:</strong> choose your Obsidian vault folder.</li>
          <li><strong>Step 2:</strong> choose the DM Helper folder name that will be created inside the vault.</li>
          <li><strong>Step 3:</strong> decide whether Loremaster should read the whole vault or only the DM Helper folder for AI context.</li>
          <li><strong>Step 4:</strong> sync sessions, NPCs, quests, locations, kingdom notes, and hex map notes into markdown files, or write the current AI output back into a note.</li>
        </ol>
        <p class="small">Current vault: <span class="mono">${escapeHtml(settings.vaultPath || "Not set")}</span></p>
        <p class="small">Sync folder: <span class="mono">${escapeHtml(syncFolderPreview)}</span></p>
        <p class="small">Vault check: ${escapeHtml(
          settings.vaultPath ? (looksLikeVault ? "Looks like a real Obsidian vault." : "Folder selected, but no .obsidian folder was detected yet.") : "No vault selected."
        )}</p>
        <p class="small">Last sync: ${escapeHtml(settings.lastSyncAt || "Never")} ${settings.lastSyncSummary ? `• ${escapeHtml(settings.lastSyncSummary)}` : ""}</p>
        <p class="small">AI read scope: ${escapeHtml(readScope)} • Note limit ${escapeHtml(String(settings.aiContextNoteLimit))} • Character budget ${escapeHtml(String(settings.aiContextCharLimit))}</p>
        <p class="small">Last AI note: ${escapeHtml(settings.lastAiNotePath || "Never written")} ${settings.lastAiNoteAt ? `• ${escapeHtml(settings.lastAiNoteAt)}` : ""}</p>
        ${ui.obsidianMessage ? `<p class="small">${escapeHtml(ui.obsidianMessage)}</p>` : ""}
      </section>

      <section class="grid grid-2">
        <article class="panel">
          <h2>Vault Settings</h2>
          <form data-form="obsidian-settings">
            <label>Vault Folder
              <input name="vaultPath" value="${escapeHtml(settings.vaultPath || "")}" placeholder="C:\\Users\\Chris Bender\\Documents\\Obsidian Vault" />
            </label>
            <label>DM Helper Folder Inside Vault
              <input name="baseFolder" value="${escapeHtml(settings.baseFolder || "DM Helper")}" placeholder="DM Helper" />
            </label>
            <label>AI Note Folder Inside DM Helper Folder
              <input name="aiWriteFolder" value="${escapeHtml(settings.aiWriteFolder || "AI Notes")}" placeholder="AI Notes" />
            </label>
            <label>
              <input type="checkbox" name="useForAiContext" ${settings.useForAiContext ? "checked" : ""} />
              Let Loremaster read vault notes for AI context
            </label>
            <label>
              <input type="checkbox" name="readWholeVault" ${settings.readWholeVault ? "checked" : ""} />
              Read the whole vault, not just the DM Helper folder
            </label>
            <div class="row">
              <label>AI Context Note Limit
                <input name="aiContextNoteLimit" type="number" min="1" max="12" step="1" value="${escapeHtml(String(settings.aiContextNoteLimit || 6))}" />
              </label>
              <label>AI Context Character Budget
                <input name="aiContextCharLimit" type="number" min="800" max="12000" step="100" value="${escapeHtml(String(settings.aiContextCharLimit || 3600))}" />
              </label>
            </div>
            <div class="toolbar">
              <button class="btn btn-primary" type="submit">Save Vault Settings</button>
              <button class="btn btn-secondary" type="button" data-action="obsidian-choose-vault">Choose Vault Folder</button>
              <button class="btn btn-secondary" type="button" data-action="obsidian-open-vault" ${settings.vaultPath ? "" : "disabled"}>Open Vault Folder</button>
            </div>
          </form>
        </article>

        <article class="panel">
          <h2>AI + Sync Actions</h2>
          <ul class="flow-list">
            <li><strong>Campaign Home:</strong> summary note with quick links.</li>
            <li><strong>Sessions:</strong> one markdown note per session.</li>
            <li><strong>NPCs / Quests / Locations:</strong> one markdown note per record.</li>
            <li><strong>Kingdom:</strong> kingdom sheet snapshot, regions, settlements, and recent turns.</li>
            <li><strong>Hex Map:</strong> party position, trail, forces, markers, and region notes.</li>
            <li><strong>AI Read Context:</strong> compact excerpts from the most relevant vault notes are added to Loremaster prompts when enabled.</li>
            <li><strong>AI Write Back:</strong> writes the current Loremaster output into <span class="mono">${escapeHtml(settings.aiWriteFolder || "AI Notes")}</span>.</li>
          </ul>
          <p class="small">This sync overwrites matching DM Helper notes, but it does not delete stale files you removed from the app.</p>
          <div class="toolbar">
            <button class="btn btn-primary" type="button" data-action="obsidian-sync" ${ui.obsidianBusy ? "disabled" : ""}>${syncLabel}</button>
            <button class="btn btn-secondary" type="button" data-action="obsidian-write-current-ai" ${(hasAiOutput && !ui.obsidianBusy) ? "" : "disabled"}>${writeLabel}</button>
          </div>
        </article>
      </section>
    </div>
  `;
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

function isRulesQuestionInput(inputText) {
  const lower = str(inputText).toLowerCase();
  if (!lower) return false;
  return (
    /\b(how does|how do|what are the requirements|what is the dc|what does this condition do|how does persistent damage|how does bleed|how does invisibility|how do reactions work)\b/.test(
      lower
    ) ||
    /\b(bleed|persistent damage|condition|trait|feat|spell|focus spell|reaction|free action|activity|saving throw|armor class|ac|fortitude|reflex|will|proficiency|remaster)\b/.test(
      lower
    )
  );
}

function isCampaignLookupInput(inputText) {
  const lower = str(inputText).toLowerCase();
  if (!lower) return false;
  return (
    /\b(who was|who is|what happened|what was the name|which faction|which npc|where did|when did|remind me|recall|what is our current|what is the current|what do we know about)\b/.test(
      lower
    ) ||
    /\b(last town|last session|three sessions ago|our kingdom problem|current kingdom problem|swamp town|mayor|bandits)\b/.test(lower)
  );
}

function isSessionSummaryInput(inputText) {
  const lower = str(inputText).toLowerCase();
  if (!lower) return false;
  return /\b(summarize|summary|recap|session notes|table log|clean these notes|turn these notes|write a recap|session wrap-up|wrap up)\b/.test(lower);
}

function inferEntityTargetFromInput(tabId, inputText) {
  if (tabId === "npcs") return "npc";
  if (tabId === "quests") return "quest";
  if (tabId === "locations") return "location";
  const lower = str(inputText).toLowerCase();
  if (/\bnpc\b/.test(lower)) return "npc";
  if (/\bquest\b/.test(lower)) return "quest";
  if (/\b(location|village|town|city|settlement|hex|place)\b/.test(lower)) return "location";
  return "";
}

function isNoteUpdateInput(tabId, inputText) {
  const lower = str(inputText).toLowerCase();
  if (!lower) return false;
  if (isStructuredWorldTab(tabId)) return !isCopilotSmallTalkInput(lower);
  return /\b(create|make|build|draft|invent|design|rewrite|update|turn this into|clean up|make an npc note|make a quest|make a location)\b/.test(lower);
}

function isKingdomHelperInput(tabId, inputText) {
  if (tabId === "kingdom") return true;
  const lower = str(inputText).toLowerCase();
  return /\b(kingdom|settlement|charter|government|heartland|consumption|control dc|ruin|unrest|economy|loyalty|stability|resource points|kingdom turn)\b/.test(lower);
}

function isMapHelperInput(tabId, inputText) {
  if (tabId === "hexmap") return true;
  const lower = str(inputText).toLowerCase();
  return /\b(hex|map|terrain|region|work site|party marker|encounter marker|event marker|road|river|travel route)\b/.test(lower);
}

function getCopilotTaskMeta(taskType) {
  return COPILOT_TASK_META[str(taskType)] || COPILOT_TASK_META.general_prep;
}

function buildRulesQuestionPrompt(requestText) {
  const compactInput = compactCopilotRequestText(requestText, 420);
  return [
    "Answer this as a Pathfinder 2e rules question using local grounded context first.",
    "Separate confirmed rules from house-rule or campaign inference.",
    "If the local context is thin, say that clearly instead of guessing.",
    "Return:",
    "Rules Answer:",
    "- 3 to 6 concise bullets",
    "Official vs Local Notes:",
    "- Confirmed official rule",
    "- House rule or campaign override if present",
    "Source Trail:",
    "- bullet",
    "",
    "GM question:",
    compactInput,
  ].join("\n");
}

function buildCampaignLookupPrompt(requestText) {
  const compactInput = compactCopilotRequestText(requestText, 420);
  return [
    "Answer this as a campaign-memory lookup grounded in app records, vault notes, and retrieved PDFs when relevant.",
    "Keep the answer concise and tied to actual campaign state.",
    "Return:",
    "Campaign Answer:",
    "- 2 to 5 bullets",
    "Key Linked Entities:",
    "- bullet",
    "Open Threads:",
    "- bullet",
    "",
    "GM question:",
    compactInput,
  ].join("\n");
}

function buildTaskRoutedRequest(tabId, userInput, autoRun) {
  const seedPrompt = buildGlobalCopilotSeedPrompt(tabId);
  const baseMode = getGlobalCopilotMode(tabId);
  const cleanInput = str(userInput);
  const meta = (taskType, extra = {}) => ({
    taskType,
    taskLabel: getCopilotTaskMeta(taskType).label,
    saveTarget: getCopilotTaskMeta(taskType).saveTarget,
    routeReason: str(extra.routeReason || ""),
    entityType: str(extra.entityType || ""),
  });

  const defaultTaskByTab = (() => {
    if (tabId === "sessions" || tabId === "capture") return "session_summary";
    if (tabId === "npcs" || tabId === "quests" || tabId === "locations") return "note_update";
    if (tabId === "rules") return "rules_question";
    if (tabId === "kingdom") return "kingdom_helper";
    if (tabId === "hexmap") return "map_helper";
    if (tabId === "pdf") return "pdf_lookup";
    if (tabId === "obsidian") return "vault_workflow";
    if (tabId === "foundry") return "foundry_handoff";
    if (tabId === "writing") return "writing_cleanup";
    return "general_prep";
  })();

  if (autoRun || !cleanInput) {
    return {
      mode: getCopilotTaskMeta(defaultTaskByTab).mode || baseMode,
      input: seedPrompt,
      isChat: false,
      ...meta(defaultTaskByTab, { routeReason: autoRun ? "automatic tab run" : "empty input defaults to tab workflow" }),
    };
  }

  if (isCopilotSmallTalkInput(cleanInput)) {
    return {
      mode: "assistant",
      input: cleanInput,
      isChat: true,
      ...meta("small_talk", { routeReason: "short conversational input" }),
    };
  }

  if (tabId === "obsidian") {
    return {
      mode: "prep",
      input: `${seedPrompt}\n\nRequest:\n${compactCopilotRequestText(cleanInput, 420)}`,
      isChat: false,
      ...meta("vault_workflow", { routeReason: "obsidian tab request" }),
    };
  }

  if (tabId === "foundry") {
    return {
      mode: "prep",
      input: `${seedPrompt}\n\nRequest:\n${compactCopilotRequestText(cleanInput, 380)}`,
      isChat: false,
      ...meta("foundry_handoff", { routeReason: "foundry export workflow" }),
    };
  }

  if (tabId === "writing") {
    return {
      mode: "session",
      input: `${seedPrompt}\n\nDraft to clean up:\n${cleanInput}`,
      isChat: false,
      ...meta("writing_cleanup", { routeReason: "writing helper request" }),
    };
  }

  if (tabId === "rules") {
    return {
      mode: "assistant",
      input: buildRulesQuestionPrompt(cleanInput),
      isChat: false,
      ...meta("rules_question", { routeReason: "rules tab workflow" }),
    };
  }

  if (tabId === "pdf" || isPdfGroundedQuestion(cleanInput)) {
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
      ...meta("pdf_lookup", { routeReason: "pdf-specific wording or pdf tab" }),
    };
  }

  if (isKingdomHelperInput(tabId, cleanInput)) {
    return {
      mode: "session",
      input: `${buildGlobalCopilotSeedPrompt("kingdom")}\n\nAdditional request:\n${compactCopilotRequestText(cleanInput, 420)}`,
      isChat: false,
      ...meta("kingdom_helper", { routeReason: "kingdom terms detected" }),
    };
  }

  if (isMapHelperInput(tabId, cleanInput)) {
    return {
      mode: "location",
      input: `${buildGlobalCopilotSeedPrompt("hexmap")}\n\nAdditional request:\n${compactCopilotRequestText(cleanInput, 420)}`,
      isChat: false,
      ...meta("map_helper", { routeReason: "hex map terms detected" }),
    };
  }

  if (isRulesQuestionInput(cleanInput)) {
    return {
      mode: "assistant",
      input: buildRulesQuestionPrompt(cleanInput),
      isChat: false,
      ...meta("rules_question", { routeReason: "rules question phrasing detected" }),
    };
  }

  if (isSessionSummaryInput(cleanInput) || tabId === "capture") {
    return {
      mode: "session",
      input: `${buildGlobalCopilotSeedPrompt(tabId === "capture" ? "capture" : "sessions")}\n\nSource text or request:\n${cleanInput}`,
      isChat: false,
      ...meta("session_summary", { routeReason: "summary / recap wording detected" }),
    };
  }

  if (isPdfGroundedWorldRequest(tabId, cleanInput)) {
    return {
      mode: baseMode,
      input: buildPdfGroundedWorldPrompt(tabId, cleanInput),
      isChat: false,
      ...meta("note_update", { routeReason: "world note request grounded in selected PDF", entityType: inferEntityTargetFromInput(tabId, cleanInput) }),
    };
  }

  const inferredMode = inferStructuredModeFromInput(cleanInput);
  if (inferredMode) {
    const targetTab = getSeedTabForMode(inferredMode);
    const compactInput = compactCopilotRequestText(cleanInput, inferredMode === "npc" ? 420 : 340);
    return {
      mode: inferredMode,
      input: `${buildGlobalCopilotSeedPrompt(targetTab)}\n${getStructuredWorldDetailInstruction(targetTab)}\n\nAdditional request:\n${compactInput}`,
      isChat: false,
      ...meta("note_update", { routeReason: "entity-creation wording detected", entityType: inferredMode }),
    };
  }

  if (isNoteUpdateInput(tabId, cleanInput)) {
    const entityType = inferEntityTargetFromInput(tabId, cleanInput);
    const targetTab = entityType ? getSeedTabForMode(entityType) : tabId;
    const compactInput = compactCopilotRequestText(cleanInput, targetTab === "npcs" ? 420 : 340);
    return {
      mode: entityType || baseMode,
      input: `${buildGlobalCopilotSeedPrompt(targetTab)}\n${getStructuredWorldDetailInstruction(targetTab)}\n\nAdditional request:\n${compactInput}`,
      isChat: false,
      ...meta("note_update", { routeReason: "note creation/update wording detected", entityType }),
    };
  }

  if (isCampaignLookupInput(cleanInput)) {
    return {
      mode: "assistant",
      input: buildCampaignLookupPrompt(cleanInput),
      isChat: false,
      ...meta("campaign_lookup", { routeReason: "campaign memory lookup wording detected" }),
    };
  }

  if (shouldUseCopilotChatMode(tabId, cleanInput)) {
    return {
      mode: "assistant",
      input: cleanInput,
      isChat: true,
      ...meta("campaign_lookup", { routeReason: "short direct GM question" }),
    };
  }

  if (tabId === "dashboard") {
    return {
      mode: baseMode,
      input: `${seedPrompt}\n\nGM request:\n${cleanInput}`,
      isChat: false,
      ...meta("general_prep", { routeReason: "dashboard prep request" }),
    };
  }

  return {
    mode: baseMode,
    input: `${seedPrompt}\n\nAdditional request:\n${cleanInput}`,
    isChat: false,
    ...meta(defaultTaskByTab, { routeReason: "tab workflow default" }),
  };
}

function shouldUseCopilotChatMode(tabId, inputText) {
  const text = str(inputText);
  if (!text) return false;
  const lower = text.toLowerCase();
  if (tabId === "hexmap" && /\b(hex|map|encounter|event|terrain|region|work site|building|settlement)\b/.test(lower)) {
    return false;
  }
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
  return buildTaskRoutedRequest(tabId, userInput, autoRun);
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
  if (tabId === "rules") {
    return [
      "Answer this as a Pathfinder 2e rules question using official rules context first and local rulings second.",
      "Return:",
      "Rules Answer:",
      "- 3 to 6 concise bullets",
      "Official vs Local Notes:",
      "- Confirmed official rule",
      "- Local override or house rule if present",
      "Source Trail:",
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
  if (tabId === "hexmap") {
    return [
      "Build one table-ready hex brief for the selected kingdom hex.",
      "Return:",
      "Hex Summary:",
      "- 2 to 4 bullets",
      "What The Party Notices First:",
      "- bullet",
      "- bullet",
      "Encounter / Event Seeds:",
      "- 3 concise hooks or complications",
      "What To Record On The Map:",
      "- status, terrain, work site, or marker ideas",
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
  if (tabId === "obsidian") {
    return [
      "Help organize DM Helper data into useful Obsidian markdown notes.",
      "Return concise, practical note structure advice or markdown examples only.",
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
  } else if (tabId === "rules") {
    const aiMemory = ensureAiMemoryState();
    const selectedRule = getSelectedRulesResult(ui.rulesResults);
    const recentRules = getRecentRuleCaptureEntries(4);
    const savedRules = ensureRulesStore();
    tabContext = `Rules query: ${ui.rulesSearchQuery || "(none)"} | Selected official rule: ${selectedRule?.title || "none"} | Official results loaded: ${
      Array.isArray(ui.rulesResults) ? ui.rulesResults.length : 0
    } | Local rulings digest: ${clipDigestText(aiMemory.rulingsDigest || "none", 320)} | Recent rule captures: ${
      recentRules.map((entry) => entry.note).filter(Boolean).join(" | ") || "none"
    } | Saved rules store entries: ${savedRules.length} | Canon memory digest: ${clipDigestText(aiMemory.canonSummary || "none", 220)}`;
  } else if (tabId === "kingdom") {
    const kingdom = getKingdomState();
    const profile = getActiveKingdomProfile();
    tabContext = `Kingdom: ${kingdom.name || "Unnamed kingdom"} | Turn: ${kingdom.currentTurnLabel || "Not set"} | Level ${kingdom.level} | Size ${kingdom.size} | Control DC ${
      kingdom.controlDC
    } | Unrest ${kingdom.unrest} | Renown ${kingdom.renown} | Fame ${kingdom.fame} | Infamy ${kingdom.infamy} | Settlements ${
      kingdom.settlements.length
    } | Claimed regions ${kingdom.regions.length} | Active profile: ${profile?.shortLabel || profile?.label || "Unknown"}`;
  } else if (tabId === "hexmap") {
    const hexMap = getHexMapState();
    const selectedHex = getHexMapSelectedHex(hexMap);
    const region = getKingdomRegionByHex(selectedHex);
    const markers = getHexMapMarkersForHex(selectedHex);
    const forces = getHexMapForcesForHex(selectedHex);
    const party = getHexMapParty(hexMap);
    tabContext = `Hex map: ${hexMap.mapName} | Selected hex: ${selectedHex} | Region status: ${region?.status || "none"} | Terrain: ${
      region?.terrain || "none"
    } | Work site: ${region?.workSite || "none"} | Markers in hex: ${markers.length} | Forces in hex: ${forces.length} | Party hex: ${
      party.hex || "not placed"
    } | Party trail points: ${party.trail.length} | Background loaded: ${
      hexMap.backgroundName || "no"
    }`;
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
  } else if (tabId === "obsidian") {
    const settings = ensureObsidianSettings();
    tabContext = `Obsidian vault path: ${settings.vaultPath || "(not set)"} | Base folder: ${settings.baseFolder || "DM Helper"} | AI read enabled: ${
      settings.useForAiContext ? "yes" : "no"
    } | Read scope: ${settings.readWholeVault ? "whole vault" : "DM Helper folder only"} | AI note folder: ${
      settings.aiWriteFolder || "AI Notes"
    } | Last sync: ${settings.lastSyncAt || "never"} | Last sync summary: ${settings.lastSyncSummary || "none"}`;
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
    aonRulesMatches: tabId === "rules" ? sortRulesResults(ui.rulesSearchQuery, ui.rulesResults).slice(0, 4) : [],
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

function extractRetrievalTerms(text) {
  return [...new Set(
    str(text)
      .toLowerCase()
      .replace(/[^a-z0-9\s]/g, " ")
      .split(/\s+/)
      .map((word) => word.trim())
      .filter((word) => word.length >= 3)
  )];
}

function countRetrievalTokenHits(text, terms) {
  const haystack = str(text).toLowerCase();
  if (!haystack || !Array.isArray(terms) || !terms.length) return 0;
  let hits = 0;
  for (const term of terms) {
    if (haystack.includes(term)) hits += 1;
  }
  return hits;
}

function getRetrievalProfile(taskType, entityType = "") {
  const sourceWeights = {
    general_prep: { session: 2.4, quest: 2.2, npc: 2.1, location: 2, kingdom: 1.8, hex: 1.8, rule: 1.5, canon: 2.3, aon_rule: 1.6, vault: 1.7, pdf_summary: 1.6 },
    rules_question: { rule: 3.8, canon: 1.6, aon_rule: 4.2, pdf_summary: 3.2, vault: 2.4, session: 0.8, quest: 0.7, npc: 0.6, location: 0.6, kingdom: 0.9, hex: 0.6 },
    campaign_lookup: { session: 3.2, quest: 3.1, npc: 3.1, location: 3, kingdom: 1.8, hex: 1.5, rule: 1, canon: 3.2, aon_rule: 1.1, vault: 2.2, pdf_summary: 1.2 },
    session_summary: { session: 3.4, quest: 2.4, npc: 2.2, location: 2.1, kingdom: 1.6, hex: 1.3, rule: 1.2, canon: 2.1, aon_rule: 1.1, vault: 1.8, pdf_summary: 1.4 },
    note_update: { session: 2, quest: 2, npc: 2, location: 2, kingdom: 1.4, hex: 1.2, rule: 1, canon: 2.4, aon_rule: 1.2, vault: 1.8, pdf_summary: 2.2 },
    kingdom_helper: { kingdom: 4, hex: 2.6, session: 1.8, quest: 1.2, npc: 1.1, location: 1.4, rule: 1.6, canon: 1.6, aon_rule: 1.3, vault: 1.8, pdf_summary: 1.3 },
    map_helper: { hex: 4, location: 2.9, kingdom: 2.4, session: 1.4, quest: 1.2, npc: 1.1, rule: 1, canon: 1.4, aon_rule: 0.9, vault: 1.5, pdf_summary: 1.1 },
    pdf_lookup: { pdf_summary: 4, rule: 2, canon: 1, aon_rule: 1.4, vault: 1.8, session: 1.1, quest: 1, npc: 1, location: 1, kingdom: 0.9, hex: 0.8 },
    vault_workflow: { vault: 4, session: 2, quest: 1.6, npc: 1.6, location: 1.6, kingdom: 1.4, hex: 1.2, rule: 1.2, canon: 1.2, aon_rule: 1.1, pdf_summary: 1.2 },
    foundry_handoff: { session: 2.8, quest: 2, npc: 2, location: 2, kingdom: 1.2, hex: 1.1, rule: 0.8, canon: 1.8, aon_rule: 0.8, vault: 1.3, pdf_summary: 1 },
    writing_cleanup: { session: 2.4, quest: 1.8, npc: 1.8, location: 1.8, kingdom: 1.2, hex: 1, rule: 1, canon: 1.5, aon_rule: 0.8, vault: 1.4, pdf_summary: 1.1 },
    small_talk: { session: 0, quest: 0, npc: 0, location: 0, kingdom: 0, hex: 0, rule: 0, canon: 0, aon_rule: 0, vault: 0, pdf_summary: 0 },
  };

  const profile = {
    maxChunks: 8,
    minScore: 1.4,
    sourceWeights: sourceWeights[taskType] || sourceWeights.general_prep,
  };

  if (taskType === "rules_question") profile.maxChunks = 7;
  if (taskType === "campaign_lookup") profile.maxChunks = 9;
  if (taskType === "note_update" && entityType === "npc") profile.maxChunks = 7;
  if (taskType === "note_update" && entityType === "location") profile.maxChunks = 7;
  if (taskType === "small_talk") profile.maxChunks = 0;

  return profile;
}

function buildUnifiedRetrievalQuery({ userInput, route, baseContext, tabId }) {
  const parts = [
    str(userInput),
    str(route?.taskLabel),
    str(route?.entityType),
    str(baseContext?.tabContext),
    str(baseContext?.latestSession?.title),
    str(baseContext?.latestSession?.summary),
    tabId === "pdf" ? str(baseContext?.selectedPdfFile) : "",
  ]
    .map((value) => value.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  return parts.join(" ").trim();
}

function computeRetrievalCandidateScore({ query, terms, title, text, sourceType, taskType, entityType, updatedAt }) {
  const combined = `${str(title)} ${str(text)}`.trim();
  if (!combined) return 0;
  const lowerCombined = combined.toLowerCase();
  const phraseScore = relevanceScore(query, title) * 1.8 + relevanceScore(query, combined) * 0.6;
  const tokenHits = countRetrievalTokenHits(lowerCombined, terms);
  const sourceWeight = Number(getRetrievalProfile(taskType, entityType).sourceWeights?.[sourceType] || 0);
  const updatedKey = safeDate(updatedAt);
  const recencyBoost = Number.isFinite(updatedKey) ? Math.max(0, Math.min(1.4, (updatedKey - (Date.now() - 1000 * 60 * 60 * 24 * 21)) / (1000 * 60 * 60 * 24 * 21))) : 0;
  return Number((phraseScore + tokenHits * 1.15 + sourceWeight + recencyBoost).toFixed(2));
}

function summarizeHexMapForRetrieval() {
  const hexMap = getHexMapState();
  const selectedHex = getHexMapSelectedHex(hexMap);
  const selectedRegion = getKingdomRegionByHex(selectedHex);
  const markers = getHexMapMarkersForHex(selectedHex);
  const forces = getHexMapForcesForHex(selectedHex);
  const party = getHexMapParty(hexMap);
  return {
    title: `Hex Map ${selectedHex || "(no selected hex)"}`,
    text: [
      `Map ${hexMap.mapName || "Hex Map"}`,
      `Selected hex ${selectedHex || "none"}`,
      `Region status ${selectedRegion?.status || "none"}`,
      `Terrain ${selectedRegion?.terrain || "none"}`,
      `Work site ${selectedRegion?.workSite || "none"}`,
      `Markers ${markers.map((marker) => `${marker.type}: ${marker.notes || "no notes"}`).join(" | ") || "none"}`,
      `Forces ${forces.map((force) => `${force.type}: ${force.notes || "no notes"}`).join(" | ") || "none"}`,
      `Party hex ${party.hex || "not placed"}`,
      `Party trail ${Array.isArray(party.trail) ? party.trail.join(" -> ") : "none"}`,
    ].join(" • "),
    updatedAt: hexMap.updatedAt || "",
  };
}

function buildUnifiedRetrievalContext({ tabId, userInput, route, baseContext, obsidianNotes = [] }) {
  const taskType = str(route?.taskType || baseContext?.taskType || "general_prep");
  const entityType = str(route?.entityType || baseContext?.entityType || "");
  const profile = getRetrievalProfile(taskType, entityType);
  const query = buildUnifiedRetrievalQuery({ userInput, route, baseContext, tabId });
  const terms = extractRetrievalTerms(query);
  const candidates = [];

  if (!profile.maxChunks || taskType === "small_talk") {
    return {
      retrievalSummary: {
        taskType,
        taskLabel: str(route?.taskLabel || baseContext?.taskLabel || ""),
        query,
        generatedAt: new Date().toLocaleTimeString(),
        sourceCounts: {},
        chunkCount: 0,
      },
      retrievedChunks: [],
    };
  }

  const pushCandidate = ({ sourceType, sourceLabel, title, text, updatedAt, reason }) => {
    const cleanTitle = str(title);
    const cleanText = str(text).replace(/\s+/g, " ").trim();
    if (!cleanTitle || !cleanText) return;
    const score = computeRetrievalCandidateScore({
      query,
      terms,
      title: cleanTitle,
      text: cleanText,
      sourceType,
      taskType,
      entityType,
      updatedAt,
    });
    if (score < profile.minScore) return;
    candidates.push({
      sourceType,
      sourceLabel: str(sourceLabel || sourceType),
      title: cleanTitle,
      text: compactLine(cleanText, 420),
      updatedAt: str(updatedAt),
      score,
      reason: str(reason),
    });
  };

  const recentSessions = [...(state.sessions || [])]
    .sort((a, b) => safeDate(b.date || b.updatedAt || b.createdAt) - safeDate(a.date || a.updatedAt || a.createdAt))
    .slice(0, 8);
  for (const session of recentSessions) {
    pushCandidate({
      sourceType: "session",
      sourceLabel: "Session",
      title: session.title || "Untitled session",
      text: [
        `Summary ${session.summary || "none"}`,
        `Next prep ${session.nextPrep || "none"}`,
        `Arc ${session.arc || "none"}`,
      ].join(" • "),
      updatedAt: session.date || session.updatedAt || session.createdAt || "",
      reason: "recent session record",
    });
  }

  for (const quest of state.quests || []) {
    pushCandidate({
      sourceType: "quest",
      sourceLabel: "Quest",
      title: quest.title || "Untitled quest",
      text: [
        `Status ${quest.status || "unknown"}`,
        `Objective ${quest.objective || "none"}`,
        `Giver ${quest.giver || "none"}`,
        `Stakes ${quest.stakes || "none"}`,
      ].join(" • "),
      updatedAt: quest.updatedAt || quest.createdAt || "",
      reason: "quest record",
    });
  }

  for (const npc of state.npcs || []) {
    pushCandidate({
      sourceType: "npc",
      sourceLabel: "NPC",
      title: npc.name || "Unnamed NPC",
      text: [
        `Role ${npc.role || "none"}`,
        `Agenda ${npc.agenda || "none"}`,
        `Disposition ${npc.disposition || "none"}`,
        `Notes ${npc.notes || "none"}`,
      ].join(" • "),
      updatedAt: npc.updatedAt || npc.createdAt || "",
      reason: "NPC record",
    });
  }

  for (const location of state.locations || []) {
    pushCandidate({
      sourceType: "location",
      sourceLabel: "Location",
      title: location.name || "Unnamed location",
      text: [
        `Hex ${location.hex || "none"}`,
        `Changed ${location.whatChanged || "none"}`,
        `Notes ${location.notes || "none"}`,
      ].join(" • "),
      updatedAt: location.updatedAt || location.createdAt || "",
      reason: "location record",
    });
  }

  const kingdom = getKingdomState();
  pushCandidate({
    sourceType: "kingdom",
    sourceLabel: "Kingdom",
    title: kingdom.name || "Kingdom Sheet",
    text: [
      `Turn ${kingdom.currentTurnLabel || "none"}`,
      `Level ${kingdom.level || 0}`,
      `Control DC ${kingdom.controlDC || 0}`,
      `Unrest ${kingdom.unrest || 0}`,
      `Resources RP ${kingdom.resourcePoints || 0} Food ${kingdom.food || 0} Lumber ${kingdom.lumber || 0} Ore ${kingdom.ore || 0} Stone ${kingdom.stone || 0} Luxuries ${kingdom.luxuries || 0}`,
      `Settlements ${(kingdom.settlements || []).map((item) => item.name).filter(Boolean).join(", ") || "none"}`,
      `Notes ${kingdom.notes || "none"}`,
    ].join(" • "),
    updatedAt: kingdom.updatedAt || "",
    reason: "kingdom state",
  });

  const hexSummary = summarizeHexMapForRetrieval();
  pushCandidate({
    sourceType: "hex",
    sourceLabel: "Hex Map",
    title: hexSummary.title,
    text: hexSummary.text,
    updatedAt: hexSummary.updatedAt,
    reason: "hex map state",
  });

  const ruleEntries = [...(state.liveCapture || [])]
    .filter((entry) => {
      const kind = str(entry?.kind).toLowerCase();
      return kind === "rule" || kind === "retcon" || kind === "house rule";
    })
    .sort((a, b) => safeDate(b.timestamp) - safeDate(a.timestamp))
    .slice(0, 10);
  for (const entry of ruleEntries) {
    pushCandidate({
      sourceType: "rule",
      sourceLabel: "Ruling",
      title: `${entry.kind || "Rule"} ${formatAiHistoryTimestamp(entry.timestamp) || ""}`.trim(),
      text: entry.note || "",
      updatedAt: entry.timestamp || "",
      reason: "captured ruling",
    });
  }

  for (const entry of ensureRulesStore()) {
    pushCandidate({
      sourceType: entry.kind === "canon_memory" ? "canon" : "rule",
      sourceLabel: entry.kind === "canon_memory" ? "Canon Memory" : "Saved Rule",
      title: `${getRuleStoreKindLabel(entry.kind)} • ${entry.title}`.trim(),
      text: [entry.text || "", entry.sourceTitle ? `Source ${entry.sourceTitle}` : "", entry.tags?.length ? `Tags ${entry.tags.join(", ")}` : ""]
        .filter(Boolean)
        .join(" • "),
      updatedAt: entry.updatedAt || entry.createdAt || "",
      reason: entry.kind === "canon_memory" ? "saved canon memory" : "saved rule store",
    });
  }

  for (const summary of Object.values(getPdfSummaryMap())) {
    pushCandidate({
      sourceType: "pdf_summary",
      sourceLabel: "PDF Summary",
      title: str(summary?.fileName || "Indexed PDF"),
      text: str(summary?.summary || ""),
      updatedAt: str(summary?.updatedAt || summary?.summaryUpdatedAt || ""),
      reason: "saved PDF memory",
    });
  }

  for (const note of Array.isArray(obsidianNotes) ? obsidianNotes : []) {
    pushCandidate({
      sourceType: "vault",
      sourceLabel: "Vault Note",
      title: str(note?.title || note?.relativePath || "Vault note"),
      text: [str(note?.relativePath), str(note?.excerpt)].filter(Boolean).join(" • "),
      updatedAt: str(note?.modifiedAt || ""),
      reason: "vault retrieval",
    });
  }

  for (const match of Array.isArray(baseContext?.aonRulesMatches) ? baseContext.aonRulesMatches : []) {
    pushCandidate({
      sourceType: "aon_rule",
      sourceLabel: "AoN Rule",
      title: str(match?.title || "PF2e rule"),
      text: [str(match?.snippet), str(match?.url)].filter(Boolean).join(" • "),
      updatedAt: "",
      reason: "live AoN rules lookup",
    });
  }

  candidates.sort((a, b) => b.score - a.score || a.title.localeCompare(b.title));
  const deduped = [];
  const seen = new Set();
  for (const candidate of candidates) {
    const key = `${candidate.sourceType}::${candidate.title}`.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push(candidate);
    if (deduped.length >= profile.maxChunks) break;
  }

  const sourceCounts = {};
  for (const chunk of deduped) {
    sourceCounts[chunk.sourceType] = (sourceCounts[chunk.sourceType] || 0) + 1;
  }

  return {
    retrievalSummary: {
      taskType,
      taskLabel: str(route?.taskLabel || baseContext?.taskLabel || ""),
      query: compactLine(query, 220),
      generatedAt: new Date().toLocaleTimeString(),
      sourceCounts,
      chunkCount: deduped.length,
    },
    retrievedChunks: deduped,
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
  if (activeTab === "obsidian") return;
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
  const runtimeContext = await buildCopilotRuntimeContext(
    activeTab,
    userInput || input,
    isSmallTalk
      ? { baseContext: buildMinimalCopilotContext(activeTab), skipVault: true, route: request }
      : { route: request }
  );
  ui.copilotRetrievalPreview = runtimeContext?.retrievalSummary
    ? {
        ...runtimeContext.retrievalSummary,
        chunks: Array.isArray(runtimeContext?.retrievedChunks) ? runtimeContext.retrievedChunks : [],
      }
    : null;
  const vaultNoteCount = Array.isArray(runtimeContext?.obsidianVaultNotes) ? runtimeContext.obsidianVaultNotes.length : 0;
  const aonRuleCount = Array.isArray(runtimeContext?.aonRulesMatches) ? runtimeContext.aonRulesMatches.length : 0;
  const vaultContextSuffix = vaultNoteCount ? ` Used vault context from ${vaultNoteCount} note${vaultNoteCount === 1 ? "" : "s"}.` : "";
  const aonContextSuffix = aonRuleCount ? ` Used AoN rules context from ${aonRuleCount} rule page${aonRuleCount === 1 ? "" : "s"}.` : "";
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
      contextOverride: runtimeContext,
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
          contextOverride: runtimeContext,
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
      const selectedPdfFile = str(runtimeContext?.selectedPdfFile) || "the selected PDF";
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
        contextOverride: runtimeContext,
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
        contextOverride: runtimeContext,
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
      ui.copilotMessage = `Recovered reply using ${getAiModelDisplayName(recoveredWithModel)}. Default model switched to this one.${vaultContextSuffix}`;
      ui.copilotMessage += aonContextSuffix;
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
      ui.copilotMessage += aonContextSuffix;
    } else {
      ui.copilotPendingFallbackMemory = null;
      if (request.isChat) {
        ui.copilotMessage = `Chat response generated with ${getAiModelDisplayName(str(finalAttempt.result?.model) || effectiveConfig.model)}.`;
      } else {
        ui.copilotMessage = autoRun
          ? `Auto-generated ${getTabLabel(activeTab)} output.`
          : "Generated.";
      }
      ui.copilotMessage += vaultContextSuffix;
      ui.copilotMessage += aonContextSuffix;
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
  if (activeTab === "hexmap") {
    appendHexMapAiNote(text);
    ui.copilotMessage = `Appended AI output to ${getHexMapSelectedHex()} notes.`;
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
  if (activeTab === "obsidian") {
    ui.copilotMessage = "No direct apply target on the Obsidian tab. Use Copy or Sync To Vault.";
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

function normalizeAiMemoryState(rawMemory) {
  const sourceCounts =
    rawMemory?.sourceCounts && typeof rawMemory.sourceCounts === "object" && !Array.isArray(rawMemory.sourceCounts)
      ? rawMemory.sourceCounts
      : {};
  return {
    campaignSummary: str(rawMemory?.campaignSummary),
    recentSessionSummary: str(rawMemory?.recentSessionSummary),
    activeQuestsSummary: str(rawMemory?.activeQuestsSummary),
    activeEntitiesSummary: str(rawMemory?.activeEntitiesSummary),
    canonSummary: str(rawMemory?.canonSummary),
    rulingsDigest: str(rawMemory?.rulingsDigest),
    manualRulings: str(rawMemory?.manualRulings),
    updatedAt: str(rawMemory?.updatedAt),
    sourceCounts: {
      sessions: Number.parseInt(String(sourceCounts.sessions || "0"), 10) || 0,
      openQuests: Number.parseInt(String(sourceCounts.openQuests || "0"), 10) || 0,
      npcs: Number.parseInt(String(sourceCounts.npcs || "0"), 10) || 0,
      locations: Number.parseInt(String(sourceCounts.locations || "0"), 10) || 0,
      ruleEntries: Number.parseInt(String(sourceCounts.ruleEntries || "0"), 10) || 0,
      canonEntries: Number.parseInt(String(sourceCounts.canonEntries || "0"), 10) || 0,
    },
  };
}

function ensureAiMemoryState() {
  if (!state?.meta) return normalizeAiMemoryState({});
  state.meta.aiMemory = normalizeAiMemoryState(state.meta.aiMemory);
  return state.meta.aiMemory;
}

function clipDigestText(value, limit = 560) {
  const clean = str(value).replace(/\s+/g, " ").trim();
  if (clean.length <= limit) return clean;
  const sliced = clean.slice(0, limit);
  const breakAt = Math.max(sliced.lastIndexOf(". "), sliced.lastIndexOf("; "), sliced.lastIndexOf(", "));
  if (breakAt > Math.floor(limit * 0.6)) {
    return `${sliced.slice(0, breakAt + 1).trim()} ...`;
  }
  return `${sliced.trim()}...`;
}

function normalizeRulesStoreEntry(rawEntry = {}) {
  const kindRaw = str(rawEntry?.kind).toLowerCase();
  const kind = Object.prototype.hasOwnProperty.call(RULE_STORE_KIND_LABELS, kindRaw) ? kindRaw : "accepted_ruling";
  const text = str(rawEntry?.text || rawEntry?.body || rawEntry?.note).slice(0, 8000);
  const titleSeed = str(rawEntry?.title || text.split(/\n+/)[0] || "").replace(/\s+/g, " ").trim();
  const title = titleSeed || `${RULE_STORE_KIND_LABELS[kind]} ${formatAiHistoryTimestamp(rawEntry?.updatedAt || rawEntry?.createdAt) || ""}`.trim();
  const tags = Array.isArray(rawEntry?.tags)
    ? rawEntry.tags
    : str(rawEntry?.tags)
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean);
  return {
    id: str(rawEntry?.id) || `rule-store-${uid()}`,
    title: title.slice(0, 180),
    kind,
    text,
    sourceTitle: str(rawEntry?.sourceTitle || ""),
    sourceUrl: str(rawEntry?.sourceUrl || ""),
    sourceOrigin: str(rawEntry?.sourceOrigin || ""),
    tags: [...new Set(tags.map((item) => str(item).toLowerCase()).filter(Boolean))].slice(0, 12),
    createdAt: str(rawEntry?.createdAt) || new Date().toISOString(),
    updatedAt: str(rawEntry?.updatedAt || rawEntry?.createdAt) || new Date().toISOString(),
  };
}

function normalizeRulesStore(rawStore) {
  return Array.isArray(rawStore) ? rawStore.map((entry) => normalizeRulesStoreEntry(entry)).filter((entry) => entry.text) : [];
}

function ensureRulesStore() {
  state.rulesStore = normalizeRulesStore(state.rulesStore);
  return state.rulesStore;
}

function getRuleStoreKindLabel(kind) {
  return RULE_STORE_KIND_LABELS[str(kind).toLowerCase()] || "Saved Entry";
}

function buildRuleStoreTags({ title = "", text = "", kind = "" } = {}) {
  const seeds = [str(title), str(text), str(kind)].join(" ");
  return extractRetrievalTerms(seeds).slice(0, 8);
}

function upsertRulesStoreEntry(input = {}) {
  const store = ensureRulesStore();
  const normalized = normalizeRulesStoreEntry({
    ...input,
    tags: Array.isArray(input.tags) && input.tags.length ? input.tags : buildRuleStoreTags(input),
  });
  const titleKey = str(normalized.title).toLowerCase();
  const kindKey = str(normalized.kind).toLowerCase();
  const existing = store.find((entry) => str(entry.title).toLowerCase() === titleKey && str(entry.kind).toLowerCase() === kindKey);
  if (existing) {
    existing.text = normalized.text;
    existing.sourceTitle = normalized.sourceTitle || existing.sourceTitle;
    existing.sourceUrl = normalized.sourceUrl || existing.sourceUrl;
    existing.sourceOrigin = normalized.sourceOrigin || existing.sourceOrigin;
    existing.tags = [...new Set([...(existing.tags || []), ...(normalized.tags || [])])].slice(0, 12);
    existing.updatedAt = new Date().toISOString();
    return existing;
  }
  store.unshift(normalized);
  state.rulesStore = normalizeRulesStore(store).slice(0, 240);
  return normalized;
}

function deleteRulesStoreEntry(id) {
  const targetId = str(id);
  if (!targetId) return false;
  const before = ensureRulesStore().length;
  state.rulesStore = ensureRulesStore().filter((entry) => entry.id !== targetId);
  return state.rulesStore.length !== before;
}

function buildAiMemoryDigests(sourceState = state) {
  const working = sourceState || createStarterState();
  const previousMemory = normalizeAiMemoryState(working?.meta?.aiMemory);
  const sessions = [...(Array.isArray(working.sessions) ? working.sessions : [])].sort(
    (a, b) => safeDate(b.date || b.updatedAt || b.createdAt) - safeDate(a.date || a.updatedAt || a.createdAt)
  );
  const latest = sessions[0] || null;
  const openQuests = (Array.isArray(working.quests) ? working.quests : []).filter((q) => q.status !== "completed" && q.status !== "failed");
  const npcs = Array.isArray(working.npcs) ? working.npcs : [];
  const locations = Array.isArray(working.locations) ? working.locations : [];
  const hexMap = working === state ? getHexMapState() : normalizeHexMapState(working.hexMap);
  const party = getHexMapParty(hexMap);
  const kingdom = working === state ? getKingdomState() : normalizeKingdomState(working.kingdom);
  const recentRuleEntries = [...(Array.isArray(working.liveCapture) ? working.liveCapture : [])]
    .filter((entry) => {
      const kind = str(entry?.kind).toLowerCase();
      return kind === "rule" || kind === "retcon";
    })
    .sort((a, b) => safeDate(b.timestamp) - safeDate(a.timestamp))
    .slice(0, 6);
  const rulesStore = working === state ? ensureRulesStore() : normalizeRulesStore(working.rulesStore);
  const savedRuleEntries = rulesStore
    .filter((entry) => entry.kind !== "canon_memory")
    .sort((a, b) => safeDate(b.updatedAt || b.createdAt) - safeDate(a.updatedAt || a.createdAt))
    .slice(0, 8);
  const canonEntries = rulesStore
    .filter((entry) => entry.kind === "canon_memory")
    .sort((a, b) => safeDate(b.updatedAt || b.createdAt) - safeDate(a.updatedAt || a.createdAt))
    .slice(0, 8);

  const sessionHeadline = latest
    ? `${latest.title || "Latest session"}: ${clipDigestText(latest.summary || latest.nextPrep || "No summary yet.", 240)}`
    : "No sessions recorded yet.";
  const pressureLine = openQuests.length
    ? `Open pressure: ${openQuests
        .slice(0, 4)
        .map((quest) => `${quest.title}${quest.stakes ? ` (${clipDigestText(quest.stakes, 90)})` : ""}`)
        .join("; ")}`
    : "Open pressure: none tracked.";
  const kingdomLine = kingdom?.name || kingdom?.currentTurnLabel || Number(kingdom?.level || 0) > 0
    ? `Kingdom: ${kingdom.name || "Unnamed kingdom"} | Turn ${kingdom.currentTurnLabel || "not set"} | Control DC ${kingdom.controlDC} | Unrest ${kingdom.unrest} | Size ${kingdom.size}`
    : "Kingdom: no active kingdom state.";
  const mapLine = party?.hex
    ? `Party position: ${party.hex} with ${Array.isArray(party.trail) ? party.trail.length : 0} trail point(s).`
    : "Party position: not currently placed on the hex map.";
  const campaignSummary = [sessionHeadline, pressureLine, kingdomLine, mapLine].join("\n");

  const recentSessionSummary = sessions.length
    ? sessions
        .slice(0, 4)
        .map((session, index) => `${index + 1}. ${session.title || "Session"}${session.date ? ` (${session.date})` : ""} - ${clipDigestText(session.summary || session.nextPrep || "No summary yet.", 180)}`)
        .join("\n")
    : "No recent sessions recorded.";

  const activeQuestsSummary = openQuests.length
    ? openQuests
        .slice(0, 6)
        .map((quest, index) => `${index + 1}. ${quest.title} - ${clipDigestText(quest.objective || quest.stakes || quest.giver || "No objective yet.", 160)}`)
        .join("\n")
    : "No active quests tracked.";

  const latestNarrative = [latest?.summary, latest?.nextPrep, ...sessions.slice(0, 3).map((session) => session.summary || session.nextPrep || "")]
    .map((item) => str(item))
    .join(" ");
  const mentionedNpcNames = findEntityMentions(latestNarrative, npcs, "name", 5);
  const mentionedLocationNames = findEntityMentions(latestNarrative, locations, "name", 4);
  const mentionedQuestTitles = findEntityMentions(latestNarrative, openQuests, "title", 4);
  const fallbackNpcNames = [...npcs]
    .sort((a, b) => safeDate(b.updatedAt || b.createdAt) - safeDate(a.updatedAt || a.createdAt))
    .map((npc) => str(npc.name))
    .filter(Boolean)
    .slice(0, 5);
  const fallbackLocationNames = [...locations]
    .sort((a, b) => safeDate(b.updatedAt || b.createdAt) - safeDate(a.updatedAt || a.createdAt))
    .map((location) => str(location.name))
    .filter(Boolean)
    .slice(0, 4);
  const activeEntitiesSummary = [
    `NPCs: ${(mentionedNpcNames.length ? mentionedNpcNames : fallbackNpcNames).join(", ") || "None tracked"}`,
    `Locations: ${(mentionedLocationNames.length ? mentionedLocationNames : fallbackLocationNames).join(", ") || "None tracked"}`,
    `Quests: ${mentionedQuestTitles.join(", ") || openQuests.slice(0, 4).map((quest) => quest.title).join(", ") || "None tracked"}`,
  ].join("\n");

  const derivedRulings = recentRuleEntries.length
    ? recentRuleEntries
        .map((entry, index) => `${index + 1}. ${clipDigestText(entry.note || "", 180)}`)
        .join("\n")
    : "No Rule or Retcon capture entries recorded yet.";
  const savedRulingsDigest = savedRuleEntries.length
    ? savedRuleEntries
        .map((entry, index) => `${index + 1}. ${getRuleStoreKindLabel(entry.kind)}: ${clipDigestText(entry.text || entry.title || "", 180)}`)
        .join("\n")
    : "";
  const rulingsDigest = previousMemory.manualRulings
    ? clipDigestText(previousMemory.manualRulings, 2200)
    : [savedRulingsDigest, derivedRulings].filter(Boolean).join("\n");
  const canonSummary = canonEntries.length
    ? canonEntries.map((entry, index) => `${index + 1}. ${entry.title} - ${clipDigestText(entry.text || "", 180)}`).join("\n")
    : "No canon memory saved yet.";

  return {
    ...previousMemory,
    campaignSummary,
    recentSessionSummary,
    activeQuestsSummary,
    activeEntitiesSummary,
    canonSummary,
    rulingsDigest,
    updatedAt: new Date().toISOString(),
    sourceCounts: {
      sessions: sessions.length,
      openQuests: openQuests.length,
      npcs: npcs.length,
      locations: locations.length,
      ruleEntries: recentRuleEntries.length + savedRuleEntries.length,
      canonEntries: canonEntries.length,
    },
  };
}

function refreshAiMemoryDigests(options = {}) {
  const nextMemory = buildAiMemoryDigests(state);
  state.meta.aiMemory = nextMemory;
  if (options.persist) {
    saveState();
  }
  if (!options.silent) {
    ui.dashboardMessage = `AI memory refreshed at ${nextMemory.updatedAt}.`;
    render();
  }
  return nextMemory;
}

function collectAiCampaignContext() {
  const aiMemory = buildAiMemoryDigests(state);
  state.meta.aiMemory = aiMemory;
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
    aiMemory: {
      campaignSummary: aiMemory.campaignSummary,
      recentSessionSummary: aiMemory.recentSessionSummary,
      activeQuestsSummary: aiMemory.activeQuestsSummary,
      activeEntitiesSummary: aiMemory.activeEntitiesSummary,
      canonSummary: aiMemory.canonSummary,
      rulingsDigest: aiMemory.rulingsDigest,
      updatedAt: aiMemory.updatedAt,
    },
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
  if (tabId === "rules") {
    return [
      "Rules Answer:",
      "- Review the selected official PF2e rule match first.",
      "- Compare it against the local rulings digest before making a table ruling.",
      "- If the official source and local digest conflict, state both clearly and choose one before play resumes.",
      "",
      "Official vs Local Notes:",
      "- Official rule: confirm on the selected Archives of Nethys page.",
      "- Local override: check the Manual Rulings Digest and recent Rule / Retcon captures.",
      "",
      "Source Trail:",
      "- PF2e Rules tab search results",
      "- Local Rulings Digest",
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
  const fields = getFormFields(form);

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

  if (type === "hexmap-settings") {
    applyHexMapSettings(fields);
    saveState();
    ui.hexMapMessage = "Hex map settings updated.";
    render();
    return;
  }

  if (type === "obsidian-settings") {
    const settings = ensureObsidianSettings();
    settings.vaultPath = str(fields.vaultPath);
    settings.baseFolder = str(fields.baseFolder) || "DM Helper";
    settings.aiWriteFolder = str(fields.aiWriteFolder) || "AI Notes";
    settings.useForAiContext = fields.useForAiContext === "on";
    settings.readWholeVault = fields.readWholeVault === "on";
    settings.aiContextNoteLimit = Math.max(1, Math.min(12, Number.parseInt(String(fields.aiContextNoteLimit || settings.aiContextNoteLimit || 6), 10) || 6));
    settings.aiContextCharLimit = Math.max(
      800,
      Math.min(12000, Number.parseInt(String(fields.aiContextCharLimit || settings.aiContextCharLimit || 3600), 10) || 3600)
    );
    saveState();
    ui.obsidianMessage = "Obsidian vault settings saved.";
    render();
    return;
  }

  if (type === "ai-memory-rulings") {
    const aiMemory = ensureAiMemoryState();
    aiMemory.manualRulings = str(fields.manualRulings);
    state.meta.aiMemory = buildAiMemoryDigests(state);
    saveState();
    ui.dashboardMessage = "Manual rulings digest saved.";
    ui.rulesMessage = "Manual rulings digest saved.";
    render();
    return;
  }

  if (type === "rules-store-entry") {
    const text = str(fields.text);
    if (!text) {
      ui.rulesMessage = "Add some text before saving a local rule or canon entry.";
      render();
      return;
    }
    const entry = upsertRulesStoreEntry({
      title: str(fields.title),
      kind: str(fields.kind),
      text,
      sourceUrl: str(fields.sourceUrl),
      sourceOrigin: "manual",
    });
    state.meta.aiMemory = buildAiMemoryDigests(state);
    saveState();
    form.reset();
    ui.rulesMessage = `Saved ${getRuleStoreKindLabel(entry.kind).toLowerCase()}: ${entry.title}.`;
    render();
    return;
  }

  if (type === "rules-search") {
    ui.rulesSearchQuery = normalizeRulesSearchQuery(fields.query);
    ui.rulesSearchLimit = Math.max(1, Math.min(6, Number.parseInt(String(fields.limit || ui.rulesSearchLimit || 5), 10) || 5));
    ui.rulesScope = ["both", "official", "local"].includes(str(fields.scope)) ? str(fields.scope) : "both";
    await runRulesSearch(ui.rulesSearchQuery, ui.rulesSearchLimit, false);
    return;
  }

  if (type === "hexmap-region") {
    const record = upsertHexMapRegion(fields);
    setHexMapSelectedHex(record.hex);
    saveState();
    ui.hexMapMessage = `Saved hex record for ${record.hex}.`;
    render();
    return;
  }

  if (type === "hexmap-marker") {
    const marker = createHexMapMarker(fields);
    setHexMapSelectedHex(marker.hex);
    saveState();
    form.reset();
    ui.hexMapMessage = `Added ${marker.type.toLowerCase()} marker to ${marker.hex}.`;
    render();
    return;
  }

  if (type === "hexmap-force") {
    const force = createHexMapForce(fields);
    setHexMapSelectedHex(force.hex);
    saveState();
    form.reset();
    ui.hexMapMessage = `Added ${force.type.toLowerCase()} to ${force.hex}.`;
    render();
    return;
  }

  if (type === "hexmap-party") {
    const party = upsertHexMapParty(fields);
    setHexMapSelectedHex(party.hex);
    saveState();
    ui.hexMapMessage = `${party.label || "Party"} moved to ${party.hex}.`;
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
  if (clean === "hexMapMarkers") return getHexMapState().markers;
  if (clean === "hexMapForces") return getHexMapState().forces;
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
    if ((cleanCollection === "kingdomRegions" || cleanCollection === "hexMapMarkers" || cleanCollection === "hexMapForces") && field === "hex") {
      out[field] = normalizeHexCoordinate(value) || str(value);
      continue;
    }
    if (cleanCollection === "hexMapMarkers" && field === "type") {
      out[field] = HEX_MAP_MARKER_TYPES.includes(str(value)) ? str(value) : "Note";
      continue;
    }
    if (cleanCollection === "hexMapForces" && field === "type") {
      out[field] = HEX_MAP_FORCE_TYPES.includes(str(value)) ? str(value) : "Allied Force";
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
  if (activeTab === "hexmap" && (collection === "hexMapMarkers" || collection === "hexMapForces" || collection === "kingdomRegions")) {
    render();
  }
}

function getPdfSummaryMap() {
  if (!state.meta.pdfSummaries || typeof state.meta.pdfSummaries !== "object" || Array.isArray(state.meta.pdfSummaries)) {
    state.meta.pdfSummaries = {};
  }
  return state.meta.pdfSummaries;
}

async function runRulesSearch(query, limit = 5, force = false) {
  const cleanQuery = normalizeRulesSearchQuery(query);
  ui.rulesSearchQuery = cleanQuery;
  ui.rulesSearchLimit = Math.max(1, Math.min(6, Number.parseInt(String(limit || ui.rulesSearchLimit || 5), 10) || 5));

  if (!cleanQuery) {
    ui.rulesResults = [];
    ui.rulesSelectedUrl = "";
    ui.rulesMessage = "Type a PF2e rule term first.";
    render();
    return;
  }

  if (!desktopApi?.searchAonRules) {
    ui.rulesResults = [];
    ui.rulesSelectedUrl = "";
    ui.rulesMessage = "Official PF2e rules lookup is only available in the desktop runtime.";
    render();
    return;
  }

  ui.rulesBusy = true;
  ui.rulesMessage = force
    ? `Refreshing official PF2e rules for "${cleanQuery}"...`
    : `Searching official PF2e rules for "${cleanQuery}"...`;
  render();
  try {
    const result = await desktopApi.searchAonRules({
      query: cleanQuery,
      limit: ui.rulesSearchLimit,
      force: force === true,
    });
    ui.rulesResults = sortRulesResults(cleanQuery, Array.isArray(result?.results) ? result.results : []);
    ui.rulesSelectedUrl = str(ui.rulesResults[0]?.url || "");
    ui.rulesIndexedAt = str(result?.indexedAt || "");
    ui.rulesMessage = ui.rulesResults.length
      ? `Loaded ${ui.rulesResults.length} official PF2e rule match(es) for "${cleanQuery}".`
      : `No official PF2e rule matches found for "${cleanQuery}".`;
  } catch (err) {
    ui.rulesResults = [];
    ui.rulesSelectedUrl = "";
    ui.rulesMessage = `Official rules lookup failed: ${readableError(err)}`;
  } finally {
    ui.rulesBusy = false;
    render();
  }
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

async function chooseHexMapBackground() {
  if (!desktopApi?.pickMapBackground) {
    ui.hexMapMessage = "Map background picker is only available in the desktop build.";
    render();
    return;
  }
  try {
    const picked = await desktopApi.pickMapBackground();
    if (!picked?.fileUrl) return;
    const imageMeta = await loadImageMetadata(picked.fileUrl);
    const hexMap = getHexMapState();
    hexMap.backgroundPath = str(picked.path);
    hexMap.backgroundUrl = str(picked.fileUrl);
    hexMap.backgroundName = str(picked.name) || str(picked.path).split(/[\\/]/).pop() || "Map background";
    hexMap.backgroundNaturalWidth = Math.max(0, Number.parseFloat(String(imageMeta?.width || "0")) || 0);
    hexMap.backgroundNaturalHeight = Math.max(0, Number.parseFloat(String(imageMeta?.height || "0")) || 0);
    hexMap.backgroundScale = 1;
    hexMap.backgroundOffsetX = 0;
    hexMap.backgroundOffsetY = 0;
    saveState();
    ui.hexMapMessage = `Loaded map background: ${hexMap.backgroundName}. Adjust scale and offsets until the overlay hexes line up with the printed map.`;
    render();
  } catch (err) {
    ui.hexMapMessage = `Background load failed: ${readableError(err)}`;
    render();
  }
}

function ensureObsidianSettings() {
  if (!state?.meta?.obsidian || typeof state.meta.obsidian !== "object" || Array.isArray(state.meta.obsidian)) {
    state.meta.obsidian = {
      vaultPath: "",
      baseFolder: "DM Helper",
      lastSyncAt: "",
      lastSyncSummary: "",
      looksLikeVault: false,
      useForAiContext: true,
      readWholeVault: true,
      aiContextNoteLimit: 6,
      aiContextCharLimit: 3600,
      aiWriteFolder: "AI Notes",
      lastAiNoteAt: "",
      lastAiNotePath: "",
    };
  }
  state.meta.obsidian.vaultPath = str(state.meta.obsidian.vaultPath);
  state.meta.obsidian.baseFolder = str(state.meta.obsidian.baseFolder) || "DM Helper";
  state.meta.obsidian.lastSyncAt = str(state.meta.obsidian.lastSyncAt);
  state.meta.obsidian.lastSyncSummary = str(state.meta.obsidian.lastSyncSummary);
  state.meta.obsidian.looksLikeVault = state.meta.obsidian.looksLikeVault === true;
  state.meta.obsidian.useForAiContext = state.meta.obsidian.useForAiContext !== false;
  state.meta.obsidian.readWholeVault = state.meta.obsidian.readWholeVault !== false;
  state.meta.obsidian.aiContextNoteLimit = Math.max(1, Math.min(12, Number.parseInt(String(state.meta.obsidian.aiContextNoteLimit || 6), 10) || 6));
  state.meta.obsidian.aiContextCharLimit = Math.max(800, Math.min(12000, Number.parseInt(String(state.meta.obsidian.aiContextCharLimit || 3600), 10) || 3600));
  state.meta.obsidian.aiWriteFolder = str(state.meta.obsidian.aiWriteFolder) || "AI Notes";
  state.meta.obsidian.lastAiNoteAt = str(state.meta.obsidian.lastAiNoteAt);
  state.meta.obsidian.lastAiNotePath = str(state.meta.obsidian.lastAiNotePath);
  return state.meta.obsidian;
}

async function buildCopilotRuntimeContext(tabId, userInput, options = {}) {
  const baseContext = options?.baseContext || buildGlobalCopilotContext(tabId);
  const route = options?.route || null;
  const aiConfig = ensureAiConfig();
  const taskContext = {
    taskType: str(route?.taskType || baseContext?.taskType),
    taskLabel: str(route?.taskLabel || baseContext?.taskLabel),
    taskSaveTarget: str(route?.saveTarget || baseContext?.taskSaveTarget),
    routeReason: str(route?.routeReason || baseContext?.routeReason),
    entityType: str(route?.entityType || baseContext?.entityType),
  };
  let runtimeContext = {
    ...baseContext,
    ...taskContext,
    aonRulesEnabled: false,
    aonRulesMatches: Array.isArray(baseContext?.aonRulesMatches) ? baseContext.aonRulesMatches : [],
  };

  if (options?.skipVault) {
    if (
      aiConfig.useAonRules &&
      taskContext.taskType === "rules_question" &&
      desktopApi?.searchAonRules
    ) {
      try {
        const aonResult = await desktopApi.searchAonRules({ query: str(userInput), limit: aiConfig.compactContext ? 3 : 4 });
        runtimeContext.aonRulesMatches = Array.isArray(aonResult?.results) ? aonResult.results : [];
        runtimeContext.aonRulesEnabled = runtimeContext.aonRulesMatches.length > 0;
      } catch {
        runtimeContext.aonRulesEnabled = false;
      }
    }
    const retrieval = buildUnifiedRetrievalContext({
      tabId,
      userInput,
      route,
      baseContext: runtimeContext,
      obsidianNotes: [],
    });
    return {
      ...runtimeContext,
      ...retrieval,
    };
  }
  const settings = ensureObsidianSettings();
  if (!settings.useForAiContext || !settings.vaultPath || !desktopApi?.getObsidianVaultContext) {
    if (
      aiConfig.useAonRules &&
      taskContext.taskType === "rules_question" &&
      desktopApi?.searchAonRules
    ) {
      try {
        const aonResult = await desktopApi.searchAonRules({ query: str(userInput), limit: aiConfig.compactContext ? 3 : 4 });
        runtimeContext.aonRulesMatches = Array.isArray(aonResult?.results) ? aonResult.results : [];
        runtimeContext.aonRulesEnabled = runtimeContext.aonRulesMatches.length > 0;
      } catch {
        runtimeContext.aonRulesEnabled = false;
      }
    }
    const retrieval = buildUnifiedRetrievalContext({
      tabId,
      userInput,
      route,
      baseContext: runtimeContext,
      obsidianNotes: [],
    });
    return {
      ...runtimeContext,
      obsidianContextEnabled: settings.useForAiContext && !!settings.vaultPath,
      ...retrieval,
    };
  }
  try {
    const result = await desktopApi.getObsidianVaultContext({
      vaultPath: settings.vaultPath,
      baseFolder: settings.baseFolder,
      readWholeVault: settings.readWholeVault,
      maxNotes: settings.aiContextNoteLimit,
      maxChars: settings.aiContextCharLimit,
      query: str(userInput),
      activeTab: tabId,
    });
    runtimeContext = {
      ...runtimeContext,
      obsidianContextEnabled: true,
      obsidianContext: str(result?.summary),
      obsidianVaultNotes: Array.isArray(result?.notes) ? result.notes.slice(0, settings.aiContextNoteLimit) : [],
      obsidianContextNoteCount: Number(result?.notes?.length || 0),
      obsidianContextSourceRoot: str(result?.sourceRoot),
      obsidianContextLooksLikeVault: result?.looksLikeVault === true,
    };
    if (
      aiConfig.useAonRules &&
      taskContext.taskType === "rules_question" &&
      desktopApi?.searchAonRules
    ) {
      try {
        const aonResult = await desktopApi.searchAonRules({ query: str(userInput), limit: aiConfig.compactContext ? 3 : 4 });
        runtimeContext.aonRulesMatches = Array.isArray(aonResult?.results) ? aonResult.results : [];
        runtimeContext.aonRulesEnabled = runtimeContext.aonRulesMatches.length > 0;
      } catch {
        runtimeContext.aonRulesEnabled = false;
      }
    }
    const retrieval = buildUnifiedRetrievalContext({
      tabId,
      userInput,
      route,
      baseContext: runtimeContext,
      obsidianNotes: runtimeContext.obsidianVaultNotes,
    });
    return {
      ...runtimeContext,
      ...retrieval,
    };
  } catch (err) {
    ui.obsidianMessage = `Vault AI read failed: ${readableError(err)}`;
    if (
      aiConfig.useAonRules &&
      taskContext.taskType === "rules_question" &&
      desktopApi?.searchAonRules
    ) {
      try {
        const aonResult = await desktopApi.searchAonRules({ query: str(userInput), limit: aiConfig.compactContext ? 3 : 4 });
        runtimeContext.aonRulesMatches = Array.isArray(aonResult?.results) ? aonResult.results : [];
        runtimeContext.aonRulesEnabled = runtimeContext.aonRulesMatches.length > 0;
      } catch {
        runtimeContext.aonRulesEnabled = false;
      }
    }
    const retrieval = buildUnifiedRetrievalContext({
      tabId,
      userInput,
      route,
      baseContext: runtimeContext,
      obsidianNotes: [],
    });
    return {
      ...runtimeContext,
      obsidianContextEnabled: false,
      ...retrieval,
    };
  }
}

function buildObsidianAiNoteTitle() {
  const firstLine = str(ui.copilotDraft.input).split(/\r?\n/).find((line) => str(line).trim()) || "";
  const basis = firstLine || `${getTabLabel(activeTab)} AI Output`;
  const compact = basis.replace(/\s+/g, " ").trim();
  const clipped = compact.length > 72 ? `${compact.slice(0, 72).trim()}...` : compact;
  return `${getTabLabel(activeTab)} - ${clipped || "AI Output"}`;
}

async function writeCurrentAiOutputToObsidian() {
  const text = str(ui.copilotDraft.output);
  if (!text) {
    ui.copilotMessage = "No Loremaster output to write to the vault.";
    render();
    return;
  }
  const settings = ensureObsidianSettings();
  if (!settings.vaultPath) {
    ui.copilotMessage = "Choose your Obsidian vault folder first.";
    render();
    return;
  }
  if (!desktopApi?.writeObsidianAiNote) {
    ui.copilotMessage = "Vault note writing is only available in the desktop build.";
    render();
    return;
  }
  ui.obsidianBusy = true;
  ui.copilotMessage = "Writing current Loremaster output to the Obsidian vault...";
  render();
  try {
    const result = await desktopApi.writeObsidianAiNote({
      vaultPath: settings.vaultPath,
      baseFolder: settings.baseFolder,
      noteFolder: settings.aiWriteFolder,
      title: buildObsidianAiNoteTitle(),
      prompt: str(ui.copilotDraft.input),
      content: text,
      sourceTab: getTabLabel(activeTab),
      model: ensureAiConfig().model,
      generatedAt: new Date().toISOString(),
    });
    settings.lastAiNoteAt = str(result?.writtenAt) || new Date().toISOString();
    settings.lastAiNotePath = str(result?.relativePath);
    saveState();
    ui.obsidianMessage = result?.summary || "AI output written to the vault.";
    ui.copilotMessage = result?.summary || "Wrote current AI output to the vault.";
  } catch (err) {
    const message = readableError(err);
    ui.obsidianMessage = `Vault note write failed: ${message}`;
    ui.copilotMessage = `Vault note write failed: ${message}`;
  } finally {
    ui.obsidianBusy = false;
    render();
  }
}

function buildObsidianSyncPayload() {
  const settings = ensureObsidianSettings();
  const profile = getActiveKingdomProfile();
  const kingdom = getKingdomState();
  const hexMap = getHexMapState();
  const party = getHexMapParty(hexMap);
  return {
    vaultPath: settings.vaultPath,
    baseFolder: settings.baseFolder,
    campaignName: str(state?.meta?.campaignName) || "DM Helper Campaign",
    generatedAt: new Date().toISOString(),
    sessions: (state.sessions || []).map((session) => ({
      title: str(session.title),
      date: str(session.date),
      arc: str(session.arc),
      kingdomTurn: str(session.kingdomTurn),
      summary: str(session.summary),
      nextPrep: str(session.nextPrep),
      updatedAt: str(session.updatedAt || session.createdAt),
    })),
    npcs: (state.npcs || []).map((npc) => ({
      name: str(npc.name),
      role: str(npc.role),
      agenda: str(npc.agenda),
      disposition: str(npc.disposition),
      folder: str(npc.folder),
      notes: str(npc.notes),
      updatedAt: str(npc.updatedAt || npc.createdAt),
    })),
    quests: (state.quests || []).map((quest) => ({
      title: str(quest.title),
      status: str(quest.status),
      objective: str(quest.objective),
      giver: str(quest.giver),
      folder: str(quest.folder),
      stakes: str(quest.stakes),
      updatedAt: str(quest.updatedAt || quest.createdAt),
    })),
    locations: (state.locations || []).map((location) => ({
      name: str(location.name),
      hex: str(location.hex),
      folder: str(location.folder),
      whatChanged: str(location.whatChanged),
      notes: str(location.notes),
      updatedAt: str(location.updatedAt || location.createdAt),
    })),
    kingdom: {
      profileLabel: str(profile?.label || profile?.shortLabel),
      name: str(kingdom.name),
      charter: str(kingdom.charter),
      government: str(kingdom.government),
      heartland: str(kingdom.heartland),
      currentTurnLabel: str(kingdom.currentTurnLabel),
      currentDate: str(kingdom.currentDate),
      level: Number(kingdom.level || 0),
      size: Number(kingdom.size || 0),
      controlDC: Number(kingdom.controlDC || 0),
      resourcePoints: Number(kingdom.resourcePoints || 0),
      culture: Number(kingdom.culture || 0),
      economy: Number(kingdom.economy || 0),
      loyalty: Number(kingdom.loyalty || 0),
      stability: Number(kingdom.stability || 0),
      unrest: Number(kingdom.unrest || 0),
      fame: Number(kingdom.fame || 0),
      infamy: Number(kingdom.infamy || 0),
      notes: str(kingdom.notes),
      leaders: (kingdom.leaders || []).map((leader) => ({
        role: str(leader.role),
        name: str(leader.name),
        ability: str(leader.ability),
        leadershipBonus: Number(leader.leadershipBonus || 0),
        notes: str(leader.notes),
      })),
      settlements: (kingdom.settlements || []).map((settlement) => ({
        name: str(settlement.name),
        size: str(settlement.size),
        influence: Number(settlement.influence || 0),
        civicStructure: str(settlement.civicStructure),
        resourceDice: Number(settlement.resourceDice || 0),
        consumption: Number(settlement.consumption || 0),
        notes: str(settlement.notes),
      })),
      regions: (kingdom.regions || []).map((region) => ({
        hex: str(region.hex),
        status: str(region.status),
        terrain: str(region.terrain),
        workSite: str(region.workSite),
        notes: str(region.notes),
      })),
      turns: [...(kingdom.turns || [])].slice(-12).map((turn) => ({
        title: str(turn.title),
        date: str(turn.date),
        summary: str(turn.summary),
        resourceDelta: Number(turn.resourceDelta || 0),
        unrestDelta: Number(turn.unrestDelta || 0),
        notes: str(turn.notes),
      })),
    },
    hexMap: {
      mapName: str(hexMap.mapName),
      backgroundName: str(hexMap.backgroundName),
      party,
      forces: (hexMap.forces || []).map((force) => ({
        type: str(force.type),
        name: str(force.name),
        hex: str(force.hex),
        notes: str(force.notes),
      })),
      markers: (hexMap.markers || []).map((marker) => ({
        type: str(marker.type),
        title: str(marker.title),
        hex: str(marker.hex),
        notes: str(marker.notes),
      })),
    },
    liveCapture: [...(state.liveCapture || [])].slice(-30).map((entry) => ({
      sessionId: str(entry.sessionId),
      kind: str(entry.kind),
      note: str(entry.note),
      timestamp: str(entry.timestamp),
    })),
  };
}

async function chooseObsidianVault() {
  if (!desktopApi?.pickObsidianVault) {
    ui.obsidianMessage = "Vault picker is only available in the desktop build.";
    render();
    return;
  }
  try {
    const selected = await desktopApi.pickObsidianVault();
    if (!selected?.path) return;
    const settings = ensureObsidianSettings();
    settings.vaultPath = str(selected.path);
    settings.looksLikeVault = selected.looksLikeVault === true;
    saveState();
    ui.obsidianMessage = settings.looksLikeVault
      ? `Vault selected: ${settings.vaultPath}`
      : `Folder selected: ${settings.vaultPath}. It does not look like an Obsidian vault yet, but sync can still write markdown there.`;
    render();
  } catch (err) {
    ui.obsidianMessage = `Failed to choose vault: ${readableError(err)}`;
    render();
  }
}

async function syncObsidianVault() {
  if (!desktopApi?.syncObsidianVault) {
    ui.obsidianMessage = "Obsidian sync is only available in the desktop build.";
    render();
    return;
  }
  const settings = ensureObsidianSettings();
  if (!settings.vaultPath) {
    ui.obsidianMessage = "Choose your Obsidian vault folder first.";
    render();
    return;
  }
  ui.obsidianBusy = true;
  ui.obsidianMessage = "Syncing markdown notes into the Obsidian vault...";
  render();
  try {
    const result = await desktopApi.syncObsidianVault(buildObsidianSyncPayload());
    settings.looksLikeVault = result?.looksLikeVault === true;
    settings.lastSyncAt = str(result?.syncedAt) || new Date().toISOString();
    settings.lastSyncSummary = str(result?.summary || "");
    saveState();
    ui.obsidianMessage = result?.summary
      ? `${result.summary} Notes written to ${result.rootFolder}.`
      : `Synced notes to ${result?.rootFolder || settings.vaultPath}.`;
  } catch (err) {
    ui.obsidianMessage = `Vault sync failed: ${readableError(err)}`;
  } finally {
    ui.obsidianBusy = false;
    render();
  }
}

function loadImageMetadata(fileUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      resolve({
        width: image.naturalWidth || 0,
        height: image.naturalHeight || 0,
      });
    };
    image.onerror = () => reject(new Error("Could not read the selected image."));
    image.src = fileUrl;
  });
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
  if (state?.meta) {
    state.meta.aiMemory = buildAiMemoryDigests(state);
  }
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
          useAonRules: out.meta.aiConfig.useAonRules === false ? false : true,
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
          useAonRules: true,
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
  out.meta.obsidian =
    out.meta.obsidian && typeof out.meta.obsidian === "object" && !Array.isArray(out.meta.obsidian)
      ? {
          vaultPath: str(out.meta.obsidian.vaultPath),
          baseFolder: str(out.meta.obsidian.baseFolder) || "DM Helper",
          lastSyncAt: str(out.meta.obsidian.lastSyncAt),
          lastSyncSummary: str(out.meta.obsidian.lastSyncSummary),
          looksLikeVault: out.meta.obsidian.looksLikeVault === true,
          useForAiContext: out.meta.obsidian.useForAiContext !== false,
          readWholeVault: out.meta.obsidian.readWholeVault !== false,
          aiContextNoteLimit: Math.max(1, Math.min(12, Number.parseInt(String(out.meta.obsidian.aiContextNoteLimit || 6), 10) || 6)),
          aiContextCharLimit: Math.max(800, Math.min(12000, Number.parseInt(String(out.meta.obsidian.aiContextCharLimit || 3600), 10) || 3600)),
          aiWriteFolder: str(out.meta.obsidian.aiWriteFolder) || "AI Notes",
          lastAiNoteAt: str(out.meta.obsidian.lastAiNoteAt),
          lastAiNotePath: str(out.meta.obsidian.lastAiNotePath),
        }
      : {
          vaultPath: "",
          baseFolder: "DM Helper",
          lastSyncAt: "",
          lastSyncSummary: "",
          looksLikeVault: false,
          useForAiContext: true,
          readWholeVault: true,
          aiContextNoteLimit: 6,
          aiContextCharLimit: 3600,
          aiWriteFolder: "AI Notes",
          lastAiNoteAt: "",
          lastAiNotePath: "",
        };
  out.meta.aiMemory = normalizeAiMemoryState(out.meta.aiMemory);
  out.rulesStore = normalizeRulesStore(out.rulesStore);
  out.kingdom = normalizeKingdomState(out.kingdom);
  out.sessions = Array.isArray(out.sessions) ? out.sessions : [];
  out.npcs = Array.isArray(out.npcs) ? out.npcs : [];
  out.quests = Array.isArray(out.quests) ? out.quests : [];
  out.locations = Array.isArray(out.locations) ? out.locations : [];
  out.npcs = out.npcs.map((item) => ({ ...item, folder: normalizeWorldFolderName(item?.folder) }));
  out.quests = out.quests.map((item) => ({ ...item, folder: normalizeWorldFolderName(item?.folder) }));
  out.locations = out.locations.map((item) => ({ ...item, folder: normalizeWorldFolderName(item?.folder) }));
  out.liveCapture = Array.isArray(out.liveCapture) ? out.liveCapture : [];
  out.hexMap = normalizeHexMapState(out.hexMap);
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
        useAonRules: true,
        aiProfile: "fast",
      },
      aiHistory: [],
      worldFolders: {
        npcs: ["Capital", "Frontier"],
        quests: ["Main Arc", "Waystation"],
        locations: ["Capital", "Frontier"],
      },
      obsidian: {
        vaultPath: "",
        baseFolder: "DM Helper",
        lastSyncAt: "",
        lastSyncSummary: "",
        looksLikeVault: false,
        useForAiContext: true,
        readWholeVault: true,
        aiContextNoteLimit: 6,
        aiContextCharLimit: 3600,
        aiWriteFolder: "AI Notes",
        lastAiNoteAt: "",
        lastAiNotePath: "",
      },
      aiMemory: {
        campaignSummary: "",
        recentSessionSummary: "",
        activeQuestsSummary: "",
        activeEntitiesSummary: "",
        canonSummary: "",
        rulingsDigest: "",
        manualRulings: "",
        updatedAt: "",
        sourceCounts: {
          sessions: 0,
          openQuests: 0,
          npcs: 0,
          locations: 0,
          ruleEntries: 0,
          canonEntries: 0,
        },
      },
    },
    rulesStore: [],
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
    hexMap: createStarterHexMapState(),
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

function formatTimestamp(value) {
  const ms = safeDate(value);
  if (!ms) return "";
  try {
    return new Date(ms).toLocaleString();
  } catch {
    return str(value);
  }
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
