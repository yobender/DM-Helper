# DM Helper Project Handoff

This file is the development handoff for the current `DM Helper` desktop app.

Use it on another PC to understand:

- what the app already does
- how the AI stack currently works
- what was intentionally deferred
- what to build next

## Repo

- GitHub: `https://github.com/yobender/DM-Helper`
- Local app folder on this PC: `d:\LoreBound\kingmaker-gm-studio`

## Product Direction

`DM Helper` is the main hub.

The AI is not meant to be a separate product. The built-in AI subsystem is `Loremaster`, and the intended architecture is:

- local for storage
- local for retrieval
- local for memory
- local for workflow
- model only for reasoning/synthesis

That design choice is deliberate. It keeps the app grounded in actual campaign data instead of letting the model invent continuity.

## Current App State

The app is an Electron desktop app with these major areas:

- `Dashboard`
- `Session Runner`
- `Live Capture HUD`
- `Writing Helper`
- `Kingdom`
- `Hex Map`
- `NPCs`
- `Quests`
- `Locations`
- `PF2e Rules`
- `PDF Intel`
- `Obsidian Vault`
- `Foundry Export`

## AI Architecture Already Implemented

### 1. Task Router

Loremaster now classifies requests into explicit task types instead of relying only on weak tab heuristics.

Examples:

- `rules_question`
- `campaign_lookup`
- `session_summary`
- `note_update`
- `kingdom_helper`
- `map_helper`
- `pdf_lookup`
- `vault_workflow`
- `writing_cleanup`

This is the start of the orchestrator layer from the design spec.

### 2. Memory Digests

The app builds compact local AI memory digests from campaign state.

Current digests:

- `campaignSummary`
- `recentSessionSummary`
- `activeQuestsSummary`
- `activeEntitiesSummary`
- `rulingsDigest`
- `canonSummary`

These are intentionally small. They are meant to be stable context, not giant raw dumps.

### 3. Unified Retrieval

Loremaster now retrieves from multiple local sources in one pass.

Current retrieval sources:

- sessions
- quests
- NPCs
- locations
- kingdom state
- hex map state
- live-captured rules/retcons
- saved PDF summaries
- Obsidian vault notes
- Archives of Nethys rule matches
- persistent rules/canon store

The app also shows a `Retrieved Context` preview so you can inspect what it fed to the model.

### 4. PF2e Rules Layer

There is now a `PF2e Rules` tab.

It supports:

- official Pathfinder 2e rules lookup through Archives of Nethys
- exact-title-biased rules search
- local rules/rulings digest comparison
- `Official vs Local` split view
- send-query-to-Loremaster flow

Important boundary:

- we are using targeted live lookup and local cache
- we are **not** mirroring the whole AoN site into the app

### 5. Local Rules / Canon Learning Layer

This is the newest AI-learning layer.

It is not model fine-tuning.

It is local memory improvement through accepted knowledge.

Current capabilities:

- save official rule snippets as `Official Note`
- save Loremaster output as `Accepted Ruling`
- save Loremaster output as `Canon Memory`
- create manual local entries directly in the rules tab
- retrieve those entries later in future AI answers

This is the correct version of “AI learning” for this app.

## What “Learning” Means In This App

The app does **not** retrain model weights on use.

Instead, it improves through:

- better local storage
- better local retrieval
- better memory digests
- accepted rulings
- accepted canon facts

That is aligned with the original hybrid-assistant plan.

## Obsidian Integration

Current Obsidian support:

- choose a vault folder
- sync DM Helper notes into markdown
- let Loremaster read compact vault context
- write current AI output back into the vault

This is currently strongest as:

- export/sync
- read context for AI
- write AI notes to markdown

It is **not** yet a full two-way editor.

## PDF Integration

Current PDF support:

- index local PDFs
- search snippets
- save per-file summaries
- use indexed PDF context in AI prompts
- use PDF summaries as persistent memory

This is useful for:

- campaign modules
- GM books you legally possess locally
- lore recap
- subsystem reminders

## Kingdom + Hex Map

Current kingdom state:

- tracked kingdom sheet
- creation planner
- charter/government/heartland reference
- derived modifiers and summaries

Current hex map state:

- dedicated `Hex Map` tab
- pan with mouse drag
- zoom with mouse wheel
- map background support
- party marker
- party trail
- allied/enemy/caravan force markers
- hex records and marker notes

This is still early, but it is already usable as a campaign map layer.

## Local Models

The app is currently built around local-first AI.

Important model direction:

- PF2e-focused local wrappers exist in `ollama-models`
- `lorebound-pf2e:latest` is the deeper PF2e 20b path
- smaller/faster models are available for lighter work

The app already supports model selection and local Ollama configuration.

## What Was Deliberately Deferred

These are not mistakes. They were intentionally postponed:

- cloud API provider wiring
- full rules mirror
- true two-way Obsidian editing
- full autonomous note rewrites
- agent swarm behavior
- Foundry automation
- giant graph-memory system
- model fine-tuning

## Why Cloud/API Was Deferred

The user asked for:

- local AI
- optional smarter online AI

That is a good direction, but it should come **after** the local memory layer.

Reason:

- if cloud comes too early, it becomes the source of truth
- that breaks the architecture we actually want

Correct order:

1. local memory and local canon
2. strong retrieval
3. provider abstraction
4. optional cloud reasoning

## Recommended Next Steps

This is the next practical sequence.

### 1. Provider Abstraction

Add AI provider mode:

- `Local`
- `Auto`
- `Cloud`

Design rule:

- local retrieval remains the source of truth
- cloud only receives compact context packages

### 2. Cloud/API Settings

Add optional settings for a smarter online model.

Support should be generic enough for an OpenAI-compatible endpoint.

At minimum:

- provider mode
- base URL
- API key
- model name
- timeout
- per-task routing policy

### 3. Per-Task Provider Routing

Good default approach:

- local for small prep and fast workflows
- cloud for heavy synthesis only

Examples of cloud-suitable tasks:

- long session consolidation
- difficult rules synthesis
- cross-source summary drafting
- large note cleanup

### 4. Reviewable Note Updates

The AI should draft updates, not silently overwrite state.

Add:

- preview diff / preview block
- `Apply`
- `Reject`
- `Save as Canon`
- `Save as Ruling`

### 5. Stronger PF2e Rules UX

Still needed:

- rule-type filters
- better exact-match handling
- easier save-to-store flows
- clearer official vs local precedence

### 6. Kingdom Events

After the AI/provider layer is stable:

- build a kingdom event system
- attach events to turns
- optionally tie events to map hexes

## Practical “What To Do On The Other PC”

### Code + Models

1. Clone or pull the repo:

   - `git clone https://github.com/yobender/DM-Helper.git`

2. Install dependencies:

   - `npm install`

3. Build PF2e model wrappers:

   - `powershell -ExecutionPolicy Bypass -File .\scripts\setup-ollama-models.ps1`

4. Run the app:

   - `npm run start`

5. Or build the portable exe:

   - `npm run dist`

### Campaign Data

To move actual campaign state:

1. Export campaign JSON from inside the app on the source PC
2. Import campaign JSON on the other PC
3. Reconnect:
   - PDF folder
   - Obsidian vault
   - Ollama models

### Optional Cache Copy

If you want the same PDF summary cache:

- copy the Electron app data cache for `dm-helper`

## Files To Read First On Another PC

Read these in order:

1. `PROJECT-HANDOFF.md`
2. `AI_ASSISTANT_ROADMAP.md`
3. `README.md`
4. `SETUP-OTHER-PC.md`

## Development Notes

- The repo may contain a committed `dist` portable build. Rebuild it when app logic changes materially.
- The current worktree before this handoff included ongoing app changes in:
  - `app.js`
  - `main.js`
  - `preload.js`
  - `styles.css`
  - `kingdom-rules-data.json`
  - `scripts/launch-dm-helper.cmd`
  - `dist/DM Helper 0.1.0.exe`
- The newest AI-learning work was primarily in:
  - `app.js`
  - `main.js`

## Short Summary

Current status:

- DM Helper is already a local-first hybrid campaign assistant
- task routing exists
- memory digests exist
- unified retrieval exists
- PF2e rules tab exists
- Obsidian context exists
- PDF memory exists
- local rules/canon learning now exists

The next real layer is:

- provider abstraction for `local + optional cloud`

That is the correct continuation point.
