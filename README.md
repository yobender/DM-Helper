# DM Helper (Desktop)

Standalone desktop app for running tabletop RPG sessions.

## What this app does for you as GM

1. Keeps your game-state in one place:
   - Sessions
   - NPCs
   - Quests
   - Locations/hexes
2. Prevents notebook chaos:
   - Every session has a summary + explicit next-session prep notes
   - Quests stay visible by status (open, blocked, completed, failed)
   - Smart session checklist is generated from your current campaign state
   - Prep Queue gives you 30/60/90 minute time-boxed prep tasks
   - Smart wrap-up generates next-session prep bullets automatically
   - Session packet export creates a printable markdown prep brief
3. Uses your local PDFs during prep:
   - Indexes PDF text from your folder
   - Keyword search returns snippets + file path
4. Pushes clean data to Foundry:
   - NPCs export as `Actor` JSON
   - Quests and Locations export as `JournalEntry` JSON
5. Stays local-first:
   - Campaign data saved in local storage on your machine
   - Full JSON import/export backup

## PDF folder integration

Default PDF folder is set to:

`C:\Users\Chris Bender\OneDrive\Desktop`

In the app:

1. Open `PDF Intel` tab.
2. Confirm folder path (or choose another).
3. Click `Index PDFs`.
4. Search terms like `travel hazards`, `factions`, `downtime`, `undead`.

## New smart flow

1. In `Session Runner`, use **Session Start Smart Checklist** before game.
2. After session, open **Session Close Wizard (3-Step)** and answer:
   - biggest moments
   - cliffhanger pressure
   - player intent
3. Wizard output:
   - Smart Wrap-Up bullets
   - 3 suggested scene openers
4. The app injects both into `Next Prep` using replaceable markers.

## Added DM helpers

1. **Session Packet Export**
   - Export from latest session or any session card.
   - Includes recap, scene openers, smart wrap-up priorities, prep queue, quest/NPC/location focus, Foundry handoff, and PDF checks.
2. **Prep Queue Modes**
   - Switch between 30m, 60m, and 90m prep windows.
   - Queue adapts to open quests, current threads, and your indexed PDFs.
3. **Live Capture HUD**
   - Fast timestamped notes with quick tags: NPC, Hook, Rule, Loot, Retcon, Scene, Combat.
   - Append captured notes directly into session summary as a structured log.
4. **Writing + Spellcheck**
   - Right-click spellcheck suggestions in text fields, plus add-to-dictionary.
   - `Writing Helper` tab cleans rough draft text and converts it into session recap, prep bullets, NPC/quest/location notes.
5. **Local AI Writing (Ollama) + Auto-Connect**
   - Connect to local Ollama endpoint/model from `Writing Helper`.
   - Generate polished GM notes using your campaign context (latest session, open quests, NPCs, locations).
   - Auto-connect generated output to latest session prep with linked NPC/quest/location references.
   - Optional: include indexed PDF excerpts in AI prompts (`Use indexed PDF context in AI responses`).
6. **Global Loremaster (All Tabs)**
   - Every tab now includes a tab-aware `Loremaster` floating overlay.
   - Copilot uses active-tab context plus campaign data, then supports tab-specific apply actions:
     - Sessions/Dashboard/Foundry: attach into latest session prep/summary
     - Capture: create live capture entry
     - NPCs/Quests/Locations: create new entities from AI output
     - PDF Intel: extract a query and run search
     - Writing: send output into Writing Helper
   - Auto-run on tab switch is enabled by default and can be toggled in Copilot settings.
   - Model switching: pick from installed local models via dropdown, or type a custom model tag manually.

## Run desktop app

1. Open terminal in this folder:
   - `cd <your-project-folder>`
2. Install dependencies:
   - `npm install`
3. Start desktop app:
   - `npm run start`

## Build portable Windows app (.exe)

1. From this folder:
   - `npm run dist`
2. Output will be in:
   - `dist\DM Helper*.exe`

## Move To Another PC

Detailed checklist:

- See `SETUP-OTHER-PC.md`

Helpful scripts:

- `scripts\setup-other-pc.ps1`
- `scripts\create-desktop-shortcut.ps1`

Project handoff / development recap:

- `PROJECT-HANDOFF.md`
- `AI_ASSISTANT_ROADMAP.md`

Best practical path:

1. Use the portable build for the app itself.
   - Run `npm run dist`
   - Copy the file from `dist\` to the other PC
2. Export your campaign data from inside the app.
   - Use `Export Campaign JSON`
   - On the other PC, use `Import Campaign JSON`
3. Install Ollama on the other PC and pull the same models you use here.
   - Example:
     - `ollama pull lorebound-pf2e:latest`
     - `ollama pull lorebound-pf2e-fast:latest`
     - `ollama pull lorebound-pf2e-ultra-fast:latest`
4. Copy your PDFs to the other PC.
   - Best if they live in the same folder structure
   - If not, just re-index them once in `PDF Intel`

Optional if you want the exact same indexed-summary cache:

- Copy this file:
  - `C:\Users\Chris Bender\AppData\Roaming\dm-helper\pdf-index-cache.v1.json`

Recommended for GitHub:

1. Put this app in its own repo, not inside another unrelated project folder.
2. Push the source code only.
3. Do not push:
   - `node_modules`
   - `dist`
   - machine-specific app data
4. Use GitHub for source sync, and use the portable `.exe` or `npm install` on the other PC.

If you want the full app state instead of just imported campaign JSON:

- The app also stores renderer state under Electron app data on this machine.
- The clean approach is still:
  - export/import campaign JSON
  - re-index PDFs
  - pull the same Ollama models

## PF2e models on another PC

Yes. The custom Pathfinder models are easy to recreate on another PC as long as Ollama is installed.

From this repo:

1. Pull the base models:
   - `ollama pull gpt-oss:20b`
   - `ollama pull qwen2.5-coder:1.5b-base`
   - `ollama pull qwen2.5:3b`
2. Build the custom models:
   - `powershell -ExecutionPolicy Bypass -File .\scripts\setup-ollama-models.ps1`

That will build:

- `gpt-oss-20b-optimized:latest`
- `lorebound-pf2e:latest`
- `lorebound-pf2e-fast:latest`
- `lorebound-pf2e-ultra-fast:latest`
- `lorebound-pf2e-pure:latest`

If the other PC is weaker, use:

- `lorebound-pf2e-fast:latest` or
- `lorebound-pf2e-ultra-fast:latest`

## Local AI quick setup (optional)

1. Install and run Ollama locally.
2. Pull a model, for example:
   - `ollama pull llama3.1:8b`
3. In app `Writing Helper`:
   - keep endpoint `http://127.0.0.1:11434`
   - set model `llama3.1:8b`
   - click `Test Local AI`
   - use `Generate With Local AI`

## File map

- `index.html` - app shell
- `styles.css` - UI theme/layout
- `app.js` - frontend logic and state
- `main.js` - Electron main process + PDF indexing/search
- `preload.js` - safe desktop API bridge
- `package.json` - run/build scripts
