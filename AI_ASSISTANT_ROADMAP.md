# DM Helper AI Assistant Roadmap

## Bottom Line

Yes, we should use most of this design.

No, we should not adopt it literally as written.

The current DM Helper app already matches a large part of the architecture:

- built-in assistant inside the app, not a separate product
- local-first campaign data
- local Obsidian vault integration
- local PDF indexing and retrieval
- task-specific AI modes
- a single orchestrator-style request flow

The main part that does **not** match is the "cloud side" assumption.

Right now DM Helper is built around a **local reasoning model** through Ollama. That is consistent with what you asked for earlier and is the right fit for this app today.

So the correct version of the philosophy for this project is:

**Local for storage, retrieval, memory, workflow, and currently also reasoning.**

Later, if we want, we can generalize that to:

**Local for storage, retrieval, memory, workflow.  
Reasoning model can be local first, with optional cloud later.**

## What Already Fits DM Helper Well

These parts should be kept and used as the backbone of the project.

### 1. The app stays the main hub

This is already correct.

DM Helper already contains:

- campaign tracking
- session logs
- NPC / quest / location records
- kingdom tools
- hex map tools
- PDF Intel
- Obsidian vault integration
- built-in Loremaster panel

The assistant should continue to live inside this app.

### 2. One orchestrator, not a swarm

This is also correct.

Current DM Helper behavior already resembles a single orchestrator:

- classify current tab / mode
- build context
- gather campaign state
- gather PDF snippets when needed
- gather vault note context when enabled
- call the model
- post-process output
- optionally save or apply output

That is the right version 1 architecture.

We should **not** add a multi-agent system now.

### 3. Context engineering matters more than model choice

This is strongly correct for this project.

In DM Helper, answer quality is mostly controlled by:

- what tab the user is on
- what mode is selected
- what local state is gathered
- what PDF / vault context is injected
- what output structure is requested

That is exactly the right design direction.

### 4. Task-specific pipelines

This also fits what we already have and should be expanded.

Current DM Helper already has partial pipelines for:

- general assistant chat
- session prep / recap
- NPC creation
- quest creation
- location creation
- PDF lookup
- writing cleanup

The writeup's four core job types are a good fit and should become more explicit.

### 5. Local source of truth

This is fully aligned.

DM Helper already stores or syncs:

- campaign state
- sessions
- NPCs
- quests
- locations
- kingdom data
- hex map state
- PDF summaries / retrieval chunks
- Obsidian notes

That should remain the source of truth.

## What Needs To Be Changed Before Using This Spec

These are the parts of the writeup that should be rewritten for DM Helper instead of copied directly.

### 1. Replace "cloud model" with "reasoning model"

The writeup assumes remote reasoning as the main plan.

That is not the best fit for the app right now.

DM Helper currently works best as:

- local data layer
- local retrieval layer
- local reasoning model through Ollama
- optional cloud later if we explicitly add it

So the architecture should say:

- local-first by default
- cloud optional, not required
- prompt/context builder should be model-provider agnostic

### 2. Do not assume full Pathfinder rules truth from the model

The writeup is right to worry about D&D bleed.

For DM Helper, the correct stance is:

- retrieve PF2e-grounded local context first
- prefer indexed PDF and vault notes
- keep explicit PF2e instructions in the model prompt
- mark uncertain answers clearly

That means the assistant should behave as a **grounded PF2e copilot**, not as a freeform fantasy chatbot.

### 3. The "retriever" is only partially done

The writeup assumes a mature hybrid retriever across all note types.

DM Helper is not there yet.

What already exists:

- PDF keyword search
- PDF semantic / hybrid retrieval
- vault note relevance scoring
- campaign state packaging

What is still missing:

- a true cross-source retriever for sessions + NPCs + quests + locations + kingdom + vault + PDFs in one ranking pass
- metadata filters by note type / campaign / recency / source
- exact-match bias for rules terms across all stores

### 4. Memory layer exists, but not as named digests

The writeup suggests dedicated memory files like:

- campaign summary
- recent session summary
- active quests
- active entities
- rulings / house rules

DM Helper currently has memory-like context, but it is assembled live from app state instead of stored as dedicated digest notes.

That is workable, but a formal memory layer would improve consistency.

## What DM Helper Already Has That Matches The Writeup

### Current local data layer

Already present:

- app-managed campaign state
- Obsidian vault sync and vault AI context
- PDF indexing and summaries
- kingdom records
- hex map records

### Current orchestrator behavior

Already present:

- current tab influences mode
- prompt recipes differ by task
- app state is gathered before calling the model
- output is post-processed
- output can be applied back into app records

### Current note-update behavior

Already present in partial form:

- AI can create NPC / quest / location entries
- AI can attach output to sessions
- AI can write to Obsidian vault

Still missing:

- safe structured diff / patch review before save
- duplicate entity detection strong enough to merge instead of clone

### Current retrieval behavior

Already present in partial form:

- PDF snippets can be retrieved into model context
- Obsidian vault notes can be compacted into model context

Still missing:

- shared retrieval layer across every note source
- rules-specific exact search pipeline

## What We Should Build Next

This is the part of the writeup that should become the actual roadmap.

### Phase 1: Formalize the current architecture

Do this now.

#### A. Introduce explicit request classes

Add a real router that classifies requests into:

- rules_question
- campaign_lookup
- session_summary
- note_update
- kingdom_helper
- map_helper

This should become a first-class part of Loremaster instead of being inferred indirectly from tabs and prompts.

#### B. Add context recipes by request class

For each request class, define:

- required system instructions
- local sources to search
- memory buckets allowed
- output format
- save behavior

This is the single biggest quality upgrade we can make.

#### C. Add a proper memory layer

Create local memory summaries such as:

- `campaign_summary`
- `recent_session_summary`
- `active_quests`
- `active_entities`
- `rulings_digest`

These can be stored in app state, Obsidian, or both.

The point is to stop rebuilding everything from scratch every time.

### Phase 2: Build a true cross-source retriever

Do this after the router/context recipes.

Target behavior:

- one retrieval pass across:
  - sessions
  - NPCs
  - quests
  - locations
  - kingdom records
  - hex map records
  - Obsidian notes
  - PDF chunks
- exact term boosting for:
  - PF2e rule names
  - feat names
  - spell names
  - conditions
  - entity names
- metadata filters:
  - note type
  - date / recency
  - campaign area
  - source bucket

This is the missing middle layer between "raw app state" and "good grounded answers."

### Phase 3: Add rules-grounding as a first-class feature

This is especially important because the app is PF2e-focused.

Needed:

- official PF2e rules summaries separated from house rules
- exact-match bias for rules terms
- response format that clearly labels:
  - official rule
  - house rule
  - inference / uncertainty

Right now DM Helper can answer PF2e-style questions, but the pipeline is not yet strict enough for reliable rules lookup.

### Phase 4: Upgrade note updates to reviewable structured saves

Do not let AI silently rewrite canonical notes.

Instead:

- show a draft
- show target note
- highlight additions / replacements
- let the user approve the save

This is the right way to keep campaign memory clean.

## What Should Wait

These ideas are fine long term, but they should not be built now.

### 1. Multi-agent architecture

Do not do this now.

The current app does not need it.

### 2. Full fine-tuning on Pathfinder corpus

Do not do this now.

Retrieval + strong PF2e prompts + structured context will get far more value for less work.

### 3. Always-listening / voice / live speech

Not a version 1 feature.

### 4. Fully autonomous note rewriting

Do not let the assistant silently mutate your campaign canon.

Review first.

### 5. Heavy graph-memory system

Not needed yet.

A good digest layer plus retriever is enough for now.

## Recommended Revised Philosophy For DM Helper

Use this instead of the original wording:

### DM Helper AI architecture

- DM Helper is the main campaign hub.
- Loremaster is an internal subsystem, not a separate app.
- Local data is the source of truth.
- Retrieval and memory selection matter more than model size.
- The reasoning model should be provider-agnostic:
  - local first
  - cloud optional later
- Every request should use a task-specific context recipe.
- AI outputs that change campaign canon should be reviewable before saving.

## Recommended Version 1 Build Target

If we turn the writeup into an actual implementation target for the current app, version 1 should be:

### Build now

- explicit task router
- formal context recipes
- memory digests
- cross-source retriever
- stricter PF2e rules pipeline
- reviewable note update workflow

### Keep as-is for now

- built-in assistant inside DM Helper
- local Ollama reasoning
- PDF Intel
- Obsidian vault integration
- session / NPC / quest / location workflows

### Do later

- cloud provider option
- more advanced rules corpora
- agent splitting
- voice / live table features

## Final Recommendation

Use this writeup as a **roadmap and architecture guide**, not as an exact implementation spec.

The strongest parts to keep are:

- built-in assistant inside the app
- local source of truth
- orchestrator model
- context engineering
- task-specific pipelines
- hybrid retrieval
- reviewable saves

The main change we should make is:

**replace "cloud reasoning engine" with "reasoning engine abstraction, local first."**

That keeps the design consistent with DM Helper as it exists today and avoids rebuilding the system around an assumption that is not actually true in the app.
