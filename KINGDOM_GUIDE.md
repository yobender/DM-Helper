# Kingdom Guide

This guide explains the built-in PF2e Kingmaker kingdom support in DM Helper.

## Rules profile

DM Helper uses a versioned kingdom rules profile so the app can stay aligned with the exact rules you want to run.

The current default profile is:

- `Vance & Kerenshara Remastered Kingdom Rules`
- Version date in app: `2026-03-16`

This profile stacks:

1. `Vance and Kerenshara's Comprehensive Pathfinder 2e Kingdom Building Rule Changes`
2. `Vance & Kerenshara's Remastered Kingdom Building Rules (Changes Only)`

The remastered document assumes the original patch is already in place, so the app treats them as one combined rules layer.

## What the Kingdom tab stores

The Kingdom tab keeps a structured kingdom sheet inside the campaign save:

- kingdom identity: name, charter, government, heartland, capital
- turn state: current turn label, current date, level, size, Control DC, resource die, RP, XP
- core stats: Culture, Economy, Loyalty, Stability
- pressure tracks: Renown, Fame, Infamy, Unrest, Ruin
- commodities: Food, Lumber, Luxuries, Ore, Stone
- leaders: role, name, PC/NPC status, leadership bonus, relevant skills, specialized kingdom skills, notes
- settlements: size, influence, civic structure, resource dice, local consumption, notes
- regions: hex, status, terrain, work site, notes
- turn history: recorded turn summary, deltas, risks, pending projects

## What the AI knows

Loremaster is now told that it is operating inside DM Helper and that the Kingdom tab is a structured subsystem, not just freeform notes.

When you use Loremaster from the Kingdom tab, it receives:

- the active kingdom rules profile summary
- current turn structure
- the current kingdom sheet
- leaders, settlements, regions, recent turns, and pending projects

Loremaster should help with:

- action order
- leader assignments
- settlement planning
- turn risk review
- what to record after a turn

Loremaster should not silently change math. The sheet stays deterministic and the AI stays advisory.

## Deterministic vs AI

Use the sheet for truth.

Deterministic data:

- current stats
- commodities
- unrest / ruin / renown / fame / infamy
- leader roster
- settlement records
- region records
- saved turn history

AI assistance:

- explain what the active rules imply
- suggest the safest or strongest turn order
- assign likely best leaders to actions
- turn rough plans into table-ready checklists
- summarize what should be saved back into the sheet

## Suggested workflow

### 1. Set up the kingdom sheet

Fill in:

- name
- charter
- government
- heartland
- capital
- level
- size
- Control DC
- resource die
- RP / XP
- core stats
- commodities
- pressure tracks

### 2. Add leaders

Track each invested role and note:

- who fills it
- whether they are a PC or NPC
- Leadership Bonus
- relevant skills
- specialized kingdom skills

### 3. Add settlements and regions

At minimum, keep one settlement and your claimed hexes current.

This matters because the remastered rules care about:

- civic structures
- local influence
- local resource dice
- work sites
- settlement-specific action limits

### 4. Ask Loremaster before the turn

Useful prompts:

- `Using our V&K kingdom rules, plan my next kingdom turn.`
- `Tell me which leader should handle each action this turn and why.`
- `Given our unrest, ruin, and commodities, what is the safest kingdom turn this month?`
- `Turn this settlement expansion plan into a step-by-step kingdom turn.`

### 5. Run the turn and record deltas

Use the turn form to record:

- RP delta
- commodity deltas
- unrest / renown / fame / infamy changes
- ruin changes
- summary
- risks
- pending projects

This updates the kingdom sheet and stores a turn log.

## Turn structure in the current profile

The active profile expects turns in this order:

1. Upkeep
2. Activities
3. Construction
4. Event

Key profile assumptions:

- each PC leader gets 3 actions
- each NPC leader gets 2 actions
- Consumption cannot go below 0
- Ruin threshold is 5
- settlements matter more because civic structures grant local actions and local build capacity

## Why the rules profile is versioned

The app stores the active profile id with the kingdom so later rule updates do not silently reinterpret old turn logs.

This makes it possible to support later variants, such as:

- updated V&K releases
- house-ruled profiles
- closer-to-RAW profiles

without breaking existing campaign history.

## Current limitations

This first version is a kingdom tracker and turn planner, not a full map simulator.

Included now:

- structured kingdom sheet
- leaders / settlements / regions / turns
- app-aware AI support
- built-in guide based on the active rules profile

Not included yet:

- full army combat engine
- settlement grid builder
- visual hex map editor
- automatic rules adjudication for every kingdom action

## Best practice

Treat the Kingdom tab as the source of truth and treat Loremaster as the analyst sitting beside it.

Ask AI for:

- options
- sequencing
- risk analysis
- record-ready summaries

Keep the actual numbers and state changes in the sheet.
