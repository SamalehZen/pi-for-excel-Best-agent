# Task: Multi-Agent Architecture (Option A — Tool-based)

**Priority:** P0
**Effort:** ~3-4 weeks
**Branch:** `feat/multi-agent-orchestrator`

---

## Context

Pi is currently a single monolithic agent with 1 system prompt, 20+ tools, and 11 templates. When the user asks for complex tasks (e.g. "Build me a DCF model with recent Tesla data, formatted properly"), the single agent must juggle reading, writing, formatting, researching, and modeling in one long conversation — it loses context, forgets steps, and produces inconsistent results.

The template system (`apply_template` with `design_only` mode) fails on complex data layouts because the auto-detection of headers/sections/totals is too naive. Simple flat data works, but multi-section spreadsheets with sub-totals, merged headers, and nested structures get misformatted.

## Goal

Implement an **Orchestrator + Sub-Agents** system using the **Tool-based approach** (Option A). The main agent becomes an orchestrator that delegates tasks to specialized sub-agents. Each sub-agent runs as an inner LLM loop invoked through a `delegate_task` tool, with its own focused system prompt and restricted tool set.

## Key Decision: Template Builder always runs

**Decision:** When a user selects a template from the dropdown gallery, the sub-agent Template Builder is **always** invoked — regardless of data complexity (simple or complex). There is no detection heuristic or fallback to the naive `apply_template` `design_only` auto-detect.

**Rationale:**
- Consistent quality on every apply — no "sometimes good, sometimes broken"
- No complexity-detection logic to build and maintain (one less thing to break)
- The user already chose the template from the dropdown — the sub-agent only needs to analyze the data layout and apply the template's palette/typography to the correct zones
- Acceptable tradeoff: ~3-8 seconds extra (LLM analysis) in exchange for reliable results every time

**Flow:**
```
User clicks template in dropdown gallery
  → System extracts palette + typography from chosen template
  → delegate_task({ role: "template-builder", task: "...", context: { palette, typography } })
  → Template Builder reads data, detects zones, applies design
  → Result displayed
```

The existing `apply_template` tool with `action: "list"`, `action: "preview"`, and `mode: "full"` (blank sheet) remains unchanged. Only `mode: "design_only"` (apply to existing data) is replaced by the sub-agent path.

---

## Architecture Overview

```
User prompt
    │
    ▼
┌─────────────────────────────────────────┐
│  Orchestrator (Pi main agent)           │
│  - Analyzes intent                      │
│  - Plans delegation sequence            │
│  - Calls delegate_task tool             │
│  - Synthesizes results for user         │
└────────┬────────┬────────┬──────────────┘
         │        │        │
    ┌────▼──┐ ┌──▼────┐ ┌▼───────┐
    │Analyst│ │Builder│ │Stylist │  ... more roles
    └───────┘ └───────┘ └────────┘
    Each sub-agent:
    - Own system prompt (short, focused)
    - Restricted tool subset
    - Runs as inner LLM loop
    - Returns structured result
```

### Why Tool-based (Option A)

- Fits the existing `pi-agent-core` tool loop — no framework changes
- `WorkbookCoordinator` serializes mutations automatically
- `workbook_history` checkpoints work unchanged
- Experimental tool gates still apply
- Progressive rollout: add roles incrementally

---

## Sub-Agent Roles

### 1. Analyst (read-only)
- **Role:** Read, understand, summarize data. Never modifies the workbook.
- **Tools:** `get_workbook_overview`, `read_range`, `search_workbook`, `trace_dependencies`, `explain_formula`
- **System prompt focus:** Data comprehension, pattern detection, anomaly identification, summary generation.
- **Use cases:** "What's in this spreadsheet?", "Find all formulas with errors", "Summarize the revenue trend"

### 2. Builder (structure + content)
- **Role:** Create structures, write formulas, build from scratch.
- **Tools:** `get_workbook_overview`, `read_range`, `write_cells`, `fill_formula`, `modify_structure`, `execute_office_js`
- **System prompt focus:** Spreadsheet architecture, formula best practices, sheet organization, named ranges.
- **Use cases:** "Create a DCF model", "Add a summary sheet", "Build an amortization schedule"

### 3. Stylist (formatting)
- **Role:** Format, style, apply themes, conditional formatting.
- **Tools:** `get_workbook_overview`, `read_range`, `format_cells`, `conditional_format`, `view_settings`, `conventions`, `apply_template`
- **System prompt focus:** Visual design, Excel formatting best practices, color theory, data visualization via formatting.
- **Use cases:** "Make this look professional", "Apply financial formatting", "Add data bars to the revenue column"

### 4. Template Builder (structure analysis + intelligent template application)
- **Role:** Analyze existing data structure and apply templates intelligently to complex layouts.
- **Tools:** `get_workbook_overview`, `read_range`, `search_workbook`, `format_cells`, `conditional_format`, `view_settings`, `modify_structure`, `execute_office_js`
- **System prompt focus:** Layout detection (titles, headers, sub-headers, data zones, sub-totals, grand totals, spacer rows, merged cells), zone-by-zone formatting, adaptive template application.
- **Use cases:** "Apply a professional template to this P&L", "Format this complex report", "Make this data presentable"
- **Key differentiator:** Uses LLM intelligence to semantically understand the data layout before formatting, unlike the current naive auto-detect.

### 5. Researcher (external data)
- **Role:** Search the web, fetch data, call APIs.
- **Tools:** `web_search`, `fetch_page`, `mcp`, `python_run`, `files`
- **System prompt focus:** Data sourcing, fact verification, structured data extraction.
- **Use cases:** "Find Tesla's latest revenue figures", "Get the current exchange rates", "Research industry benchmarks"

### 6. Modeler (financial/quantitative)
- **Role:** Financial modeling, complex calculations, Python-based analysis.
- **Tools:** `get_workbook_overview`, `read_range`, `write_cells`, `fill_formula`, `python_run`, `python_transform_range`, `execute_office_js`
- **System prompt focus:** Financial modeling standards (loaded from `financial-modeling` skill), formula construction rules, assumption separation, cross-sheet linking.
- **Use cases:** "Build a 3-statement model", "Calculate WACC", "Run a sensitivity analysis"

### 7. Debugger (audit + repair)
- **Role:** Trace errors, audit formulas, fix broken references, repair data.
- **Tools:** `get_workbook_overview`, `read_range`, `search_workbook`, `trace_dependencies`, `explain_formula`, `write_cells`, `workbook_history`
- **System prompt focus:** Error diagnosis, formula auditing, circular reference detection, data validation.
- **Use cases:** "Why does cell F15 show #REF?", "Audit all formulas in this sheet", "Fix the broken references"

---

## Implementation Plan

### Phase 1: Core Infrastructure (~1 week)

#### Task 1.1: Sub-agent role definitions
- **File:** `src/agents/roles.ts`
- **Content:** Type definitions for sub-agent roles, each containing:
  - `id`: unique role identifier
  - `name`: display name
  - `description`: what this role does
  - `systemPrompt`: specialized prompt text
  - `allowedTools`: list of tool names this role can use
  - `requiredContext`: what auto-context to inject (blueprint, selection, changes)
  - `maxTurns`: maximum inner loop iterations (safety limit)
  - `skillsToPreload`: skills to auto-load (e.g. `financial-modeling` for Modeler)
- **Roles to define:** analyst, builder, stylist, template-builder, researcher, modeler, debugger
- **Export:** `SUB_AGENT_ROLES` registry, `SubAgentRole` type, `getRole(id)` lookup

#### Task 1.2: Sub-agent runner
- **File:** `src/agents/sub-agent-runner.ts`
- **Purpose:** Execute a sub-agent as an inner LLM loop
- **Interface:**
  ```typescript
  interface SubAgentRequest {
    roleId: string;
    task: string;                    // natural language task description
    context?: string;                // additional context from orchestrator
    parentModel?: string;            // inherit model from orchestrator
    maxTurns?: number;               // override role default
  }

  interface SubAgentResult {
    roleId: string;
    status: "completed" | "failed" | "max_turns_reached";
    summary: string;                 // human-readable summary of what was done
    toolCallCount: number;
    turnsUsed: number;
    errors?: string[];
  }
  ```
- **Behavior:**
  1. Look up role from registry
  2. Filter the full tool list to only allowed tools for this role
  3. Build specialized system prompt (role prompt + workbook context)
  4. Run inner LLM loop using the same `streamFn` / provider as the parent agent
  5. Collect tool results, track mutations
  6. Return structured `SubAgentResult`
- **Constraints:**
  - Must use the same `WorkbookCoordinator` as the parent (serialized mutations)
  - Must trigger `workbook_history` checkpoints for mutations
  - Must respect execution mode (Auto/Confirm)
  - Must respect tool output truncation
  - Inner loop capped at `maxTurns` (default: 15) to prevent runaway

#### Task 1.3: delegate_task tool
- **File:** `src/tools/delegate-task.ts`
- **Purpose:** Tool that the orchestrator calls to invoke a sub-agent
- **Schema:**
  ```typescript
  {
    role: string;           // "analyst" | "builder" | "stylist" | "template-builder" | "researcher" | "modeler" | "debugger"
    task: string;           // what to accomplish
    context?: string;       // additional info from orchestrator
    wait_for?: string[];    // role IDs of tasks that must complete first (for sequential chaining)
  }
  ```
- **Execution policy:** `mutate/content` (sub-agents may write to the workbook)
- **Result:** Returns `SubAgentResult` as structured tool result with details
- **Registration:**
  - Add to `src/tools/index.ts` in `createAllTools()`
  - Add to `src/tools/names.ts` if treated as core, or keep as auxiliary
  - Add to `src/ui/tool-renderers.ts` (render sub-agent results in sidebar)
  - Add to `src/ui/humanize-params.ts` (humanize role + task params)

#### Task 1.4: Orchestrator system prompt update
- **File:** `src/prompt/system-prompt.ts`
- **Changes:**
  - Add a new section `## Delegation` describing available sub-agent roles
  - Include guidance on when to delegate vs handle directly:
    - Simple single-tool tasks → handle directly (no delegation overhead)
    - Complex multi-step tasks → delegate to appropriate role(s)
    - Tasks spanning multiple domains → sequential delegation (research → build → format)
  - Include role descriptions so the orchestrator knows what each sub-agent can do
  - Update TOOLS section to include `delegate_task`

### Phase 2: Template Builder Sub-Agent (~1 week)

This is the highest-priority sub-agent because it solves the known pain point with `design_only` mode on complex data.

#### Task 2.1: Template Builder role prompt
- **File:** `src/agents/roles/template-builder.ts`
- **System prompt content:**
  - Identity: "You are the Template Builder, specialized in analyzing spreadsheet layouts and applying formatting."
  - Step-by-step workflow:
    1. READ the entire used range with `read_range` in detailed mode
    2. IDENTIFY zones: title rows, meta/info rows, column headers, sub-section headers, data rows, sub-total rows, grand total rows, spacer/empty rows, merged areas
    3. CLASSIFY each zone type based on content patterns:
       - Title: single cell spanning multiple columns, large text, often row 1-2
       - Headers: row with short text labels, often bold or colored
       - Sub-headers: row with category labels followed by data rows
       - Data: rows with mixed text/numbers following a header pattern
       - Sub-totals: rows with SUM/SUBTOTAL formulas or bold numeric values
       - Grand totals: bottom summary row, often with double borders
       - Spacer: empty rows separating sections
    4. PLAN formatting for each zone (specific format_cells calls)
    5. EXECUTE the plan zone-by-zone
    6. VERIFY by reading back a sample to confirm formatting applied correctly
  - Formatting rules by zone type:
    - Titles: merged, bold, larger font, colored background
    - Headers: bold, background fill, bottom border, frozen row
    - Data: number formatting, alternating rows (optional), alignment
    - Sub-totals: bold, top border thin, number formatting matching data
    - Grand totals: bold, top border medium or double, number formatting
  - Error handling: if structure is ambiguous, format conservatively and report uncertainty

#### Task 2.2: Template gallery integration (always-sub-agent path)
- **Goal:** When a user selects a template from the gallery dropdown, the Template Builder sub-agent is ALWAYS invoked instead of the naive `design_only` auto-detect
- **Changes to gallery flow:**
  1. User clicks a template in the dropdown gallery
  2. System calls `apply_template` with `action: "preview"` to extract the palette + typography of the chosen template
  3. System invokes `delegate_task` with `role: "template-builder"`, passing the extracted palette/typography as context
  4. Template Builder reads the data, detects zones, applies the template design to the correct zones
  5. Result is displayed to the user
- **Where to wire this:**
  - `src/template-gallery/template-catalog.ts` — intercept the apply action
  - `src/tools/apply-template.ts` — when `mode: "design_only"`, route through delegate_task instead of the naive auto-detect
  - The `mode: "full"` path (blank sheet with sample data) remains unchanged
- **Template Builder receives:**
  - `palette`: TemplatePalette (titleBg, headerBg, accentBg, totalBg, etc.)
  - `typography`: TemplateTypography (fontFamily, titleSize, headerSize, bodySize)
  - `alternatingRows`: boolean
  - `templateName`: for reference in the result summary
- **No fallback:** There is NO fallback to the old auto-detect. The sub-agent handles all cases (simple and complex data).

#### Task 2.3: Complex layout test cases
- **Create:** `tests/agents/template-builder.test.ts`
- **Test scenarios:**
  - Simple flat table (5 cols, 20 rows) → should format correctly
  - Multi-section P&L (Revenue section, COGS section, OpEx section, each with sub-totals)
  - Spreadsheet with title + meta fields + header + data + total
  - Sheet with merged header cells spanning multiple columns
  - Mixed data types (currency, percentages, dates, text in same sheet)
  - Sheet with existing formatting that should be overridden
  - Sheet with formulas in total rows (verify formulas preserved, only formatting changed)

### Phase 3: Remaining Sub-Agents (~1-2 weeks)

#### Task 3.1: Analyst role prompt
- Focus on data comprehension, pattern detection, summarization
- Read-only — never calls write/mutate tools
- Pre-loaded context: full workbook blueprint

#### Task 3.2: Builder role prompt
- Focus on structure creation, formula writing
- Pre-loads `power-user-patterns` skill for advanced formulas
- Includes formula construction rules from `financial-modeling` skill

#### Task 3.3: Researcher role prompt
- Focus on data sourcing, web search, fact extraction
- Returns structured data (JSON/CSV) that other sub-agents can consume
- Includes attribution/source tracking

#### Task 3.4: Modeler role prompt
- Pre-loads `financial-modeling` skill automatically
- Includes DCF, LBO, 3-statement model patterns
- Color-coding conventions (blue inputs, black formulas, green cross-sheet)

#### Task 3.5: Debugger role prompt
- Focus on error diagnosis, formula tracing
- Systematic approach: find all errors → trace each → propose fix → apply with user approval

#### Task 3.6: Stylist role prompt
- Focus on visual design, formatting best practices
- Knows Excel's limitations (available colors, font support, border styles)
- Can apply consistent themes across multiple sheets

### Phase 4: UI + Polish (~3-5 days)

#### Task 4.1: Sub-agent tool card renderer
- **File:** `src/ui/tool-renderers.ts`
- **Content:** Custom renderer for `delegate_task` tool results showing:
  - Which sub-agent was invoked (role badge with color)
  - Task description
  - Status (completed/failed/max_turns)
  - Number of tool calls made
  - Summary of actions taken
  - Expandable detail view with inner tool calls

#### Task 4.2: Sub-agent input humanizer
- **File:** `src/ui/humanize-params.ts`
- **Content:** Humanize `delegate_task` params:
  - Role → colored badge ("🔍 Analyst", "🔨 Builder", "🎨 Stylist", etc.)
  - Task → truncated task description
  - Context → collapsed detail

#### Task 4.3: System prompt fix — template count
- **File:** `src/prompt/system-prompt.ts` and `src/tools/capabilities.ts`
- **Fix:** Update the template count from "6 bundled" to "11 bundled" with correct names
- **Also update:** `CORE_TOOL_CAPABILITY_METADATA.apply_template.promptDescription`

#### Task 4.4: Delegation status indicator
- **File:** `src/ui/pi-sidebar.ts` (or new component)
- **Content:** When a sub-agent is running, show a status bar:
  - "Template Builder working... (3/15 turns)"
  - Animated indicator
  - Cancel button to abort the sub-agent

---

## Files to Create

| File | Purpose |
|---|---|
| `src/agents/roles.ts` | Role registry + types |
| `src/agents/sub-agent-runner.ts` | Inner LLM loop executor |
| `src/agents/roles/analyst.ts` | Analyst role definition + prompt |
| `src/agents/roles/builder.ts` | Builder role definition + prompt |
| `src/agents/roles/stylist.ts` | Stylist role definition + prompt |
| `src/agents/roles/template-builder.ts` | Template Builder role definition + prompt |
| `src/agents/roles/researcher.ts` | Researcher role definition + prompt |
| `src/agents/roles/modeler.ts` | Modeler role definition + prompt |
| `src/agents/roles/debugger.ts` | Debugger role definition + prompt |
| `src/tools/delegate-task.ts` | delegate_task tool definition |
| `tests/agents/template-builder.test.ts` | Template Builder test cases |
| `tests/agents/sub-agent-runner.test.ts` | Sub-agent runner unit tests |
| `tests/agents/delegate-task.test.ts` | delegate_task tool tests |

## Files to Modify

| File | Change |
|---|---|
| `src/tools/index.ts` | Add `createDelegateTaskTool()` to `createAllTools()` |
| `src/tools/names.ts` | Add `delegate_task` if core, or keep auxiliary |
| `src/tools/capabilities.ts` | Fix template count (6 → 11), add delegate_task metadata |
| `src/prompt/system-prompt.ts` | Add Delegation section, fix template count |
| `src/ui/tool-renderers.ts` | Add renderer for delegate_task results |
| `src/ui/humanize-params.ts` | Add humanizer for delegate_task params |
| `src/context/tool-disclosure.ts` | Ensure delegate_task is in full bundle |

## Constraints

- **No breaking changes** to existing tools or behavior
- **Same LLM provider** used for sub-agents as the main agent (inherit model selection)
- **WorkbookCoordinator** must be shared (not per-sub-agent)
- **Token budget awareness:** sub-agent inner loops should use compact context (no full workbook blueprint every turn)
- **TypeScript policy:** no `@ts-ignore`, no `any`, strict types
- **Follow AGENTS.md:** update registry.ts, tool-renderers.ts, humanize-params.ts, tool-disclosure.ts, system-prompt.ts in the same PR when adding delegate_task

## Verification

- `npm run check` — lint + typecheck pass
- `npm run build` — production build succeeds, check chunk sizes
- `npm run test:context` — tool disclosure + context tests pass
- New tests: `tests/agents/*.test.ts` all pass
- Manual Excel smoke test: delegate_task works end-to-end with Template Builder on a complex spreadsheet

## Open Questions

1. Should sub-agent conversations be visible in the sidebar message history, or collapsed into the tool card?
2. Should sub-agents inherit the parent's `conventions` and `instructions`?
3. Cost control: should there be a per-session token budget for sub-agent calls?
4. Should the user be able to pick which sub-agent handles a task, or always let the orchestrator decide?
5. ~~Should Template Builder only run on complex data?~~ **DECIDED: NO — Template Builder always runs on every template apply, simple or complex.**
