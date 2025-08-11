## Ultrasheets — Architecture & Implementation Notes

### TL;DR

- **LLM** runs server-side via Vercel AI SDK (`streamText`) in `app/api/chat/route.ts` and drives the app by calling declarative "tools".
- **Tools** are split in two layers:
  - Server-declared tool intents that return a `clientSideAction` (e.g., execute a Univer tool in the browser).
  - Modern Univer tools implemented with a shared context layer in `lib/` (e.g., `format_currency_column`).
- **Frontend** (`components/chat-sidebar.tsx`) streams assistant parts and executes `clientSideAction`s against the live Univer workbook, logging actions for follow‑ups.

---

### How the LLM works

- Model: `openai("gpt-4o-mini")` orchestrated by `streamText` in `app/api/chat/route.ts`.
- System prompt enforces a strict workflow:
  - “Analyze first” by calling `get_workbook_snapshot` (no guessing columns/ranges).
  - Choose the right tool based on sheet intelligence (tables, numeric columns, spatial zones).
  - Return short confirmations; do not dump tables in chat.
- Tools are declared with `tool({...})`. Most tools return a `clientSideAction` that the frontend performs using Univer’s live API (no stale state).
- Deterministic behavior: temperature 0.2 and explicit guidance (e.g., currency formatting policies, totals policies, chart rules).

---

### Flow: from a message to an action

1. User types in the sidebar (`useChat` in `components/chat-sidebar.tsx`).
2. Request hits `POST /api/chat` with lightweight `clientEnv` and recent action hints.
3. The model responds with text + tool invocations.
4. The frontend reads streaming parts and executes each `clientSideAction` exactly once (deduped by `toolCallId`).
5. Univer APIs mutate the live sheet; we recalc formulas and append an entry to `window.ultraActionLog` for context‑aware follow‑ups.

```mermaid
sequenceDiagram
  participant U as User
  participant FE as Chat Sidebar (React)
  participant API as /api/chat (LLM)
  participant UA as Univer API (Browser)

  U->>FE: "add price per kg"
  FE->>API: messages + recentActions
  API->>API: system prompt + tools
  API-->>FE: text + tool(invocation)
  FE->>UA: executeUniverTool(smart_add_column)
  UA-->>FE: ok; recalc; log ultraActionLog
  FE->>API: next user msg "format it in USD"
  API-->>FE: tool(format_currency_column, columnName=LAST_ADDED)
  FE->>UA: apply number format on target range
```

---

### How a tool is made

There are two complementary layers.

- Server-declared tool wrappers (request orchestration)

  - Location: `app/api/chat/route.ts` under `tools: { ... }`.
  - Shape: `tool({ description, parameters, execute })`.
  - Behavior: return `{ clientSideAction: { type, toolName, params } }` so the frontend can execute it against the live workbook.

- Modern Univer tools (execution logic + unified context)
  - Location: `lib/tools/modern-tools.ts` and shared primitives in `lib/`.
  - Created via `createSimpleTool` from `lib/tool-executor.ts`.
  - Run with a rich `UniversalToolContext` from `lib/universal-context.ts` that provides:
    - `tables`, `columns`, `spatialMap`, `findTable`, `findColumn`, `getColumnRange`, `buildSumFormula`, `findOptimalPlacement`.
  - Example: `FormatCurrencyColumnTool` finds the target column, derives the data range, maps the currency to a format pattern, and calls Univer to set the number format.

Minimal recipe to introduce a new modern tool

1. Implement it with `createSimpleTool` and export in `MODERN_TOOLS` (see `lib/tools/modern-tools.ts`).
2. Register it on the client (your Univer bootstrap) so `window.executeUniverTool` can run it.
3. Add a corresponding server tool stub in `app/api/chat/route.ts` that returns a `clientSideAction` mapping user intent to the modern tool.

---

### Hacks / tricks used here

- **Recent action memory for pronouns**

  - We append compact entries to `window.ultraActionLog` after each client‑side execution.
  - The server includes `RECENT ACTIONS` and `LAST_ADDED_COLUMN` in the prompt, so follow‑ups like “format it” resolve correctly.
  - `format_currency_column` prefers the last added column before fuzzy matching.

- **Special range tokens**

  - `format_cells` accepts `"__AUTO_HEADERS__"` and resolves it to the header row of the primary data region on the client with `getCleanSheetContext`.

- **Deterministic tool execution**

  - Streaming parts are deduped by `toolCallId` in `components/chat-sidebar.tsx` to prevent duplicate executions during SSE updates.

- **Formula recalculation**

  - After writes or inserts we explicitly call `univerAPI.getFormula().executeCalculation()` to update dependent values.

- **Context caching with safety**

  - `lib/universal-context.ts` caches an analyzed snapshot for ~5s and exposes helpers; tools can invalidate/refresh when needed.

- **Robust guards**
  - Server route now uses null‑safe checks when `workbookData.sheets` is not provided to avoid crashes in selection handling.

---

### Example tool (pseudocode): format_currency_column

Goal: find the right currency column and apply the chosen currency format to data cells (no headers).

1. Modern tool logic (browser; uses Univer API via our context)

```
on format_currency_column(currency = "USD", decimals = 2, columnName?) {
  table = context.findTable()

  // pick the column
  target = columnName ? context.findColumn(columnName)
         : lastAddedColumnFrom(ultraActionLog)
         : detectByNameOrType(table.columns)  // price/cost/amount/... or isCurrency

  dataRange = context.getColumnRange(target.name, includeHeader = false)
  pattern = currencyMap[currency] or `${currency}#,##0.${repeat("0", decimals)}`
  context.fWorksheet.getRange(dataRange).setNumberFormat(pattern)
  return { column: target.name, range: dataRange, currency }
}
```

2. Registration (made available to the runtime)

```
export MODERN_TOOLS = [ ..., format_currency_column, ... ]
```

3. Server wrapper (LLM-visible tool)

```
tool name: format_currency_column
params: { currency, decimals, columnName? }
execute: return clientSideAction { type: executeUniverTool, toolName: format_currency_column, params }
```

4. Frontend execution

```
on tool result with clientSideAction:
  window.executeUniverTool(toolName, params)
  // This calls our modern tool in the browser against the live sheet
```

Effect: “format it in USD” after “add price per kg” targets the just-added column and formats its data cells correctly.

---

### Key files to skim

- `app/api/chat/route.ts` — LLM system prompt, tool declarations, guidance.
- `components/chat-sidebar.tsx` — streaming UI, `clientSideAction` executor, ultraActionLog.
- `lib/universal-context.ts` — one-stop context with smart helpers.
- `lib/tool-executor.ts` — tool base class, registry, and `createSimpleTool` helper.
- `lib/tools/modern-tools.ts` — modern tools (e.g., `format_currency_column`, `add_smart_totals`, `generate_chart`).

---

### Demo script (suggested)

1. “Add price per kg.” → new column appears with per‑row formula.
2. “Format it in USD.” → resolves to last added column and applies currency format.
3. “Add totals.” → `add_smart_totals` inserts SUMs below numeric columns.
4. “Generate a column chart titled ‘Weight vs Price’.” → chart appears to the right using spatial placement.
