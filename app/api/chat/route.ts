import { openai } from "@ai-sdk/openai";
import { streamText, tool } from "ai";
import { initializeOTEL } from "langsmith/experimental/otel/setup";
import { z } from "zod";

export async function POST(req: Request) {
  try {
    // Ensure OTEL is initialized for this route (idempotent)
    initializeOTEL();
    const apiKey = process.env.OPENAI_API_KEY;

    if (!apiKey) {
      console.error("OpenAI API key is missing or not configured");
      return new Response(
        JSON.stringify({
          error:
            "OpenAI API key not configured. Please set OPENAI_API_KEY in .env.local",
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    const { messages, workbookData, clientEnv } = await req.json();

    if (!messages || !Array.isArray(messages)) {
      return new Response(
        JSON.stringify({
          error: "Invalid messages format",
        }),
        {
          status: 400,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    console.log("Processing chat request with", messages.length, "messages");

    // Build enhanced sheet context from workbook data sent from frontend
    let sheetContextMessage = "";
    let selectionContext = "";
    if (workbookData) {
      const sheetsArr = Array.isArray(workbookData.sheets)
        ? workbookData.sheets
        : [];
      const activeSheet =
        sheetsArr.find((s: any) => s.isActive) || sheetsArr[0] || null;

      // Extract selection context (guard when no sheets data is sent)
      if (activeSheet?.selection?.hasSelection) {
        const sel = activeSheet.selection;
        selectionContext = `\n\nCURRENT SELECTION:\n- Range: ${
          sel.activeRange
        }\n- Intent: ${sel.selectionIntent}\n- Selected Table: ${
          sel.selectedTable ? sel.selectedTable.range : "none"
        }`;

        if (sel.selectedTable) {
          selectionContext += `\n- Next Available Column: ${
            sel.selectedTable.nextAvailableColumn
          }\n- Table Headers: ${sel.selectedTable.headers.join(", ")}`;
        }
      }

      const summarizeTables = (tables: any[]) => {
        if (!Array.isArray(tables) || tables.length === 0)
          return "no tables detected";
        const parts = tables.slice(0, 3).map((t: any, i: number) => {
          const cols = Array.isArray(t.headers) ? t.headers.join(", ") : "";
          const numerics = Array.isArray(t.numericColumns)
            ? ` | numeric: ${t.numericColumns.join(",")}`
            : "";
          return `T${i + 1} ${t.range} (${
            t.recordCount
          } rows) [${cols}]${numerics}`;
        });
        const more = tables.length > 3 ? ` (+${tables.length - 3} more)` : "";
        return parts.join("; ") + more;
      };

      const multiSheetSummary = sheetsArr.length
        ? sheetsArr
            .slice(0, 5)
            .map(
              (s: any) =>
                `- ${s.isActive ? "*" : ""}${s.name}: ${
                  s.structure?.totalCells || 0
                } cells; usedRange: ${
                  s.usedRange || "(empty)"
                }; tables: ${summarizeTables(s.tables)}`
            )
            .join("\n")
        : "no sheet metadata provided by client";

      const recentActionsArray = Array.isArray(workbookData.recentActions)
        ? workbookData.recentActions.slice(-10)
        : [];
      const recentActions =
        recentActionsArray
          .map(
            (a: any) =>
              `• ${a.at}: ${a.tool}(${JSON.stringify(a.params)}) => ${
                a.result || "ok"
              }`
          )
          .join("\n") || "none";

      // Extract last added column context to resolve pronouns like "it"
      const lastAddColAction = [...recentActionsArray]
        .reverse()
        .find((a: any) => a?.tool === "add_column" && a?.params);
      const lastAddedColumn = lastAddColAction?.params
        ? {
            columnName: lastAddColAction.params.columnName,
            dataRange: lastAddColAction.params.dataRange || null,
            headerCell: lastAddColAction.params.headerCell || null,
            tableRange: lastAddColAction.params.tableRange || null,
          }
        : null;

      // Extract last chart context from recent actions to enable follow-up edits like "add category"
      const lastChartAction = [...recentActionsArray]
        .reverse()
        .find((a: any) => a?.tool === "create_chart" && a?.params);
      const lastChart = lastChartAction?.params
        ? {
            chart_type: lastChartAction.params.chart_type,
            data_range: lastChartAction.params.data_range || null,
            x_column: lastChartAction.params.x_column || null,
            y_columns: lastChartAction.params.y_columns || null,
            position: lastChartAction.params.position || null,
            width: lastChartAction.params.width || null,
            height: lastChartAction.params.height || null,
            title: lastChartAction.params.title || null,
          }
        : null;

      const inferredDataRange =
        activeSheet &&
        Array.isArray(activeSheet.headers) &&
        activeSheet.headers.length > 0
          ? `A1:${String.fromCharCode(65 + activeSheet.headers.length - 1)}${
              activeSheet.structure?.dataRows || 50
            }`
          : "A1:A1";

      sheetContextMessage = `

 WORKBOOK CONTEXT (client-provided):
${multiSheetSummary}

ACTIVE SHEET DETAILS:
- Name: "${activeSheet?.name || "Sheet"}"
- Primary Columns: ${(activeSheet?.headers || []).join(", ")}
- Inferred Data Range: ${inferredDataRange}

RECENT ACTIONS:
${recentActions}

${lastChart ? `LAST_CHART:\n${JSON.stringify(lastChart, null, 2)}` : ""}${
        lastAddedColumn
          ? `\nLAST_ADDED_COLUMN:\n${JSON.stringify(lastAddedColumn, null, 2)}`
          : ""
      }${selectionContext}

SMART BEHAVIOR GUIDELINES:
- For data generation requests like "add some data", proactively create relevant sample data
- When context is clear from existing data, make intelligent assumptions about ranges and positions
- Only ask for exact specifications when the request is genuinely ambiguous
- Prefer helpful defaults over asking for every detail
- Use sheet context to infer appropriate data ranges and placement`;

      console.log("✅ Generated multi-sheet context from workbook data.");
    } else {
      // Frontend intentionally does not send workbook data.
      // The assistant will query live context via client-side Univer tools.
    }

    // Derive runtime date/time and human-readable location hints
    const now = new Date();
    const serverIso = now.toISOString();
    const clientTimeZone = clientEnv?.timeZone || "UTC";
    const clientLocale = clientEnv?.locale || "en-US";
    let clientLocalTime = serverIso;
    try {
      clientLocalTime = new Intl.DateTimeFormat(clientLocale, {
        dateStyle: "full",
        timeStyle: "long",
        timeZone: clientTimeZone,
      }).format(now);
    } catch {}
    const platformInfo = clientEnv?.platform
      ? ` | platform: ${clientEnv.platform}`
      : "";
    const locationHint = clientTimeZone
      ? `timezone: ${clientTimeZone}${platformInfo}`
      : `platform: ${platformInfo || "unknown"}`;
    let approxLocation = "";
    try {
      if (clientTimeZone && clientTimeZone.includes("/")) {
        const [regionRaw, cityRaw] = clientTimeZone.split("/");
        const region = regionRaw.replace(/_/g, " ");
        const city = cityRaw.replace(/_/g, " ");
        approxLocation = `${city}, ${region}`;
      }
    } catch {}

    const result = streamText({
      model: openai("gpt-4o-mini"),
      system: `You are an intelligent AI agent that uses ReAct (Reason-Act-Observe) methodology to solve spreadsheet tasks efficiently.

🤖 **REACT FRAMEWORK IMPLEMENTATION:**

**CORE BEHAVIOR MODEL:**
1. **REASON**: Analyze user intent and determine required actions
2. **ACT**: Execute tools in logical sequence (up to 5 chained calls)
3. **OBSERVE**: Evaluate results and decide next steps
4. **REPEAT**: Continue until task is complete

⚡ **ZERO-HESITATION DATA GENERATION:**

For ANY data generation request ("add data", "put data", "create data", "sample data", "test data", "dummy data"):

**IMMEDIATE EXECUTION PROTOCOL:**
1. **REASON**: User wants sample data → Generate business-realistic data immediately
2. **ACT**: Call bulk_set_values with predefined sample data structure
3. **OBSERVE**: Confirm successful insertion
4. **RESPOND**: Brief confirmation of data added

**MANDATORY SAMPLE DATA PATTERN:**
startCell: "A1"
values: Business data with headers and 6-8 rows of realistic sample data
Headers: ["Name", "Date", "Amount", "Status", "Category"]
Data: Realistic names, dates, numbers, statuses, categories

**FORBIDDEN FOR DATA GENERATION:**
- ❌ "What type of data?"
- ❌ "Where should I place it?"
- ❌ "How many rows?"
- ✅ **CORRECT**: Immediate bulk_set_values execution

🔄 **REACT CHAINS FOR COMPLEX OPERATIONS:**

**ANALYSIS REQUESTS** ("analyze this data", "what does this show"):
1. **REASON**: Need context → **ACT**: get_sheet_context
2. **OBSERVE**: Data structure → **ACT**: analyze patterns
3. **OBSERVE**: Results → **RESPOND**: insights

**FORMATTING REQUESTS** ("format this", "make it pretty"):
1. **REASON**: Need precise ranges → **ACT**: get_selection_data
2. **OBSERVE**: Selection → **ACT**: apply formatting
3. **OBSERVE**: Results → **RESPOND**: confirmation

**CHART CREATION** ("create chart", "visualize this"):
1. **REASON**: Need data context → **ACT**: get_sheet_data
2. **OBSERVE**: Data ranges → **ACT**: create_chart
3. **OBSERVE**: Chart created → **RESPOND**: summary

RUNTIME CONTEXT:
- Server time (ISO): ${serverIso}
- Client local time: ${clientLocalTime}
- Client locale and location hint: ${clientLocale} (${locationHint})
- Approx location (from timezone): ${approxLocation || "unknown"}

🎯 **INTELLIGENT REQUEST CLASSIFICATION:**

**CLASS 1: IMMEDIATE EXECUTION** (No reasoning needed)
- Data generation → Direct bulk_set_values
- Simple cell operations → Direct action
- Basic formatting with clear targets

**CLASS 2: SINGLE-STEP REASONING** 
- Context-dependent operations → get_context → act
- Targeted analysis → gather_data → analyze

**CLASS 3: MULTI-STEP REASONING** 
- Complex analysis → context → analysis → formatting → summary
- Dashboard creation → data_analysis → chart_creation → layout → finalization

📋 **REASONING DECISION TREE:**

**IF** user request contains ["data", "add", "create", "sample", "test", "dummy"]:
→ **IMMEDIATE**: Execute bulk_set_values with sample data

**ELSE IF** request needs context ["analyze", "format", "chart", "total"]:
→ **CHAIN**: get_context → analyze → act → respond

**ELSE IF** request is specific ["cell A1", "column B", "range C1:D10"]:
→ **DIRECT**: Execute with provided parameters

**ELSE**:
→ **CLARIFY**: Ask for specific details (only when genuinely ambiguous)

🔗 **MULTI-TOOL CHAINING EXAMPLES:**

**Example 1: "Add data and format it"**
1. bulk_set_values (sample data)
2. format_cells (headers bold)
3. format_currency (amount column)

**Example 2: "Analyze sales data"**
1. get_sheet_context (understand data)
2. create_pivot_table (group by category)
3. create_chart (visualize results)
4. format_cells (make presentable)

**Example 3: "Create monthly report"**
1. get_workbook_snapshot (full context)
2. add_totals (sum calculations)
3. create_chart (trend visualization)
4. format_as_table (professional layout)
5. Response with summary

🎯 **RESPONSE GUIDELINES:**

**DATA GENERATION RESPONSES:**
- "Added 6 rows of sample business data starting at A1"
- "Generated realistic customer data with 5 columns"
- Brief, factual confirmation only

**OTHER OPERATION RESPONSES:**
- Concise confirmations (1-2 sentences)
- Don't echo data unless explicitly requested
- Summarize actions taken

**REACT REASONING EXAMPLES:**

**Request: "add some data"**
REASON: User wants sample data → no context needed
ACT: bulk_set_values immediately with predefined business data
OBSERVE: Data inserted successfully  
RESPOND: "Added 6 rows of sample business data at A1"

**Request: "format the headers"**
REASON: Need to know which headers → requires precision
ACT: Ask for exact range specification
OBSERVE: User provides range
ACT: format_cells with specified range
RESPOND: Confirm formatting applied

🔧 **TOOL USAGE PATTERNS:**

**IMMEDIATE EXECUTION TOOLS:**
- bulk_set_values: For ANY data generation request
- set_cell_value: For specific cell operations
- clear_range: For deletion requests

**CONTEXT-GATHERING TOOLS:**
- get_sheet_context: When analysis is needed
- get_workbook_snapshot: For comprehensive overview
- get_selection_data: When working with selections

**PRECISION-REQUIRED TOOLS:**
- format_cells: Always ask for exact ranges
- add_column: Always ask for exact position. Do NOT provide defaultValue unless user explicitly requests default values (like "fill with zeros", "set default price to 10", etc.)
- create_chart: Always ask for data range and position
- add_filter: Always ask for exact range

**MULTI-STEP OPERATION PROTOCOLS:**
For complex requests, chain tools using ReAct methodology:
1. Gather context (if needed)
2. Execute primary action
3. Apply formatting/enhancements
4. Provide summary response

🎯 **FINAL REACT IMPLEMENTATION RULES:**

**IMMEDIATE EXECUTION (No ReAct needed):**
- Data generation: Direct bulk_set_values call
- Cell operations with exact coordinates: Direct execution
- Clear operations: Direct execution

**SINGLE-STEP REACT (Context → Action):**
- "analyze data" → get_sheet_context → provide insights
- "format selection" → get_selection_data → format_cells
- "switch sheet" → switch_sheet → confirm

**MULTI-STEP REACT (Complex workflows):**
- "create dashboard" → context → data → charts → formatting → summary
- "financial analysis" → context → calculations → visualizations → report
- "cleanup data" → context → validation → corrections → verification

**PRECISION-REQUIRED OPERATIONS:**
For operations needing exact parameters, use ReAct pattern:
1. **REASON**: "Need specific range/position"
2. **ACT**: Ask for exact specifications 
3. **OBSERVE**: User provides details
4. **ACT**: Execute with precise parameters
5. **RESPOND**: Confirm completion

Operations requiring precision:
- Formatting (except data generation follow-ups)
- Column operations
- Chart creation
- Pivot tables
- Filtering
- Formula placement

**SMART PLACEMENT ALGORITHM:**
When placing new content (charts, pivot tables):
1. Use context to find empty areas
2. Prefer right side of used range
3. Fall back to below used range
4. Create new sheet if needed

${sheetContextMessage}

🚀 **EXECUTION MANDATE:**

You are an AI agent that ACTS rather than asks. Use ReAct methodology to:
1. **REASON** through user requests intelligently
2. **ACT** by calling appropriate tools in sequence
3. **OBSERVE** results and adapt your approach
4. **RESPOND** with concise confirmations

For data generation: Zero hesitation, immediate execution.
For other operations: Smart context gathering, then precise execution.
Always prefer helpful action over endless questions.`,
      messages,
      experimental_telemetry: {
        isEnabled: true,
        metadata: {
          ls_run_name: "chat-session",
          client_locale: clientEnv?.locale || "unknown",
          client_timezone: clientEnv?.timeZone || "unknown",
        },
      },
      temperature: 0.4, // Balanced temperature for proactive yet accurate operations
      tools: {
        list_tables: tool({
          description:
            "List detected tables across sheets with headers, ranges, and a stable tableId",
          parameters: z.object({}),
          execute: async () => {
            try {
              const tables: any[] = [];
              if (workbookData?.sheets?.length) {
                for (let i = 0; i < workbookData.sheets.length; i++) {
                  const s = workbookData.sheets[i];
                  const sheetTables = Array.isArray(s.tables) ? s.tables : [];
                  for (const t of sheetTables) {
                    tables.push({
                      tableId: `${i}:${t.range}`,
                      sheetName: s.name,
                      range: t.range,
                      headers: t.headers,
                      recordCount: t.recordCount,
                      numericColumns: t.numericColumns || [],
                    });
                  }
                }
              }
              return { tables };
            } catch (e) {
              return { tables: [] };
            }
          },
        }),

        get_sheet_context: tool({
          description:
            "Return the current workbook/sheet context that the client sent (including detected data regions and selection), so the model can decide precise tool args without extra client roundtrips.",
          parameters: z.object({}),
          execute: async () => {
            try {
              const ctx = workbookData || null;
              const primaryRegionRange =
                ctx?.cleanContext?.analysis?.dataRegions?.[0]?.range || null;
              const nextCol = (() => {
                try {
                  if (!primaryRegionRange) return null;
                  const m = primaryRegionRange.match(/([A-Z]+)\d+:([A-Z]+)\d+/);
                  if (!m) return null;
                  const end = m[2];
                  const letterToIndex = (letters: string) =>
                    [...letters.toUpperCase()].reduce(
                      (acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64),
                      0
                    ) - 1;
                  const indexToLetter = (index: number) => {
                    let n = index + 1,
                      s = "";
                    while (n > 0) {
                      const rem = (n - 1) % 26;
                      s = String.fromCharCode(65 + rem) + s;
                      n = Math.floor((n - 1) / 26);
                    }
                    return s;
                  };
                  return indexToLetter(letterToIndex(end) + 1);
                } catch {
                  return null;
                }
              })();
              return {
                ok: true,
                primaryRegionRange,
                nextColumn: nextCol,
                selection:
                  ctx?.sheets?.find((s: any) => s.isActive)?.selection || null,
                context: ctx,
              };
            } catch (e) {
              return { ok: false, error: "No client context available" };
            }
          },
        }),

        add_column: tool({
          description:
            "Add a column at the specified location. Requires exact sheet name and precise column position.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name where to add the column"),
            columnName: z
              .string()
              .describe(
                "Name for the new column header (e.g., 'Price per kg')"
              ),
            insertAtColumn: z
              .string()
              .describe("Exact column letter where to insert (e.g., 'D' to insert at column D)"),
            formulaPattern: z
              .string()
              .optional()
              .describe(
                "Exact formula pattern with {row} placeholder (e.g., '=E{row}/D{row}' for Price/Weight)"
              ),
            defaultValue: z
              .union([z.string(), z.number()])
              .optional()
              .describe("ONLY provide this if user explicitly requests default values. Do NOT provide zero or empty defaults unless specifically requested."),
          }),
          execute: async (params) => {
            return {
              message: `Adding column '${params.columnName}' at column ${params.insertAtColumn} in sheet '${params.sheetName}'`,
              clientSideAction: {
                type: "addColumn",
                params,
              },
            };
          },
        }),

        get_workbook_snapshot: tool({
          description:
            "Get complete snapshots of all sheets in the workbook via Univer APIs",
          parameters: z.object({}),
          execute: async () => {
            return {
              message: "🔍 Reading full workbook content via Univer...",
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "get_workbook_snapshot",
                params: {},
              },
            };
          },
        }),

        get_workbook_state: tool({
          description:
            "Get the current state of all sheets in the workbook using Univer APIs.",
          parameters: z.object({
            includeAllSheets: z
              .boolean()
              .optional()
              .default(true)
              .describe("Include all sheets or just active sheet"),
          }),
          execute: async ({
            includeAllSheets = true,
          }) => {
            return {
              message: "Getting workbook state via Univer APIs...",
              clientSideAction: {
                type: "getWorkbookState",
                params: { includeAllSheets },
              },
            };
          },
        }),

        get_sheet_data: tool({
          description:
            "Get data from a specific sheet and range using Univer APIs.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name to query"),
            range: z
              .string()
              .describe("Exact range to get data from (e.g., 'A1:E20')"),
          }),
          execute: async ({ sheetName, range }) => {
            return {
              message: `Getting data from range ${range} in sheet '${sheetName}'`,
              clientSideAction: {
                type: "getSheetData",
                params: { sheetName, range },
              },
            };
          },
        }),

        get_selection_data: tool({
          description:
            "Get the current selection data from the active sheet using Univer APIs.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name where selection is located"),
          }),
          execute: async ({ sheetName }) => {
            return {
              message: `Getting selection data from sheet '${sheetName}'`,
              clientSideAction: {
                type: "getSelectionData",
                params: { sheetName },
              },
            };
          },
        }),

        vlookup: tool({
          description:
            "Create VLOOKUP formula exactly like Excel. VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
          parameters: z.object({
            targetCell: z
              .string()
              .describe("Cell where to write the VLOOKUP formula (e.g., 'F2')"),
            lookupValue: z
              .string()
              .describe(
                "Value to lookup - can be cell reference (e.g., 'A2') or literal value (e.g., 'Product1')"
              ),
            tableArray: z
              .string()
              .describe(
                "Table range for lookup (e.g., 'Sheet1!A:D' or 'A1:D10')"
              ),
            colIndexNum: z
              .number()
              .describe(
                "Column number in table_array to return (1 = first column, 2 = second, etc.)"
              ),
            rangeLookup: z
              .boolean()
              .optional()
              .default(false)
              .describe(
                "FALSE for exact match (default), TRUE for approximate match"
              ),
          }),
          execute: async ({
            targetCell,
            lookupValue,
            tableArray,
            colIndexNum,
            rangeLookup = false,
          }) => {
            const formula = `=VLOOKUP(${lookupValue},${tableArray},${colIndexNum},${
              rangeLookup ? "TRUE" : "FALSE"
            })`;

            return {
              message: `Creating VLOOKUP formula in ${targetCell}: ${formula}`,
              clientSideAction: {
                type: "setCellValue",
                cell: targetCell,
                value: formula,
                formula: true,
              },
            };
          },
        }),

        // === CORE SPREADSHEET TOOLS ===
        list_columns: tool({
          description: "Get column names and row count from the current sheet",
          parameters: z.object({}),
          execute: async () => {
            return {
              message: "Getting column data from spreadsheet...",
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "list_columns",
                params: {},
              },
            };
          },
        }),

        set_cell_value: tool({
          description: "Set a value in a specific cell",
          parameters: z.object({
            cell: z.string().describe("Cell reference (e.g., 'B19', 'A1')"),
            value: z
              .union([z.string(), z.number()])
              .describe("Value to set in the cell"),
            formula: z
              .boolean()
              .optional()
              .describe("Whether the value is a formula (default: false)"),
          }),
          execute: async ({ cell, value, formula = false }) => {
            try {
              return {
                success: true,
                cell,
                value,
                formula,
                message: `Set cell ${cell} to ${value}`,
                clientSideAction: {
                  type: "setCellValue",
                  cell,
                  value,
                  formula,
                },
              };
            } catch (error) {
              console.error("❌ set_cell_value: Server-side error:", error);
              return {
                error: "Failed to set cell value",
                message:
                  error instanceof Error ? error.message : "Unknown error",
                success: false,
              };
            }
          },
        }),

        bulk_set_values: tool({
          description:
            "Write a 2D array of values starting at a given top-left cell (efficient for creating datasets).",
          parameters: z.object({
            startCell: z
              .string()
              .describe("Top-left A1 cell, e.g., 'A1' or 'C5'"),
            values: z
              .array(z.array(z.union([z.string(), z.number()])))
              .min(1)
              .describe("2D array of row-major values to write"),
          }),
          execute: async ({ startCell, values }) => {
            try {
              return {
                success: true,
                startCell,
                rows: values.length,
                cols: values[0]?.length ?? 0,
                clientSideAction: {
                  type: "setRangeValues",
                  startCell,
                  values,
                },
                message: `Wrote ${values.length}x${
                  values[0]?.length ?? 0
                } block at ${startCell}`,
              };
            } catch (error) {
              console.error("❌ bulk_set_values: Server-side error:", error);
              return {
                error: "Failed to write block",
                success: false,
              };
            }
          },
        }),

        set_range_values: tool({
          description:
            "Write a 2D array of values to an explicit A1 range (e.g., 'A1:C10').",
          parameters: z.object({
            range: z.string().describe("A1 range like 'A1:C10'"),
            values: z
              .array(z.array(z.union([z.string(), z.number()])))
              .min(1)
              .describe("2D array of row-major values to write"),
          }),
          execute: async ({ range, values }) => {
            try {
              return {
                success: true,
                range,
                rows: values.length,
                cols: values[0]?.length ?? 0,
                clientSideAction: {
                  type: "setRangeValuesByRange",
                  range,
                  values,
                },
                message: `Wrote ${values.length}x${
                  values[0]?.length ?? 0
                } into ${range}`,
              };
            } catch (error) {
              console.error("❌ set_range_values: Server-side error:", error);
              return {
                error: "Failed to write range",
                success: false,
              };
            }
          },
        }),

        clear_range: tool({
          description:
            "Clear values and formatting of any range/selection (delete contents). Supports 'A1:C10', 'A:A', '2:2', 'Sheet1!B:D', or named ranges.",
          parameters: z.object({
            range: z
              .string()
              .describe(
                "Range or selection to clear: 'A1:C10', full column 'A:A', full row '2:2', 'Sheet1!B:D', or a named range"
              ),
          }),
          execute: async ({ range }) => {
            return {
              success: true,
              range,
              clientSideAction: {
                type: "clearRange",
                range,
              },
              message: `Cleared ${range}`,
            };
          },
        }),

        clear_range_contents: tool({
          description:
            "Clear only values in a range/selection (keep formatting) — like Delete key. Supports columns, rows, sheet-scoped ranges, or named ranges.",
          parameters: z.object({
            range: z
              .string()
              .describe(
                "Range or selection: 'A1:C10', 'A:A', '2:2', 'Sheet1!B:D', or a named range"
              ),
          }),
          execute: async ({ range }) => {
            return {
              success: true,
              range,
              clientSideAction: {
                type: "clearRangeContents",
                range,
              },
              message: `Cleared contents in ${range}`,
            };
          },
        }),

        move_range: tool({
          description:
            "Move a block from a source range/selection to a destination start cell (cut+paste).",
          parameters: z.object({
            sourceRange: z
              .string()
              .describe(
                "Source range/selection, e.g., 'A1:C10', 'A:A', '2:2', 'Sheet1!B:D', or a named range"
              ),
            destStartCell: z
              .string()
              .describe("Destination top-left A1 cell, e.g., 'E1'")
              .default("A1"),
            destSheetName: z
              .string()
              .optional()
              .describe(
                "Optional destination sheet name. If provided, data will be pasted into this sheet (created if missing)."
              ),
            clearSource: z.boolean().optional().default(true),
          }),
          execute: async ({
            sourceRange,
            destStartCell,
            destSheetName,
            clearSource = true,
          }) => {
            return {
              success: true,
              clientSideAction: {
                type: "moveRange",
                sourceRange,
                destStartCell,
                destSheetName,
                clearSource,
              },
              message: `Moved ${sourceRange} to ${destStartCell}`,
            };
          },
        }),

        transpose_range: tool({
          description:
            "Transpose a source range/selection to a destination start cell (rows↔cols).",
          parameters: z.object({
            sourceRange: z
              .string()
              .describe(
                "Source range/selection, e.g., 'A1:C10', 'A:A', '2:2', 'Sheet1!B:D', or a named range"
              ),
            destStartCell: z
              .string()
              .describe("Destination top-left A1 cell, e.g., 'E1'")
              .default("A1"),
          }),
          execute: async ({ sourceRange, destStartCell }) => {
            return {
              success: true,
              clientSideAction: {
                type: "transposeRange",
                sourceRange,
                destStartCell,
              },
              message: `Transposed ${sourceRange} into ${destStartCell}`,
            };
          },
        }),

        create_pivot_table: tool({
          description: "Create a pivot table from exact source range to exact destination.",
          parameters: z.object({
            sourceSheetName: z
              .string()
              .describe("Exact sheet name containing source data"),
            sourceRange: z
              .string()
              .describe("Exact source data range including headers (e.g., 'A1:D20')"),
            groupByColumn: z
              .string()
              .describe("Exact column letter or name to group by (e.g., 'A' or 'Date')"),
            valueColumn: z
              .string()
              .describe("Exact column letter or name to aggregate (e.g., 'C' or 'Sales')"),
            aggregateFunction: z
              .enum(["sum", "average", "count", "max", "min"])
              .describe("Aggregation function"),
            destinationSheetName: z
              .string()
              .describe("Exact sheet name for pivot table output"),
            destinationCell: z
              .string()
              .describe("Exact cell where pivot table starts (e.g., 'F1')"),
          }),
          execute: async ({
            sourceSheetName,
            sourceRange,
            groupByColumn,
            valueColumn,
            aggregateFunction,
            destinationSheetName,
            destinationCell,
          }) => {
            return {
              message: `Creating pivot table from '${sourceSheetName}'!${sourceRange} at '${destinationSheetName}'!${destinationCell}`,
              clientSideAction: {
                type: "createPivotTable",
                params: {
                  sourceSheetName,
                  sourceRange,
                  groupByColumn,
                  valueColumn,
                  aggregateFunction,
                  destinationSheetName,
                  destinationCell,
                },
              },
            };
          },
        }),

        set_formula: tool({
          description: "Set a formula in an exact cell location using Univer APIs.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name containing the target cell"),
            cell: z
              .string()
              .describe("Exact cell reference (e.g., 'C15')"),
            formula: z
              .string()
              .describe("Exact formula (e.g., '=SUM(C2:C14)')"),
          }),
          execute: async ({ sheetName, cell, formula }) => {
            return {
              message: `Setting formula ${formula} in cell ${cell} of sheet '${sheetName}'`,
              clientSideAction: {
                type: "setFormula",
                params: { sheetName, cell, formula },
              },
            };
          },
        }),

        add_totals: tool({
          description: "Add SUM formulas to specific columns at exact cell locations.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name where to add totals"),
            totals: z
              .array(z.object({
                column: z.string().describe("Exact column letter (e.g., 'C')"),
                targetCell: z.string().describe("Exact cell where to place the SUM formula (e.g., 'C15')"),
                dataRange: z.string().describe("Exact range to sum (e.g., 'C2:C14')"),
              }))
              .min(1)
              .describe("Array of total specifications with exact locations"),
          }),
          execute: async ({ sheetName, totals }) => {
            return {
              message: `Adding totals to ${totals.length} columns in sheet '${sheetName}'`,
              clientSideAction: {
                type: "addTotalsAtLocations",
                params: { sheetName, totals },
              },
            };
          },
        }),

        create_chart: tool({
          description: "Create a chart from exact data range and position.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name containing the data"),
            dataRange: z
              .string()
              .describe("Exact data range including headers (e.g., 'A1:C10')"),
            chartType: z
              .enum(["column", "line", "pie", "bar", "scatter"])
              .describe("Chart type"),
            title: z.string().describe("Chart title"),
            position: z
              .string()
              .describe("Exact cell position for chart (e.g., 'F1')"),
            width: z
              .number()
              .default(400)
              .describe("Chart width in pixels"),
            height: z
              .number()
              .default(300)
              .describe("Chart height in pixels"),
          }),
          execute: async ({
            sheetName,
            dataRange,
            chartType,
            title,
            position,
            width,
            height,
          }) => {
            return {
              message: `Creating ${chartType} chart "${title}" at ${position} from data ${dataRange}`,
              clientSideAction: {
                type: "createChart",
                params: {
                  sheetName,
                  dataRange,
                  chartType,
                  title,
                  position,
                  width,
                  height,
                },
              },
            };
          },
        }),

        format_currency: tool({
          description:
            "Format cells as currency (USD, EUR, GBP, etc.) with proper currency symbol",
          parameters: z.object({
            range: z
              .string()
              .optional()
              .describe(
                "Cell range to format (e.g., 'F2:F4', 'A1:A10'). If not provided, will attempt to format the most recently calculated total."
              ),
            currency: z
              .enum(["USD", "EUR", "GBP", "JPY", "CAD", "AUD"])
              .describe("Currency code (USD, EUR, GBP, etc.)"),
            decimals: z
              .number()
              .optional()
              .default(2)
              .describe("Number of decimal places (default: 2)"),
          }),
          execute: async ({ range, currency, decimals = 2 }) => {
            return {
              message: `Formatting ${range} as ${currency} currency...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "format_currency",
                params: { range, currency, decimals },
              },
            };
          },
        }),

        format_currency_range: tool({
          description: "Format exact cell ranges with currency formatting.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name containing the cells"),
            ranges: z
              .array(z.string())
              .min(1)
              .describe("Array of exact cell ranges to format (e.g., ['C15', 'D15:D20'])"),
            currency: z
              .string()
              .default("USD")
              .describe("Currency code (USD, EUR, GBP, JPY, etc.)"),
            decimals: z
              .number()
              .default(2)
              .describe("Number of decimal places"),
          }),
          execute: async ({ sheetName, ranges, currency, decimals }) => {
            return {
              message: `Formatting ranges ${ranges.join(', ')} as ${currency} in sheet '${sheetName}'`,
              clientSideAction: {
                type: "formatCurrencyRanges",
                params: { sheetName, ranges, currency, decimals },
              },
            };
          },
        }),


        add_filter: tool({
          description: "Add filter dropdowns to an exact range. Creates Excel-like filter controls.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name where to add filters"),
            range: z
              .string()
              .describe("Exact range including headers (e.g., 'A1:E20')"),
          }),
          execute: async ({ sheetName, range }) => {
            return {
              message: `Adding filters to range ${range} in sheet '${sheetName}'`,
              clientSideAction: {
                type: "addFilter",
                params: { sheetName, range },
              },
            };
          },
        }),

        apply_filter: tool({
          description: "Apply filter criteria to a specific column in an exact range.",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Exact sheet name containing the data"),
            range: z
              .string()
              .describe("Exact range with headers (e.g., 'A1:E20')"),
            column: z
              .string()
              .describe("Exact column letter to filter (e.g., 'C')"),
            condition: z
              .enum([
                "greater_than",
                "less_than",
                "equals",
                "not_equals",
                "contains",
                "not_contains",
              ])
              .describe("Filter condition type"),
            value: z
              .union([z.string(), z.number()])
              .describe("Value to compare against"),
          }),
          execute: async ({ sheetName, range, column, condition, value }) => {
            return {
              message: `Applying filter to column ${column} in range ${range} (${condition}: ${value})`,
              clientSideAction: {
                type: "applyFilter",
                params: { sheetName, range, column, condition, value },
              },
            };
          },
        }),

        sort_table: tool({
          description:
            "Sort table data by specified column in ascending or descending order. Works with both column names and column letters.",
          parameters: z.object({
            tableId: z
              .string()
              .optional()
              .describe(
                "Optional table ID to target. If not specified, sorts the primary table."
              ),
            column: z
              .union([z.string(), z.number()])
              .describe(
                "Column to sort by - can be column letter (e.g., 'E'), column name (e.g., 'Price'), or column index (0-based)"
              ),
            ascending: z
              .boolean()
              .optional()
              .default(true)
              .describe(
                "Sort order - true for ascending, false for descending"
              ),
            range: z
              .string()
              .optional()
              .describe(
                "Optional explicit range to sort (e.g., 'C8:K12'). If not provided, uses table range."
              ),
          }),
          execute: async ({ tableId, column, ascending = true, range }) => {
            return {
              message: `Sorting ${
                range || `table ${tableId || "primary"}`
              } by column ${column} in ${
                ascending ? "ascending" : "descending"
              } order...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "sort_table",
                params: { tableId, column, ascending, range },
              },
            };
          },
        }),

        format_cells: tool({
          description:
            "Apply text and number formatting to cells (bold, colors, fonts, date formats, etc.) - NOT for currency formatting",
          parameters: z.object({
            range: z
              .string()
              .describe("Cell range (e.g., 'A1:B1', 'A1', 'A1:A10')"),
            bold: z.boolean().optional().describe("Make text bold"),
            italic: z.boolean().optional().describe("Make text italic"),
            fontSize: z.number().optional().describe("Font size"),
            fontColor: z
              .string()
              .optional()
              .describe("Font color (hex, e.g., '#FF0000')"),
            backgroundColor: z
              .string()
              .optional()
              .describe("Background color (hex, e.g., '#FFFF00')"),
            underline: z.boolean().optional().describe("Underline text"),
            numberFormat: z
              .string()
              .optional()
              .describe(
                "Number format pattern (e.g., 'MM/DD/YYYY' for dates, '0.00' for decimals)"
              ),
            textAlign: z
              .enum(["left", "center", "right"])
              .optional()
              .describe("Horizontal text alignment - left, center, or right"),
            verticalAlign: z
              .enum(["top", "middle", "bottom"])
              .optional()
              .describe("Vertical text alignment - top, middle, or bottom"),
            textRotation: z
              .number()
              .optional()
              .describe("Text rotation angle in degrees (e.g., 45, 90, -30)"),
            textWrap: z
              .enum(["overflow", "truncate", "wrap"])
              .optional()
              .describe("Text overflow behavior - overflow, truncate, or wrap"),
          }),
          execute: async ({
            range,
            bold,
            italic,
            fontSize,
            fontColor,
            backgroundColor,
            underline,
            numberFormat,
            textAlign,
            verticalAlign,
            textRotation,
            textWrap,
          }) => {
            try {
              console.log(`🔍 format_cells: Formatting ${range}`, {
                bold,
                italic,
                fontSize,
                fontColor,
                backgroundColor,
                underline,
                numberFormat,
                textAlign,
                verticalAlign,
                textRotation,
                textWrap,
              });

              const result = {
                success: true,
                range,
                formatting: {
                  bold,
                  italic,
                  fontSize,
                  fontColor,
                  backgroundColor,
                  underline,
                  numberFormat,
                  textAlign,
                  verticalAlign,
                  textRotation,
                  textWrap,
                },
                message: `Applied formatting to range ${range}${
                  bold ? " (bold)" : ""
                }${italic ? " (italic)" : ""}${
                  fontSize ? ` (size: ${fontSize})` : ""
                }${fontColor ? ` (color: ${fontColor})` : ""}${
                  backgroundColor ? ` (background: ${backgroundColor})` : ""
                }${numberFormat ? ` (format: ${numberFormat})` : ""}${
                  textAlign ? ` (align: ${textAlign})` : ""
                }${verticalAlign ? ` (vertical: ${verticalAlign})` : ""}${
                  textRotation ? ` (rotation: ${textRotation}°)` : ""
                }${textWrap ? ` (wrap: ${textWrap})` : ""}`,
                clientSideAction: {
                  type: "formatCells",
                  range,
                  bold,
                  italic,
                  fontSize,
                  fontColor,
                  backgroundColor,
                  underline,
                  numberFormat,
                  textAlign,
                  verticalAlign,
                  textRotation,
                  textWrap,
                },
              };

              console.log("✅ format_cells: Returning result:", result);
              return result;
            } catch (error) {
              console.error("❌ format_cells: Server-side error:", error);
              return {
                error: "Failed to format cells",
                message:
                  error instanceof Error ? error.message : "Unknown error",
                success: false,
              };
            }
          },
        }),

        conditional_formatting: tool({
          description:
            "Add conditional formatting rules. You MUST provide explicit visual styles when you want colors. Do not assume defaults.",
          parameters: z.object({
            range: z
              .string()
              .describe("Range to apply (e.g., 'A2:A100' or 'B:B')"),
            ruleType: z
              .enum([
                "number_between",
                "number_gt",
                "number_gte",
                "number_lt",
                "number_lte",
                "number_eq",
                "number_neq",
                "text_contains",
                "text_not_contains",
                "text_starts_with",
                "text_ends_with",
                "not_empty",
                "empty",
                "formula",
                "color_scale",
                "data_bar",
                "unique",
                "duplicate",
              ])
              .optional()
              .describe(
                "Type of rule. Defaults to negative values red (number_lt 0)"
              ),
            min: z.number().optional(),
            max: z.number().optional(),
            equals: z.number().optional(),
            contains: z.string().optional(),
            startsWith: z.string().optional(),
            endsWith: z.string().optional(),
            formula: z.string().optional(),
            // The format object is passed through to the client tool. When specifying visual intent, set either backgroundColor or fontColor.
            format: z
              .object({
                backgroundColor: z.string().optional(),
                fontColor: z.string().optional(),
                bold: z.boolean().optional(),
                italic: z.boolean().optional(),
              })
              .optional(),
          }),
          execute: async (params) => {
            return {
              message: `Applying conditional formatting to ${params.range}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "conditional_formatting",
                params,
              },
            };
          },
        }),

        switch_sheet: tool({
          description:
            "Switch to a different sheet in the workbook or analyze a specific sheet",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Name of the sheet to switch to or analyze"),
            action: z
              .enum(["switch", "analyze"])
              .optional()
              .describe(
                "Whether to switch to the sheet or just analyze it (default: analyze)"
              ),
          }),
          execute: async ({ sheetName, action = "analyze" }) => {
            return {
              message: `${
                action === "switch" ? "Switching to" : "Analyzing"
              } sheet "${sheetName}"...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "switch_sheet",
                params: { sheetName, action },
              },
            };
          },
        }),

        create_sheet: tool({
          description: "Create a new worksheet in the workbook",
          parameters: z.object({
            sheetName: z.string().describe("Name of the new sheet to create"),
            switchToSheet: z
              .boolean()
              .optional()
              .default(true)
              .describe(
                "Whether to switch to the new sheet after creating it (default: true)"
              ),
          }),
          execute: async ({ sheetName, switchToSheet = true }) => {
            return {
              message: `Creating sheet "${sheetName}"${
                switchToSheet ? " and switching to it" : ""
              }...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "create_sheet",
                params: { sheetName, switchToSheet },
              },
            };
          },
        }),

        fit_columns: tool({
          description:
            "Fit one or more columns to their content width (like double-clicking column edge in Excel)",
          parameters: z.object({
            columns: z
              .string()
              .describe(
                "Columns to fit: 'A' for single, 'A:C' for range, or comma-separated like 'A,C,E'"
              ),
            rowsSampleLimit: z
              .number()
              .optional()
              .default(1000)
              .describe(
                "Max rows to sample for measuring content (default 1000)"
              ),
          }),
          execute: async ({ columns, rowsSampleLimit = 1000 }) => {
            return {
              message: `Fitting columns ${columns}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "fit_columns",
                params: { columns, rowsSampleLimit },
              },
            };
          },
        }),

        set_cell_formula: tool({
          description:
            "Set an A1-style formula into a specific cell. Prefer this over writing computed constants.",
          parameters: z.object({
            cell: z.string().describe("Target cell in A1 notation, e.g., E3"),
            formula: z
              .string()
              .describe(
                "Formula string in A1 notation, without or with leading = (both accepted)"
              ),
          }),
          execute: async ({ cell, formula }) => {
            return {
              message: `Setting formula in ${cell}`,
              clientSideAction: {
                type: "setCellValue",
                cell,
                value: formula,
                formula: true,
              },
            };
          },
        }),

        find_cell: tool({
          description:
            "Find the first cell that matches a given text, optionally within a sheet or column. Returns the cell address for building formulas.",
          parameters: z.object({
            text: z.string().describe("Text to match in cell value"),
            sheetName: z.string().optional(),
            match: z.enum(["exact", "contains"]).optional().default("exact"),
            column: z
              .string()
              .optional()
              .describe("Optional column letter to restrict search, e.g., 'A'"),
          }),
          execute: async ({ text, sheetName, match = "exact", column }) => {
            return {
              message: `Locating cell for "${text}"...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "find_cell",
                params: { text, sheetName, match, column },
              },
            };
          },
        }),

        format_as_table: tool({
          description:
            "Format a given A1 range as a Univer Table with optional name and theme.",
          parameters: z.object({
            range: z
              .string()
              .describe("A1 range to convert to a table, e.g., B2:F11"),
            name: z.string().optional().describe("Optional table name"),
            tableId: z
              .string()
              .optional()
              .describe("Optional table id; will be generated if omitted"),
            showHeader: z.boolean().optional().default(true),
            theme: z
              .string()
              .optional()
              .describe(
                "Optional theme name (must exist). If omitted, default theme is used."
              ),
          }),
          execute: async ({
            range,
            name,
            tableId,
            showHeader = true,
            theme,
          }) => {
            return {
              message: `Formatting ${range} as table...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "format_as_table",
                params: { range, name, tableId, showHeader, theme },
              },
            };
          },
        }),
      },
      maxSteps: 5,
    });

    return result.toDataStreamResponse({
      getErrorMessage: (error) => {
        console.error("Stream error:", error);
        return error instanceof Error ? error.message : "An error occurred";
      },
    });
  } catch (error) {
    console.error("Chat API error:", error);

    // More specific error handling
    if (error instanceof Error) {
      return new Response(
        JSON.stringify({
          error: error.message,
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    return new Response(
      JSON.stringify({
        error: "An unexpected error occurred",
      }),
      {
        status: 500,
        headers: { "Content-Type": "application/json" },
      }
    );
  }
}
