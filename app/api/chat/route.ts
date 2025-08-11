import { openai } from "@ai-sdk/openai";
import { streamText, tool } from "ai";
import { z } from "zod";

export async function POST(req: Request) {
  try {
    const apiKey = process.env.OPENAI_API_KEY;

    if (!apiKey || apiKey === "your_openai_api_key_here") {
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
        .find((a: any) => a?.tool === "smart_add_column" && a?.params);
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
        .find((a: any) => a?.tool === "generate_chart" && a?.params);
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

SMART GUIDANCE:
- ALWAYS use the nextAvailableColumn from table context when adding columns to tables
- When user wants to "add a column" to a table, place it in the nextAvailableColumn, not beyond
- Respect table boundaries: table starts at startRow, not row 0
- For table operations, use the actual table range (e.g., A5:E15), not A1-based ranges
- If selection shows add_column intent, the user wants to extend the selected table
- When multiple tables exist, pick the one matching the user's intent or selection context`;

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
      system: `You are a professional spreadsheet analysis assistant with access to powerful tools for data manipulation and visualization.

🔄 REAL-TIME CONTEXT STRATEGY:
You do NOT have static spreadsheet data. Instead, you have tools to query the current state:
- MANDATORY: Call 'get_workbook_snapshot' FIRST to see complete sheet contents and structure
- Use 'get_active_sheet_context' for detailed analysis of the current sheet  
- Use 'get_selection_context' to understand what the user has selected and their likely intent
- NEVER assume column names or ranges - ALWAYS get the snapshot first to see actual data
- This ensures you have fresh, accurate data about the current state

CRITICAL: Before ANY operation (totals, formatting, charts, etc.), call get_workbook_snapshot to see:
- Exact column headers (e.g., "Price (zł)", "Weight (kg)", "Rating")  
- Data locations and ranges
- Sheet structure and content

🧠 INTELLIGENT ANALYSIS LEVERAGE:
The get_workbook_snapshot tool now provides INTELLIGENT ANALYSIS for each sheet:
- intelligence.tables[]: All detected tables with positions, headers, column types
- intelligence.summary.calculableColumns[]: Pre-identified numeric/currency columns ready for totals
- intelligence.summary.numericColumns[]: All numeric columns across tables
- intelligence.spatialMap[]: Optimal placement zones for new content
- For "add totals" requests, use this intelligence to make smart decisions about which columns to total
- For multi-table sheets, use tableId to target specific tables
- Never guess column names - use the exact names from the intelligence analysis

RUNTIME CONTEXT:
- Server time (ISO): ${serverIso}
- Client local time: ${clientLocalTime}
- Client locale and location hint: ${clientLocale} (${locationHint})
- Approx location (from timezone): ${approxLocation || "unknown"}

INTELLIGENT WORKFLOW:
1. 🔍 ANALYZE FIRST: Start by understanding the current state with context tools
2. 🎯 UNDERSTAND INTENT: Use selection context to understand what the user wants to do
3. 🛠️ CHOOSE TOOLS: Select the most appropriate tools based on actual current data
4. ⚡ EXECUTE SMARTLY: Use spatial awareness to place results in optimal locations
5. ✅ VERIFY RESULTS: Confirm operations completed successfully

CONTEXT-AWARE DECISION MAKING:
- For multi-sheet workbooks, use switch_sheet or get_current_workbook_state first
- For table operations, get_active_sheet_context reveals table boundaries and available space
- For calculations, understand the data structure before choosing sum vs pivot vs chart
- For column operations, use spatial context to find the right location (nextAvailableColumn)
- For VLOOKUP operations, analyze table structures to suggest optimal lookup patterns

MULTI-STEP EXECUTION:
- You can call multiple tools in sequence to complete complex requests
- For "analyze this data", use: list_columns → create_pivot_table → generate_chart → summary
- For comprehensive analysis, chain tools logically
- Always provide a final summary of all actions taken

HEADER FORMATTING POLICY:
- When the user asks to "format headers" or similar, do NOT guess header position.
- First understand the active sheet structure using context tools, or use the explicit header token.
- Prefer calling format_cells with range set to "__AUTO_HEADERS__". The client resolves this to the first row of the primary detected data region using live Univer APIs.
- If the user specifies an explicit range, use it directly instead of the token.

RESPONSE STYLE:
- Be concise. Prefer 1-2 short sentences when confirming actions
- Do NOT paste or preview full datasets/tables in the chat unless the user explicitly asks to "show the data". Insert the data directly into the sheet using tools
- Avoid echoing long lists of rows, CSV, or markdown tables. Summarize instead
 - When you add or modify data, do not list the rows you added. Respond like: "Added 5 rows to Sheet1" or "Updated the total in B19". Only show concrete values on direct request

DATA CREATION POLICY:
- If the user asks to "create" data (e.g., a financial statement, a random dataset, or sample data) and the sheet is empty or missing needed values, you MUST generate plausible numeric values (not only labels) and write them into the sheet immediately.
- Prefer writing entire blocks with bulk_set_values. After writing numbers, add simple derived formulas where helpful (e.g., Gross Profit = Revenue − COGS; Net Cash Flow = SUM of cash flows).
- Do not ask the user to provide data first unless they explicitly demand realism tied to their own figures.

INTELLIGENT RESPONSES:
- When users ask vague questions like "list columns" or "do a sum", use the comprehensive sheet context to provide specific, helpful responses
- For complex financial statements, recognize that data may not be in simple tabular format - look for data regions and header patterns
- If someone says "do a sum" and there's only one numeric column, assume they want that column summed
- For multi-sheet workbooks, proactively mention other available sheets and offer to analyze them
- Be proactive: if you see time-series data, suggest trends; if you see multiple categories, suggest grouping
- For financial statements, recognize common patterns like Income Statements, Balance Sheets, Cash Flow statements
- Always call the appropriate tool to get actual data - never guess or fabricate numbers
- When encountering non-standard layouts, use the data regions and structural analysis to understand the layout

FOLLOW-UP MODIFICATION POLICY (Charts):
- If the user refers to "the chart" or "previous chart", use LAST_CHART above as the base configuration.
- If they ask to restrict to specific columns (e.g., "date and net income"), call generate_chart with x_column and y_columns set precisely to those headers.
- If they ask to "add category" or split the series by a categorical column, first create a pivot table that groups by the X column and splits series by the categorical column, then generate the chart from that pivot range.
- Place follow-up charts near the previous one unless the user specifies a new position.
- Never include extra columns when the user specifies exact columns.

TOOL SELECTION GUIDE:
- 🔄 ALWAYS START: Use get_workbook_snapshot to see complete sheet data and structure first

COLUMN OPERATIONS:
- "add column", "add new column", "create column" → Use smart_add_column tool (automatically finds correct table and position)
- "calculate X per Y", "price per kg", "cost per unit", "X divided by Y" → Use smart_add_column tool with formula pattern (e.g., formulaPattern: "=E{row}/D{row}")
- "add calculated column", "create formula column", "compute X from Y and Z" → Use smart_add_column tool
- "vlookup", "lookup", "find value" → Use intelligent_vlookup tool (context-aware table and column detection)
- "list columns" or "show columns" → Use get_active_sheet_context for full table structure (avoid the legacy list_columns tool)
- When there may be multiple tables, use list_tables first, then pass tableId or data_range to downstream tools

INTELLIGENT FILTERING:
- "add filter", "add filters", "enable filtering", "create filter" → Use add_filter tool (creates Excel-like filter dropdowns in headers)
- The add_filter tool automatically detects table range and applies filters to all columns
- After adding filters, users can interact with the dropdown arrows in headers to filter data

INTELLIGENT TOTALS SELECTION (for bottom-of-table summaries only):
- "add totals", "calculate totals", "sum all", "total everything" → Use add_smart_totals tool (automatically detects and totals ALL calculable columns)
- "sum column X" or "total column X" or "calculate total for [specific column]" → Use calculate_total tool for single column
- "totals for price and weight" → Use add_smart_totals with specific columns array
- The add_smart_totals tool leverages intelligent analysis to identify numeric/currency columns and place totals optimally
- ⚠️ IMPORTANT: These tools add summary rows BELOW the table, NOT new columns with per-row calculations

- "create pivot table" or "group by X" → Use create_pivot_table tool (writes pivot table to spreadsheet)
- "create chart" or "generate chart" → Use generate_chart tool
- When inserting or moving content (pivot tables, charts, pasted ranges), choose a destination in an empty area. Prefer:
  1) the first empty block to the right of the used range on the active sheet,
  2) or below the used range if right side is not spacious enough,
  3) or a new sheet if there is no sufficient space. Use usedRange provided in the context.
- "set cell X to Y" → Use set_cell_value tool
- "format as USD", "format as currency", "add $ symbol" → Use format_currency tool (applies proper currency formatting). For "format the total", omit range parameter to auto-detect recent totals
- "make X bold", "format headers", "center text", "align center", "rotate text", "wrap text" → Use format_cells tool (comprehensive formatting). Supports: textAlign (left/center/right), verticalAlign (top/middle/bottom), textRotation (degrees), textWrap (overflow/truncate/wrap)
- "format dates" or "make dates readable" → Use format_cells tool with numberFormat parameter
- "add filter", "filter by X", "show only X" → Use add_filter tool
- "remove filter", "clear filters" → Use add_filter tool with action parameter
- "switch to sheet X" or "analyze sheet X" → Use switch_sheet tool
- "analyze data" → Use multiple tools in sequence
- Complex financial analysis → Use financial_intelligence tool

IMPORTANT CURRENCY FORMATTING:
- For any request involving currency symbols ($, €, £, ¥), currency codes (USD, EUR, GBP), or terms like "format as currency", "format as dollars", "add currency symbol" - use these tools intelligently:

INTELLIGENT FORMATTING SELECTION:
- "format totals in USD" or "format totals as currency" AFTER adding totals → Use format_recent_totals tool (automatically finds and formats only currency-related totals)
- "format the column as USD" or "format column as currency" → Use format_currency_column tool (automatically detects and formats the Price/Cost/Amount column data range)
- "format cell A1 as USD" or specific range formatting → Use format_currency tool with explicit range
- The format_recent_totals tool is context-aware and will only format price/cost/amount columns from recent totals
- The format_currency_column tool uses intelligent analysis to identify currency columns and formats their data ranges (NOT totals)
- Never use format_cells for currency formatting - it handles text formatting (bold, colors, fonts) and number/date formatting
- For date formatting, use format_cells with numberFormat parameter (e.g., "MM/DD/YYYY", "DD-MM-YYYY")

CRITICAL COLUMN CONTEXT:
- "format the column" means format the data in the most relevant column based on context
- Use workbook intelligence to identify currency columns (Price, Cost, Amount, Revenue, etc.)
- Format the DATA RANGE of the identified column, not empty cells or wrong columns
- Example: "Price (zł)" column at G9:G18 should be formatted, not random empty column B

IMPORTANT: Tools now automatically write results to the spreadsheet:
- calculate_total places the sum in the next available cell in the same column
- create_pivot_table writes a formatted pivot table starting at the specified destination
- generate_chart creates native Univer charts embedded directly in the spreadsheet
- set_cell_value allows precise cell placement
- format_cells applies formatting like bold, italic, colors to specified ranges
- switch_sheet allows analyzing or switching between multiple sheets in the workbook

DISAMBIGUATION POLICY:
- If multiple tables are present and neither tableId nor data_range are provided for an operation, call list_tables and choose; otherwise ask for clarification.

CHART INTELLIGENCE:
- For time-series data (like Date + Search_Volume), automatically suggest line charts
- For categorical data, suggest column or bar charts
- Charts are created natively in Univer (not external URLs)
- Auto-detect appropriate data ranges when not specified

${sheetContextMessage}

Be professional, execute requests immediately, and provide specific insights based on the actual data.`,
      messages,
      temperature: 0.2, // Lower temperature for more deterministic operations
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

        smart_add_column: tool({
          description:
            "Add calculated column to table - intelligently detects table structure and creates per-row formulas. Use for 'calculate X per Y', 'price per kg', ratio calculations, etc.",
          parameters: z.object({
            columnName: z
              .string()
              .describe(
                "Name for the new column header (e.g., 'Price per kg')"
              ),
            formulaPattern: z
              .string()
              .optional()
              .describe(
                "Formula pattern with {row} placeholder (e.g., '=E{row}/D{row}' for Price/Weight). System will auto-detect correct columns if not provided."
              ),
            defaultValue: z
              .union([z.string(), z.number()])
              .optional()
              .describe("Default value if no formula"),
          }),
          execute: async ({ columnName, formulaPattern, defaultValue }) => {
            return {
              message:
                "🧠 Getting REAL sheet context to find the perfect column location...",
              clientSideAction: {
                type: "smartAddColumnWithContext",
                params: { columnName, formulaPattern, defaultValue },
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

        get_current_workbook_state: tool({
          description:
            "Get the REAL current state of workbook - no assumptions, pure intelligence",
          parameters: z.object({
            includeAllSheets: z
              .boolean()
              .optional()
              .default(true)
              .describe("Analyze all sheets or just active"),
            includeSelection: z
              .boolean()
              .optional()
              .default(true)
              .describe("Include current selection analysis"),
          }),
          execute: async ({
            includeAllSheets = true,
            includeSelection = true,
          }) => {
            return {
              message:
                "🔍 Analyzing real workbook state - no hardcoded assumptions...",
              clientSideAction: {
                type: "getCleanWorkbookState",
                params: { includeAllSheets, includeSelection },
              },
            };
          },
        }),

        get_active_sheet_context: tool({
          description:
            "Get REAL active sheet context - actual data boundaries, real table structures, true spatial awareness",
          parameters: z.object({
            analyzeDataRegions: z
              .boolean()
              .optional()
              .default(true)
              .describe("Detect all data regions intelligently"),
            findEmptyAreas: z
              .boolean()
              .optional()
              .default(true)
              .describe("Find actual empty areas for placement"),
          }),
          execute: async ({
            analyzeDataRegions = true,
            findEmptyAreas = true,
          }) => {
            return {
              message:
                "🧠 Analyzing real sheet structure - no hardcoded ranges...",
              clientSideAction: {
                type: "getCleanSheetContext",
                params: { analyzeDataRegions, findEmptyAreas },
              },
            };
          },
        }),

        get_selection_context: tool({
          description:
            "Analyze REAL selection - actual values, true intent, intelligent suggestions based on current context",
          parameters: z.object({
            analyzeIntent: z
              .boolean()
              .optional()
              .default(true)
              .describe("Determine user intent from selection"),
            findRelatedData: z
              .boolean()
              .optional()
              .default(true)
              .describe("Find related data structures"),
          }),
          execute: async ({ analyzeIntent = true, findRelatedData = true }) => {
            return {
              message: "🎯 Analyzing real selection intent - no assumptions...",
              clientSideAction: {
                type: "getCleanSelectionContext",
                params: { analyzeIntent, findRelatedData },
              },
            };
          },
        }),

        intelligent_vlookup: tool({
          description:
            "VLOOKUP with ZERO hardcoding - analyzes real tables, finds actual columns, creates precise formulas",
          parameters: z.object({
            lookupValue: z
              .string()
              .describe("Value to lookup (cell reference or literal)"),
            lookupColumn: z
              .string()
              .optional()
              .describe("Column to search in (will auto-detect best option)"),
            returnColumn: z
              .string()
              .optional()
              .describe("Column to return (will suggest based on data types)"),
            exactMatch: z
              .boolean()
              .optional()
              .default(true)
              .describe("Exact match (TRUE) or approximate (FALSE)"),
          }),
          execute: async ({
            lookupValue,
            lookupColumn,
            returnColumn,
            exactMatch = true,
          }) => {
            return {
              message:
                "🔍 Analyzing REAL table structures for perfect VLOOKUP setup...",
              clientSideAction: {
                type: "intelligentVlookupWithContext",
                params: { lookupValue, lookupColumn, returnColumn, exactMatch },
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
            clearSource: z.boolean().optional().default(true),
          }),
          execute: async ({
            sourceRange,
            destStartCell,
            clearSource = true,
          }) => {
            return {
              success: true,
              clientSideAction: {
                type: "moveRange",
                sourceRange,
                destStartCell,
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
          description: "Create a pivot table with grouping and aggregation",
          parameters: z.object({
            groupBy: z
              .string()
              .describe("Column to group by (e.g., 'Date' or 'Region')"),
            valueColumn: z
              .string()
              .describe(
                "Column to aggregate (e.g., 'Search_Volume' or 'Sales')"
              ),
            aggFunc: z
              .enum(["sum", "average", "count", "max", "min"])
              .describe("Aggregation function"),
            destination: z
              .string()
              .optional()
              .describe("Destination range (e.g., 'D1' or 'Sheet2!A1')"),
            sheetName: z
              .string()
              .optional()
              .describe(
                "Optional sheet name to create/use for the pivot table"
              ),
            data_range: z
              .string()
              .optional()
              .describe("Optional source range like 'H1:I20'"),
            tableId: z
              .string()
              .optional()
              .describe(
                "Stable tableId from list_tables: '<sheetIndex>:<range>'"
              ),
          }),
          execute: async ({
            groupBy,
            valueColumn,
            aggFunc,
            destination,
            sheetName,
            data_range,
            tableId,
          }) => {
            return {
              message: `Creating pivot table grouping by '${groupBy}' with ${aggFunc} of '${valueColumn}'...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "create_pivot_table",
                params: {
                  groupBy,
                  valueColumn,
                  aggFunc,
                  destination,
                  sheetName,
                  data_range,
                  tableId,
                },
              },
            };
          },
        }),

        calculate_total: tool({
          description: "Calculate total for a specific column",
          parameters: z.object({
            column: z
              .string()
              .describe("Column name or letter (e.g., 'Sales' or 'B')"),
            data_range: z
              .string()
              .optional()
              .describe("Optional source range like 'H1:I20'"),
            tableId: z
              .string()
              .optional()
              .describe(
                "Stable tableId from list_tables: '<sheetIndex>:<range>'"
              ),
          }),
          execute: async ({ column, data_range, tableId }) => {
            return {
              message: `Calculating total for column ${column}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "calculate_total",
                params: { column, data_range, tableId },
              },
            };
          },
        }),

        add_smart_totals: tool({
          description: `Intelligently add totals to multiple numeric/calculable columns in a table.

- Automatically detects which columns are calculable (numeric, currency, or have calculable names like "Price", "Weight", etc.)
- Can add totals to all calculable columns at once or specific columns
- Places SUM formulas below the data table in the appropriate columns
- Perfect for "add totals" requests where the user wants totals for all relevant columns

When to use:
- User asks "add totals" without specifying columns
- User wants totals for multiple columns  
- User asks to "calculate totals for all numeric columns"

IMPORTANT: This tool leverages intelligent table analysis from get_workbook_snapshot to automatically identify calculable columns and optimal placement.`,
          parameters: z.object({
            columns: z
              .array(z.string())
              .optional()
              .describe(
                "Optional array of specific column names or letters to total"
              ),
            tableId: z
              .string()
              .optional()
              .describe("Specific table ID if multiple tables exist"),
          }),
          execute: async ({ columns, tableId }) => {
            return {
              message: `Adding smart totals${
                columns
                  ? ` for ${columns.join(", ")}`
                  : " for all calculable columns"
              }...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "add_smart_totals",
                params: { columns, tableId },
              },
            };
          },
        }),

        generate_chart: tool({
          description: "Generate native Univer chart from spreadsheet data",
          parameters: z.object({
            data_range: z
              .string()
              .optional()
              .describe(
                "Data range (e.g., 'A1:B18') - if not provided, will auto-detect data"
              ),
            x_column: z
              .string()
              .optional()
              .describe(
                "Optional X-axis column header (e.g., 'Date'). When provided, restrict chart to this X column."
              ),
            y_columns: z
              .union([z.string(), z.array(z.string())])
              .optional()
              .describe(
                "Optional Y-axis column header(s). When provided, restrict chart to these Y columns."
              ),
            tableId: z
              .string()
              .optional()
              .describe(
                "Stable tableId from list_tables: '<sheetIndex>:<range>'"
              ),
            chart_type: z
              .enum(["column", "line", "pie", "bar", "scatter"])
              .describe("Chart type - column, line, pie, bar, or scatter"),
            title: z.string().describe("Chart title"),
            position: z
              .string()
              .optional()
              .describe("Chart position cell (e.g., 'D1')"),
            width: z
              .number()
              .optional()
              .describe("Chart width in pixels (default: 400)"),
            height: z
              .number()
              .optional()
              .describe("Chart height in pixels (default: 300)"),
          }),
          execute: async ({
            data_range,
            x_column,
            y_columns,
            tableId,
            chart_type,
            title,
            position,
            width = 400,
            height = 300,
          }) => {
            return {
              message: `Creating ${chart_type} chart "${title}"...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "generate_chart",
                params: {
                  data_range,
                  x_column,
                  y_columns,
                  tableId,
                  chart_type,
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

        format_recent_totals: tool({
          description: `Intelligently format recently added totals with currency formatting.

- Automatically detects and formats currency/price-related totals from recent add_smart_totals operations
- Identifies columns like "Price", "Cost", "Amount", "Revenue", "Value" automatically
- Perfect for requests like "format totals in USD" after adding totals

When to use:
- User asks to "format totals" or "format totals in USD" after adding totals
- User wants to apply currency formatting to recent calculations
- User asks to "make the price total show as dollars"

IMPORTANT: This tool understands the context of recently added totals and intelligently selects only currency-relevant ones.`,
          parameters: z.object({
            currency: z
              .string()
              .default("USD")
              .describe("Currency code (USD, EUR, GBP, JPY, etc.)"),
            decimals: z
              .number()
              .default(2)
              .describe("Number of decimal places"),
            columnPattern: z
              .string()
              .optional()
              .describe(
                "Optional pattern to match specific column names (e.g., 'price' to match only price columns)"
              ),
          }),
          execute: async ({ currency, decimals, columnPattern }) => {
            return {
              message: `Formatting recent currency totals as ${currency}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "format_recent_totals",
                params: { currency, decimals, columnPattern },
              },
            };
          },
        }),

        format_currency_column: tool({
          description: `Intelligently format a currency/price column with proper currency formatting.

- Uses workbook intelligence to identify which column contains currency data (Price, Cost, Amount, Revenue, etc.)
- Formats the ENTIRE DATA RANGE of the identified currency column, not just totals
- Perfect for requests like "format the column as USD" or "format the price column as currency"

When to use:
- User asks to "format the column as USD" or "format column as currency"
- User wants to format the data values in a currency column (not just totals)
- User refers to "the column" and context suggests they mean a price/currency column

IMPORTANT: This tool uses intelligent analysis to find the most relevant currency column and formats its data range.`,
          parameters: z.object({
            currency: z
              .string()
              .default("USD")
              .describe("Currency code (USD, EUR, GBP, JPY, etc.)"),
            decimals: z
              .number()
              .default(2)
              .describe("Number of decimal places"),
            columnName: z
              .string()
              .optional()
              .describe(
                "Optional specific column name to format (e.g., 'Price'). If not provided, will auto-detect currency column."
              ),
          }),
          execute: async ({ currency, decimals, columnName }) => {
            return {
              message: `Formatting currency column as ${currency}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "format_currency_column",
                params: { currency, decimals, columnName },
              },
            };
          },
        }),

        add_filter: tool({
          description: `Add filter dropdowns to a table (like Excel AutoFilter). Creates interactive filter controls in the header row.

- Automatically detects the table range using intelligent analysis
- Creates filter dropdowns for each column header
- Users can then click the dropdowns to filter data interactively
- Works with any table structure and position

When to use:
- User asks to "add filter", "add filters", "enable filtering", "create filter"
- User wants to filter data like in Excel
- Follow-up to table creation when filtering is needed

The tool intelligently finds the table and applies filters to the entire table range including headers.`,
          parameters: z.object({
            tableId: z
              .string()
              .optional()
              .describe(
                "Optional table ID to target. If not specified, applies to the primary table detected in the sheet."
              ),
          }),
          execute: async ({ tableId }) => {
            return {
              message: `Adding filter to table${
                tableId ? ` ${tableId}` : ""
              }...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "add_filter",
                params: { tableId },
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
            "Add conditional formatting rules (e.g., negatives red, values between X and Y, text contains)",
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
            background: z
              .string()
              .optional()
              .describe("Background color (e.g., '#FF0000')"),
            fontColor: z
              .string()
              .optional()
              .describe("Font color (e.g., '#00FF00')"),
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
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
          description:
            "Create a new worksheet in the workbook",
          parameters: z.object({
            sheetName: z
              .string()
              .describe("Name of the new sheet to create"),
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
              message: `Creating sheet "${sheetName}"${switchToSheet ? ' and switching to it' : ''}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "create_sheet",
                params: { sheetName, switchToSheet },
              },
            };
          },
        }),

        auto_fit_columns: tool({
          description:
            "Auto-fit one or more columns to their content width (like double-clicking column edge in Excel)",
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
              message: `Auto-fitting columns ${columns}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "auto_fit_columns",
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
