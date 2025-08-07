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

    const { messages, workbookData } = await req.json();

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

    // Build sheet context from workbook data sent from frontend
    let sheetContextMessage = "";
    if (workbookData && workbookData.sheets && workbookData.sheets.length > 0) {
      const activeSheet = workbookData.sheets.find((s: any) => s.isActive) || workbookData.sheets[0];
      
      if (activeSheet && activeSheet.headers && activeSheet.headers.length > 0) {
        sheetContextMessage = `

CURRENT SHEET CONTEXT:
- Spreadsheet: "${activeSheet.name}"
- Columns: ${activeSheet.headers.join(", ")} (${activeSheet.headers.length} total)
- Data Range: A1:${String.fromCharCode(65 + activeSheet.headers.length - 1)}${activeSheet.structure?.dataRows || 50}
- Summary: ${activeSheet.structure?.totalCells || 0} cells with data

Based on the available data, you can help users with:
- Column analysis: The spreadsheet has ${activeSheet.headers.length} columns: ${activeSheet.headers.join(", ")}
- Calculations: Available columns for operations
- Data insights: Structured data ready for analysis`;

        console.log("✅ Generated sheet context from workbook data:", {
          sheetName: activeSheet.name,
          headers: activeSheet.headers,
          totalCells: activeSheet.structure?.totalCells,
        });
      }
    } else {
      console.log("⚠️ No workbook data received from frontend");
    }


    const result = streamText({
      model: openai("gpt-4o-mini"),
      system: `You are a professional spreadsheet analysis assistant with access to powerful tools for data manipulation and visualization.

INSTRUCTIONS:
- You have direct access to comprehensive spreadsheet data from ALL sheets in the workbook
- The system provides detailed structural analysis of each sheet including headers, data regions, and numeric columns
- Use the appropriate tool based on the user's request and available data structure
- For multi-sheet workbooks, use switch_sheet to analyze or switch between sheets
- For complex financial statements, look for data regions and header patterns rather than assuming simple tabular structure
- For basic data queries, use list_columns to understand the current sheet structure
- For calculations, use calculate_total for simple sums or financial_intelligence for complex operations
- For data analysis, use create_pivot_table to group and aggregate data
- For visualizations, use generate_chart to create charts from data
- Always consider the comprehensive sheet analysis provided in the context

MULTI-STEP EXECUTION:
- You can call multiple tools in sequence to complete complex requests
- For "analyze this data", use: list_columns → create_pivot_table → generate_chart → summary
- For comprehensive analysis, chain tools logically
- Always provide a final summary of all actions taken

INTELLIGENT RESPONSES:
- When users ask vague questions like "list columns" or "do a sum", use the comprehensive sheet context to provide specific, helpful responses
- For complex financial statements, recognize that data may not be in simple tabular format - look for data regions and header patterns
- If someone says "do a sum" and there's only one numeric column, assume they want that column summed
- For multi-sheet workbooks, proactively mention other available sheets and offer to analyze them
- Be proactive: if you see time-series data, suggest trends; if you see multiple categories, suggest grouping
- For financial statements, recognize common patterns like Income Statements, Balance Sheets, Cash Flow statements
- Always call the appropriate tool to get actual data - never guess or fabricate numbers
- When encountering non-standard layouts, use the data regions and structural analysis to understand the layout

TOOL SELECTION GUIDE:
- "list columns" or "show columns" → Use list_columns tool
- "sum column X" or "total column X" or just "do a sum" → Use calculate_total tool (automatically places sum in spreadsheet)
- "create pivot table" or "group by X" → Use create_pivot_table tool (writes pivot table to spreadsheet)
- "create chart" or "generate chart" → Use generate_chart tool
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
- For any request involving currency symbols ($, €, £, ¥), currency codes (USD, EUR, GBP), or terms like "format as currency", "format as dollars", "add currency symbol" - ALWAYS use format_currency tool
- The format_currency tool applies proper Excel-style number formatting with currency symbols
- Never use format_cells for currency formatting - it handles text formatting (bold, colors, fonts) and number/date formatting
- For date formatting, use format_cells with numberFormat parameter (e.g., "MM/DD/YYYY", "DD-MM-YYYY")

IMPORTANT: Tools now automatically write results to the spreadsheet:
- calculate_total places the sum in the next available cell in the same column
- create_pivot_table writes a formatted pivot table starting at the specified destination
- generate_chart creates native Univer charts embedded directly in the spreadsheet
- set_cell_value allows precise cell placement
- format_cells applies formatting like bold, italic, colors to specified ranges
- switch_sheet allows analyzing or switching between multiple sheets in the workbook

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
                params: {}
              }
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
              console.log(
                `🔍 set_cell_value: Setting ${cell} = ${value} (formula: ${formula})`
              );

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
              .describe("Destination range (e.g., 'D1')"),
          }),
          execute: async ({ groupBy, valueColumn, aggFunc, destination }) => {
            return {
              message: `Creating pivot table grouping by '${groupBy}' with ${aggFunc} of '${valueColumn}'...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "create_pivot_table",
                params: { groupBy, valueColumn, aggFunc, destination }
              }
            };
          },
        }),

        calculate_total: tool({
          description: "Calculate total for a specific column",
          parameters: z.object({
            column: z
              .string()
              .describe("Column name or letter (e.g., 'Sales' or 'B')"),
          }),
          execute: async ({ column }) => {
            return {
              message: `Calculating total for column ${column}...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "calculate_total",
                params: { column }
              }
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
          execute: async ({ data_range, chart_type, title, position, width = 400, height = 300 }) => {
            return {
              message: `Creating ${chart_type} chart "${title}"...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "generate_chart", 
                params: { data_range, chart_type, title, position, width, height }
              }
            };
          },
        }),

        format_currency: tool({
          description: "Format cells as currency (USD, EUR, GBP, etc.) with proper currency symbol",
          parameters: z.object({
            range: z
              .string()
              .optional()
              .describe("Cell range to format (e.g., 'F2:F4', 'A1:A10'). If not provided, will attempt to format the most recently calculated total."),
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
                params: { range, currency, decimals }
              }
            };
          },
        }),

        format_cells: tool({
          description: "Apply text and number formatting to cells (bold, colors, fonts, date formats, etc.) - NOT for currency formatting",
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
              .describe("Number format pattern (e.g., 'MM/DD/YYYY' for dates, '0.00' for decimals)"),
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
                }${verticalAlign ? ` (vertical: ${verticalAlign})` : ""
                }${textRotation ? ` (rotation: ${textRotation}°)` : ""
                }${textWrap ? ` (wrap: ${textWrap})` : ""
                }`,
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
              message: `${action === "switch" ? "Switching to" : "Analyzing"} sheet "${sheetName}"...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "switch_sheet",
                params: { sheetName, action }
              }
            };
          },
        }),

        add_filter: tool({
          description: "Add filter to spreadsheet data to show/hide specific rows based on criteria",
          parameters: z.object({
            range: z
              .string()
              .optional()
              .describe("Data range to filter (e.g., 'A1:G51'). If not provided, will auto-detect data range"),
            column: z
              .string()
              .optional()
              .describe("Specific column to filter (e.g., 'Region', 'Product', or 'C')"),
            filterValues: z
              .array(z.string())
              .optional()
              .describe("Values to show (e.g., ['North', 'South'] to show only North and South regions)"),
            action: z
              .enum(["add", "remove", "clear"])
              .optional()
              .default("add")
              .describe("Action: 'add' creates filter, 'remove' removes specific filter, 'clear' removes all filters"),
          }),
          execute: async ({ range, column, filterValues, action = "add" }) => {
            return {
              message: action === "add" 
                ? `Adding filter${column ? ` on column ${column}` : ''}${filterValues ? ` showing: ${filterValues.join(', ')}` : ''}...`
                : action === "remove" 
                ? `Removing filter${column ? ` on column ${column}` : ''}...`
                : `Clearing all filters...`,
              clientSideAction: {
                type: "executeUniverTool",
                toolName: "add_filter",
                params: { range, column, filterValues, action }
              }
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
