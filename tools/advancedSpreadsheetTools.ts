// tools/advancedSpreadsheetTools.ts
// NOTE: These LangChain tools are unused - all tool execution happens in chat API
// This file can be removed if not needed for future LangChain integration
import { DynamicStructuredTool } from "@langchain/core/tools";
import { z } from "zod";
import { UniverService } from "@/services/univerService";

export const getSheetContextTool = new DynamicStructuredTool({
  name: "get_sheet_context",
  description: "Get current sheet context including data structure, tables, and content summary. Use this to understand what's in the spreadsheet before performing operations.",
  schema: z.object({}),
  func: async () => {
    try {
      const univerService = UniverService.getInstance();
      const context = await univerService.getSheetContext();
      return JSON.stringify(context, null, 2);
    } catch (error) {
      return `Error getting sheet context: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const setCellValueTool = new DynamicStructuredTool({
  name: "set_cell_value",
  description: "Set a value in a specific cell. Use for entering data, text, or numbers.",
  schema: z.object({
    cell: z.string().describe("The cell reference (e.g., 'A1', 'B5')"),
    value: z.union([z.string(), z.number()]).describe("The value to set in the cell"),
  }),
  func: async ({ cell, value }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.setCellValue(cell, value);
    } catch (error) {
      return `Error setting cell value: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const setFormulaTool = new DynamicStructuredTool({
  name: "set_formula",
  description: "Set a formula in a specific cell. Supports all Excel formulas like SUM, AVERAGE, VLOOKUP, etc.",
  schema: z.object({
    cell: z.string().describe("The cell reference where to place the formula (e.g., 'A1', 'B5')"),
    formula: z.string().describe("The formula to set (e.g., 'SUM(A1:A10)', 'AVERAGE(B:B)', '=A1*B1')"),
  }),
  func: async ({ cell, formula }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.setFormula(cell, formula);
    } catch (error) {
      return `Error setting formula: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const formatCellsTool = new DynamicStructuredTool({
  name: "format_cells",
  description: "Format cells with styling like bold, italic, font size, colors, etc.",
  schema: z.object({
    range: z.string().describe("The range to format (e.g., 'A1:B5', 'A:A', '1:1')"),
    bold: z.boolean().optional().describe("Make text bold"),
    italic: z.boolean().optional().describe("Make text italic"),
    fontSize: z.number().optional().describe("Font size in points"),
    fontColor: z.string().optional().describe("Font color (name or hex)"),
    backgroundColor: z.string().optional().describe("Background color (name or hex)"),
  }),
  func: async ({ range, bold, italic, fontSize, fontColor, backgroundColor }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.formatCells(range, {
        bold,
        italic,
        fontSize,
        fontColor,
        backgroundColor,
      });
    } catch (error) {
      return `Error formatting cells: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

// WARNING: Row/column operations may not work with current Univer API
export const insertRowTool = new DynamicStructuredTool({
  name: "insert_row",
  description: "Insert a new row at the specified index - WARNING: May not be supported",
  schema: z.object({
    rowIndex: z.number().describe("The row index where to insert (0-based)"),
  }),
  func: async ({ rowIndex }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.insertRow(rowIndex);
    } catch (error) {
      return `Error inserting row: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const deleteRowTool = new DynamicStructuredTool({
  name: "delete_row",
  description: "Delete a row at the specified index",
  schema: z.object({
    rowIndex: z.number().describe("The row index to delete (0-based)"),
  }),
  func: async ({ rowIndex }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.deleteRow(rowIndex);
    } catch (error) {
      return `Error deleting row: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const insertColumnTool = new DynamicStructuredTool({
  name: "insert_column",
  description: "Insert a new column at the specified index",
  schema: z.object({
    colIndex: z.number().describe("The column index where to insert (0-based)"),
  }),
  func: async ({ colIndex }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.insertColumn(colIndex);
    } catch (error) {
      return `Error inserting column: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const deleteColumnTool = new DynamicStructuredTool({
  name: "delete_column",
  description: "Delete a column at the specified index",
  schema: z.object({
    colIndex: z.number().describe("The column index to delete (0-based)"),
  }),
  func: async ({ colIndex }) => {
    try {
      const univerService = UniverService.getInstance();
      return await univerService.deleteColumn(colIndex);
    } catch (error) {
      return `Error deleting column: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});

export const smartColorTool = new DynamicStructuredTool({
  name: "smart_color_cells",
  description: "Intelligently color cells based on context. Can understand descriptions like 'table', 'headers', or 'data' and automatically detect the appropriate range.",
  schema: z.object({
    target: z.string().describe("Description of what to color (e.g., 'table', 'headers', 'first row', 'all data', or specific range like 'A1:B5')"),
    color: z.string().describe("Color name (e.g., 'blue', 'red') or hex code"),
  }),
  func: async ({ target, color }) => {
    try {
      const univerService = UniverService.getInstance();
      const context = await univerService.getSheetContext();
      return await univerService.smartColorCells(target, color, context);
    } catch (error) {
      return `Error coloring cells: ${error instanceof Error ? error.message : "Unknown error"}`;
    }
  },
});