/**
 * Additional Modern Tools
 *
 * Modern implementations for remaining legacy tools and new functionality.
 */

import { createSimpleTool } from "../tool-executor";
import type { UniversalToolContext } from "../universal-context";
import { AGGREGATED_CHART_TOOLS } from "./aggregated-chart-tools";
import { EXCEL_FUNCTION_TOOLS } from "./excel-function-tools";

/**
 * AUTO FIT COLUMNS - Modern Implementation
 * Migrated from: executeAutoFitColumns
 */
export const AutoFitColumnsTool = createSimpleTool(
  {
    name: "auto_fit_columns",
    description:
      "Auto-fit column widths to their content (like double-clicking column edge in Excel)",
    category: "structure",
    requiredContext: ["tables"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      columns?: string | string[];
      tableId?: string;
      allColumns?: boolean;
    }
  ) => {
    const { columns, tableId, allColumns = false } = params;

    let columnsToFit: string[] = [];

    if (allColumns) {
      // Auto-fit all columns in the table
      const table = context.findTable(tableId);
      if (!table) {
        throw new Error(`Table ${tableId || "primary"} not found`);
      }
      columnsToFit = table.columns.map((col: any) => col.letter);
    } else if (columns) {
      // Parse columns parameter
      if (typeof columns === "string") {
        columnsToFit = columns.includes(",")
          ? columns.split(",").map((c) => c.trim())
          : [columns];
      } else {
        columnsToFit = columns;
      }
    } else {
      throw new Error(
        "Must specify columns to auto-fit or set allColumns=true"
      );
    }

    const fittedColumns = [];

    // Auto-fit each column
    for (const column of columnsToFit) {
      try {
        // Try different auto-fit methods
        if (typeof context.fWorksheet.autoFitColumn === "function") {
          const colIndex = column.charCodeAt(0) - 65;
          context.fWorksheet.autoFitColumn(colIndex);
        } else if (typeof context.fWorksheet.autoResizeColumn === "function") {
          const colIndex = column.charCodeAt(0) - 65;
          context.fWorksheet.autoResizeColumn(colIndex);
        } else {
          // Fallback: set a reasonable default width
          const colRange = context.fWorksheet.getRange(`${column}:${column}`);
          if (typeof colRange.setColumnWidth === "function") {
            colRange.setColumnWidth(100); // Default width
          }
        }
        fittedColumns.push(column);
      } catch (error) {
        console.warn(`⚠️ Failed to auto-fit column ${column}:`, error);
      }
    }

    return {
      fittedColumns,
      totalCount: fittedColumns.length,
      message: `Auto-fitted ${
        fittedColumns.length
      } columns: ${fittedColumns.join(", ")}`,
    };
  }
);

/**
 * FIND CELL - Modern Implementation
 * Migrated from: executeFindCell
 */
export const FindCellTool = createSimpleTool(
  {
    name: "find_cell",
    description: "Find cells containing specific values or patterns",
    category: "navigation",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      searchValue: string;
      matchCase?: boolean;
      wholeWord?: boolean;
      useRegex?: boolean;
      searchRange?: string;
      tableId?: string;
    }
  ) => {
    const {
      searchValue,
      matchCase = false,
      wholeWord = false,
      useRegex = false,
      searchRange,
      tableId,
    } = params;

    if (!searchValue) {
      throw new Error("Search value is required");
    }

    let targetRange = searchRange;
    if (!targetRange && tableId) {
      const table = context.findTable(tableId);
      if (table) {
        targetRange = table.range;
      }
    }

    const foundCells = [];
    const cellData = context.activeSheetSnapshot.cellData || {};

    // Determine search criteria
    let searchPattern: RegExp | string = searchValue;
    if (useRegex) {
      try {
        searchPattern = new RegExp(searchValue, matchCase ? "g" : "gi");
      } catch (regexError) {
        throw new Error(`Invalid regex pattern: ${searchValue}`);
      }
    } else if (!matchCase) {
      searchPattern = searchValue.toLowerCase();
    }

    // Parse search range if provided
    let startRow = 0,
      endRow = 999,
      startCol = 0,
      endCol = 25; // Default range
    if (targetRange) {
      const rangeMatch = targetRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (rangeMatch) {
        const [, startColStr, startRowStr, endColStr, endRowStr] = rangeMatch;
        startRow = parseInt(startRowStr) - 1;
        endRow = parseInt(endRowStr) - 1;
        startCol =
          startColStr
            .split("")
            .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
        endCol =
          endColStr
            .split("")
            .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
      }
    }

    // Search through cells
    for (let row = startRow; row <= endRow; row++) {
      const rowData = cellData[row] || {};
      for (let col = startCol; col <= endCol; col++) {
        const cell = rowData[col];
        if (!cell || !cell.v) continue;

        const cellValue = cell.v.toString();
        let isMatch = false;

        if (useRegex && searchPattern instanceof RegExp) {
          isMatch = searchPattern.test(cellValue);
        } else if (wholeWord) {
          const exactMatch = matchCase
            ? cellValue === searchValue
            : cellValue.toLowerCase() === searchPattern;
          isMatch = exactMatch;
        } else {
          const searchText = matchCase ? cellValue : cellValue.toLowerCase();
          isMatch = searchText.includes(searchPattern as string);
        }

        if (isMatch) {
          const colLetter = String.fromCharCode(65 + col);
          const cellAddress = `${colLetter}${row + 1}`;
          foundCells.push({
            address: cellAddress,
            value: cellValue,
            row: row + 1,
            column: colLetter,
          });
        }
      }
    }

    return {
      searchValue,
      searchRange: targetRange || "entire sheet",
      foundCells,
      totalFound: foundCells.length,
      searchOptions: {
        matchCase,
        wholeWord,
        useRegex,
      },
      message:
        foundCells.length > 0
          ? `Found ${
              foundCells.length
            } cells containing '${searchValue}': ${foundCells
              .map((c) => c.address)
              .join(", ")}`
          : `No cells found containing '${searchValue}'`,
    };
  }
);

/**
 * FORMAT AS TABLE - Modern Implementation
 * Migrated from: executeFormatAsTable
 */
export const FormatAsTableTool = createSimpleTool(
  {
    name: "format_as_table",
    description:
      "Format a range as a table with borders, alternating row colors, and styling",
    category: "format",
    requiredContext: ["tables"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      tableId?: string;
      range?: string;
      style?: "light" | "medium" | "dark";
      alternatingRows?: boolean;
      headerRow?: boolean;
      borders?: boolean;
    }
  ) => {
    const {
      tableId,
      range,
      style = "medium",
      alternatingRows = true,
      headerRow = true,
      borders = true,
    } = params;

    let targetRange = range;
    if (!targetRange && tableId) {
      const table = context.findTable(tableId);
      if (table) {
        targetRange = table.range;
      }
    }

    if (!targetRange) {
      throw new Error("Either tableId or range must be provided");
    }

    // Parse the range
    const rangeMatch = targetRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${targetRange}`);
    }

    const [, startCol, startRow, endCol, endRow] = rangeMatch;
    const startRowIndex = parseInt(startRow) - 1;
    const endRowIndex = parseInt(endRow) - 1;
    const startColIndex =
      startCol
        .split("")
        .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
    const endColIndex =
      endCol
        .split("")
        .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;

    const appliedFormats = [];

    // Color schemes for different table styles
    const colorSchemes = {
      light: {
        headerBg: "#F2F2F2",
        headerText: "#000000",
        alternateRowBg: "#F9F9F9",
        borderColor: "#D0D0D0",
      },
      medium: {
        headerBg: "#4472C4",
        headerText: "#FFFFFF",
        alternateRowBg: "#F2F2F2",
        borderColor: "#A6A6A6",
      },
      dark: {
        headerBg: "#2F4F4F",
        headerText: "#FFFFFF",
        alternateRowBg: "#E8E8E8",
        borderColor: "#808080",
      },
    };

    const colors = colorSchemes[style];

    // Format header row
    if (headerRow) {
      const headerRange = `${startCol}${startRow}:${endCol}${startRow}`;
      const headerRangeObj = context.fWorksheet.getRange(headerRange);
      headerRangeObj.setBackgroundColor(colors.headerBg);
      headerRangeObj.setFontColor(colors.headerText);
      headerRangeObj.setBold(true);
      appliedFormats.push("header formatting");
    }

    // Apply alternating row colors
    if (alternatingRows) {
      const dataStartRow = headerRow ? startRowIndex + 1 : startRowIndex;
      for (let row = dataStartRow; row <= endRowIndex; row++) {
        if ((row - dataStartRow) % 2 === 1) {
          // Every other row
          const rowRange = `${startCol}${row + 1}:${endCol}${row + 1}`;
          const rowRangeObj = context.fWorksheet.getRange(rowRange);
          rowRangeObj.setBackgroundColor(colors.alternateRowBg);
        }
      }
      appliedFormats.push("alternating row colors");
    }

    // Apply borders
    if (borders) {
      const tableRangeObj = context.fWorksheet.getRange(targetRange);
      // Note: Border methods may vary by Univer version
      if (typeof tableRangeObj.setBorder === "function") {
        tableRangeObj.setBorder("all", colors.borderColor, "thin");
      } else if (typeof tableRangeObj.setBorders === "function") {
        tableRangeObj.setBorders(
          true,
          true,
          true,
          true,
          true,
          true,
          colors.borderColor,
          "thin"
        );
      }
      appliedFormats.push("borders");
    }

    return {
      range: targetRange,
      style,
      appliedFormats,
      options: {
        alternatingRows,
        headerRow,
        borders,
      },
      message: `Formatted ${targetRange} as ${style} table with: ${appliedFormats.join(
        ", "
      )}`,
    };
  }
);

/**
 * CONDITIONAL FORMATTING - Modern Implementation using proper Univer.js CF API
 * Fixed to use newConditionalFormattingRule() instead of direct cell formatting
 */
export const ConditionalFormattingTool = createSimpleTool(
  {
    name: "conditional_formatting",
    description:
      "Apply conditional formatting rules to highlight cells based on their values",
    category: "format",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      range: string;
      condition:
        | "greater_than"
        | "less_than"
        | "equal_to"
        | "between"
        | "contains"
        | "not_empty"
        | "empty";
      value?: number | string;
      value2?: number; // For 'between' condition
      format: {
        backgroundColor?: string;
        fontColor?: string;
        bold?: boolean;
        italic?: boolean;
      };
    }
  ) => {
    const { range, condition, value, value2, format } = params;

    if (!range) {
      throw new Error("Range is required for conditional formatting");
    }

    console.log(`🎨 Creating conditional formatting rule for range: ${range}`);
    console.log(`🎨 Condition: ${condition}, Value: ${value}, Format:`, format);

    // Parse range to get coordinates for conditional formatting
      const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        throw new Error(`Invalid range format: ${range}`);
      }

      const [, startCol, startRow, endCol, endRow] = rangeMatch;
      const startRowIndex = parseInt(startRow) - 1;
      const endRowIndex = parseInt(endRow) - 1;
      const startColIndex =
        startCol
          .split("")
          .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
      const endColIndex =
        endCol
          .split("")
          .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;

      // Create a new conditional formatting rule using Univer.js CF API
      let rule = context.fWorksheet.newConditionalFormattingRule();

      // Set the range for the rule using FRange object
      const fRange = context.fWorksheet.getRange(range);
      rule = rule.setRanges([fRange.getRange()]);

      // Apply the condition based on type
      switch (condition) {
        case "greater_than":
          if (typeof value === "number") {
            rule = rule.whenNumberGreaterThan(value);
          } else {
            throw new Error("greater_than condition requires a numeric value");
          }
          break;
        case "less_than":
          if (typeof value === "number") {
            rule = rule.whenNumberLessThan(value);
          } else {
            throw new Error("less_than condition requires a numeric value");
          }
          break;
        case "equal_to":
          if (typeof value === "number") {
            rule = rule.whenNumberEqualTo(value);
          } else if (typeof value === "string") {
            rule = rule.whenTextEqualTo(value);
          } else {
            throw new Error("equal_to condition requires a value");
          }
          break;
        case "between":
          if (typeof value === "number" && typeof value2 === "number") {
            rule = rule.whenNumberBetween(
              Math.min(value, value2),
              Math.max(value, value2)
            );
          } else {
            throw new Error("between condition requires two numeric values");
          }
          break;
        case "contains":
          if (typeof value === "string") {
            rule = rule.whenTextContains(value);
          } else {
            throw new Error("contains condition requires a string value");
          }
          break;
        case "not_empty":
          rule = rule.whenCellNotEmpty();
          break;
        case "empty":
          rule = rule.whenCellEmpty();
          break;
        default:
          throw new Error(`Unsupported condition: ${condition}`);
      }

      // Apply formatting to the rule
      if (format.backgroundColor) {
        rule = rule.setBackground(format.backgroundColor);
        console.log(`🎨 Applied background color: ${format.backgroundColor}`);
      }
      if (format.fontColor) {
        rule = rule.setFontColor(format.fontColor);
        console.log(`🎨 Applied font color: ${format.fontColor}`);
      }
      if (format.bold) {
        rule = rule.setBold(true);
        console.log(`🎨 Applied bold formatting`);
      }
      if (format.italic) {
        rule = rule.setItalic(true);
        console.log(`🎨 Applied italic formatting`);
      }

      // Build and apply the rule
      const builtRule = rule.build();
      console.log(`🎨 Built conditional formatting rule:`, builtRule);

      // Apply the rule to the worksheet
      const ruleId = context.fWorksheet.addConditionalFormattingRule(builtRule);
      console.log(`🎨 Applied conditional formatting rule with ID: ${ruleId}`);

      return {
        success: true,
        range,
        condition,
        conditionValue: value,
        conditionValue2: value2,
        format,
        ruleId,
        message: `Applied conditional formatting rule to ${range} based on '${condition}' condition using Univer.js CF API`,
      };
  }
);

/**
 * FIND & REPLACE - Modern Implementation
 */
export const FindReplaceTool = createSimpleTool(
  {
    name: "find_replace",
    description:
      "Find and replace text within a sheet range or table. Supports case sensitivity, whole-word, and regex.",
    category: "data",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      findText: string;
      replaceText: string;
      matchCase?: boolean;
      wholeWord?: boolean;
      useRegex?: boolean;
      range?: string; // A1 style like 'A1:D20', 'A:A', etc.
      tableId?: string; // optional table scope
      sheetName?: string; // reserved; current active sheet used
      selectionOnly?: boolean; // reserved for future selection integration
    }
  ) => {
    const {
      findText,
      replaceText,
      matchCase = false,
      wholeWord = false,
      useRegex = false,
      range,
      tableId,
    } = params;

    if (!findText || typeof findText !== "string") {
      throw new Error("findText is required");
    }

    // Determine target range
    let targetRange = range;
    if (!targetRange && tableId) {
      const table = context.findTable(tableId);
      if (!table) throw new Error(`Table ${tableId} not found`);
      targetRange = table.range;
    }

    // If no explicit range, derive a minimal used range from snapshot
    if (!targetRange) {
      const cellData = context.activeSheetSnapshot?.cellData || {};
      let minRow = Number.MAX_SAFE_INTEGER,
        maxRow = -1,
        minCol = Number.MAX_SAFE_INTEGER,
        maxCol = -1;
      const toA1Col = (n: number) => {
        let s = "";
        n += 1;
        while (n > 0) {
          const m = (n - 1) % 26;
          s = String.fromCharCode(65 + m) + s;
          n = Math.floor((n - 1) / 26);
        }
        return s;
      };
      for (const rKey of Object.keys(cellData)) {
        const r = Number(rKey);
        const rowData = cellData[r];
        if (!rowData) continue;
        minRow = Math.min(minRow, r);
        maxRow = Math.max(maxRow, r);
        for (const cKey of Object.keys(rowData)) {
          const c = Number(cKey);
          minCol = Math.min(minCol, c);
          maxCol = Math.max(maxCol, c);
        }
      }
      if (maxRow >= 0 && maxCol >= 0) {
        const startA1 = `${toA1Col(minCol)}${minRow + 1}`;
        const endA1 = `${toA1Col(maxCol)}${maxRow + 1}`;
        targetRange = `${startA1}:${endA1}`;
      } else {
        // Nothing to do
        return {
          success: true,
          replacements: 0,
          cellsUpdated: 0,
          range: "",
          message: "No data to scan on the active sheet",
        };
      }
    }

    // Parse A1 range
    const match = targetRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!match) throw new Error(`Invalid range: ${targetRange}`);
    const [, startColStr, startRowStr, endColStr, endRowStr] = match;
    const colToIndex = (s: string) =>
      s.split("").reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) -
      1;
    const startCol = colToIndex(startColStr);
    const endCol = colToIndex(endColStr);
    const startRow = parseInt(startRowStr, 10) - 1;
    const endRow = parseInt(endRowStr, 10) - 1;

    // Build search pattern
    let pattern: RegExp;
    if (useRegex) {
      pattern = new RegExp(findText, matchCase ? "g" : "gi");
    } else {
      const escaped = findText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const body = wholeWord ? `\\b${escaped}\\b` : escaped;
      pattern = new RegExp(body, matchCase ? "g" : "gi");
    }

    const worksheet = context.fWorksheet;
    const updatedCells: Array<{
      r: number;
      c: number;
      from: string;
      to: string;
    }> = [];
    let replacements = 0;

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = worksheet.getRange(r, c, 1, 1);
        const v = (cell as any).getValue ? (cell as any).getValue() : undefined;
        if (v == null) continue;
        if (typeof v === "string") {
          const before = v;
          const after = before.replace(pattern, (m) => {
            replacements += 1;
            return replaceText;
          });
          if (after !== before) {
            cell.setValue(after);
            updatedCells.push({ r, c, from: before, to: after });
          }
        }
      }
    }

    try {
      const formula = context.univerAPI.getFormula();
      formula.executeCalculation();
    } catch {}

    return {
      success: true,
      range: targetRange,
      replacements,
      cellsUpdated: updatedCells.length,
      message: `Replaced ${replacements} occurrence(s) across ${updatedCells.length} cell(s) in ${targetRange}.`,
      details: updatedCells.slice(0, 20),
    };
  }
);

/**
 * Export all additional tools
 */
export const ADDITIONAL_TOOLS = [
  AutoFitColumnsTool,
  FindCellTool,
  FormatAsTableTool,
  ConditionalFormattingTool,
  FindReplaceTool,
  ...AGGREGATED_CHART_TOOLS, // Add intelligent aggregated chart tools
  ...EXCEL_FUNCTION_TOOLS, // Add comprehensive Excel function support
];
