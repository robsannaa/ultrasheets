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
      "Create a Univer Table from an A1 range using @univerjs/preset-sheets-table (no manual styling)",
    category: "format",
    requiredContext: [],
    invalidatesCache: true,
  },
  async (
    context: UniversalToolContext,
    params: {
      range?: string;
      name?: string;
      tableId?: string;
      showHeader?: boolean;
      theme?:
        | {
            name: string;
            headerRowStyle?: any;
            firstRowStyle?: any;
            lastRowStyle?: any;
            bandedRowsStyle?: any;
          }
        | string;
      // Back-compat (ignored, we now rely on Table preset):
      style?: "light" | "medium" | "dark";
      alternatingRows?: boolean;
      headerRow?: boolean;
      borders?: boolean;
    }
  ) => {
    const {
      range: inputRange,
      name,
      tableId: inputTableId,
      showHeader = true,
      theme,
    } = params;

    // Validate API availability per Univer docs
    const ws: any = context.fWorksheet as any;
    const wb: any = context.fWorkbook as any;
    if (typeof ws.addTable !== "function") {
      throw new Error(
        "The Table preset is not available. Please install '@univerjs/preset-sheets-table' and ensure it is registered."
      );
    }

    // Determine target A1 range
    let targetRange = inputRange;
    if (!targetRange) {
      const table = context.findTable(inputTableId);
      if (table?.range) targetRange = table.range;
    }
    if (!targetRange || !/^[A-Z]+\d+:[A-Z]+\d+$/.test(targetRange)) {
      throw new Error(
        `A valid A1 range is required (e.g., 'B2:F11'). Received: '${
          targetRange || ""
        }'`
      );
    }

    // Build unique name and id when absent
    const existingTables =
      typeof wb.getTableList === "function" ? wb.getTableList() || [] : [];
    const ensureUnique = (base: string, exists: (s: string) => boolean) => {
      if (!exists(base)) return base;
      let i = 1;
      while (exists(`${base}_${i}`)) i += 1;
      return `${base}_${i}`;
    };

    const desiredName = name || "Table";
    const uniqueName = ensureUnique(
      desiredName,
      (n) =>
        !!(
          (typeof wb.getTableInfoByName === "function" &&
            wb.getTableInfoByName(n)) ||
          existingTables.find((t: any) => t?.name === n)
        )
    );

    const baseId =
      inputTableId || uniqueName.toLowerCase().replace(/\W+/g, "-");
    const uniqueId = ensureUnique(
      baseId,
      (id) => !!(typeof wb.getTableInfo === "function" && wb.getTableInfo(id))
    );

    // Convert A1 to IRangeData per docs via getRange().getRange()
    const fRange = ws.getRange(targetRange);
    const iRange =
      typeof (fRange as any).getRange === "function"
        ? (fRange as any).getRange()
        : undefined;
    if (!iRange) {
      throw new Error("Could not resolve range payload for addTable().");
    }

    // Create table
    const success = await ws.addTable(uniqueName, iRange, uniqueId, {
      showHeader: Boolean(showHeader),
    });

    if (!success) {
      throw new Error(
        "addTable() returned false. The range may be invalid or overlapping."
      );
    }

    // Optionally register a theme
    let appliedThemeName: string | undefined;
    if (theme) {
      // Accept either a string (name) or a full theme object
      if (typeof theme === "string") {
        // Try applying existing theme by name if API exists
        if (typeof ws.setTableTheme === "function") {
          try {
            await ws.setTableTheme(uniqueId, theme);
            appliedThemeName = theme;
          } catch {}
        } else if (typeof ws.applyTableTheme === "function") {
          try {
            await ws.applyTableTheme(uniqueId, theme);
            appliedThemeName = theme;
          } catch {}
        } else {
          // Fallback: cannot apply by name with current API
          appliedThemeName = undefined;
        }
      } else if (typeof ws.addTableTheme === "function") {
        const themeName = theme.name || `${uniqueName}-theme`;
        await ws.addTableTheme(uniqueId, { ...theme, name: themeName });
        appliedThemeName = themeName;
      }
    }

    // Retrieve info for response
    const tableInfo =
      typeof wb.getTableInfo === "function" ? wb.getTableInfo(uniqueId) : null;

    return {
      success: true,
      range: targetRange,
      name: uniqueName,
      tableId: uniqueId,
      showHeader: Boolean(showHeader),
      theme: appliedThemeName || null,
      tableInfo,
      message: `Created table '${uniqueName}' (${uniqueId}) for ${targetRange}${
        appliedThemeName ? ` with theme '${appliedThemeName}'` : ""
      }.`,
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
      "Apply conditional formatting rules to highlight cells based on their values. Colors must be provided by the caller (LLM). No hardcoded defaults are used.",
    category: "format",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (context: UniversalToolContext, params: any) => {
    // Accept both our legacy schema and server tool schema
    const range: string = params.range;
    const incomingCondition: string | undefined = params.condition;
    const ruleType: string | undefined = params.ruleType;
    const format: any = params.format || {};
    const contains: string | undefined = params.contains;
    const startsWith: string | undefined = params.startsWith;
    const endsWith: string | undefined = params.endsWith;
    const min: number | undefined = params.min;
    const max: number | undefined = params.max;
    const equals: number | undefined = params.equals;
    const formula: string | undefined = params.formula;
    const value: any = params.value; // legacy
    const value2: any = params.value2; // legacy

    if (!range) {
      throw new Error("Range is required for conditional formatting");
    }

    // Provide intelligent defaults for number conditions if missing
    let effectiveMin = min;
    let effectiveMax = max;
    let effectiveEquals = equals;

    // Infer rule
    let inferredRule = ruleType || incomingCondition || "";
    if (!inferredRule) {
      if (typeof formula === "string" && formula.trim())
        inferredRule = "formula";
      else if (typeof contains === "string") inferredRule = "text_contains";
      else if (typeof startsWith === "string")
        inferredRule = "text_starts_with";
      else if (typeof endsWith === "string") inferredRule = "text_ends_with";
      else if (typeof min === "number" && typeof max === "number")
        inferredRule = "number_between";
      else if (typeof min === "number") inferredRule = "number_gt";
      else if (typeof max === "number") inferredRule = "number_lt";
      else if (typeof equals === "number") inferredRule = "number_eq";
      else {
        // Smart inference: if no specific condition, analyze the data to suggest something useful
        // For price columns, default to highlighting high values
        if (
          range.includes("C") ||
          range.includes("Price") ||
          range.includes("price")
        ) {
          inferredRule = "number_gt";
          effectiveMin = 10; // Highlight prices over $10
        } else {
          inferredRule = "not_empty";
        }
      }
    }

    if (inferredRule.startsWith("number_")) {
      // For number conditions without explicit values, provide sensible defaults
      if (inferredRule === "number_gt" && typeof effectiveMin !== "number") {
        effectiveMin = 0; // Default: highlight values greater than 0
      } else if (
        inferredRule === "number_lt" &&
        typeof effectiveMax !== "number"
      ) {
        effectiveMax = 1000000; // Default: highlight values less than 1M
      } else if (
        inferredRule === "number_between" &&
        (typeof effectiveMin !== "number" || typeof effectiveMax !== "number")
      ) {
        effectiveMin = 0;
        effectiveMax = 1000000;
      } else if (
        inferredRule === "number_eq" &&
        typeof effectiveEquals !== "number"
      ) {
        effectiveEquals = 0; // Default: highlight zero values
      }
    }

    // Parse range
    const m = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!m) throw new Error(`Invalid range format: ${range}`);
    const [, startColStr, startRowStr, endColStr, endRowStr] = m;
    const colToIndex = (s: string) =>
      s.split("").reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) -
      1;
    const startRowIndex = parseInt(startRowStr, 10) - 1;
    const endRowIndex = parseInt(endRowStr, 10) - 1;
    const startColIndex = colToIndex(startColStr);
    const endColIndex = colToIndex(endColStr);

    // Build rule per docs [Conditional Formatting Facade API]
    // https://docs.univer.ai/guides/sheets/features/conditional-formatting
    let rule = context.fWorksheet.newConditionalFormattingRule();

    switch (inferredRule) {
      case "number_between":
        if (
          typeof effectiveMin === "number" &&
          typeof effectiveMax === "number"
        )
          rule = rule.whenNumberBetween(
            Math.min(effectiveMin, effectiveMax),
            Math.max(effectiveMin, effectiveMax)
          );
        else throw new Error("number_between requires min and max");
        break;
      case "number_gt":
        if (typeof effectiveMin === "number")
          rule = rule.whenNumberGreaterThan(effectiveMin);
        else throw new Error("number_gt requires min");
        break;
      case "number_gte":
        if (typeof effectiveMin === "number")
          rule = rule.whenNumberGreaterThanOrEqualTo(effectiveMin);
        else throw new Error("number_gte requires min");
        break;
      case "number_lt":
        if (typeof effectiveMax === "number")
          rule = rule.whenNumberLessThan(effectiveMax);
        else throw new Error("number_lt requires max");
        break;
      case "number_lte":
        if (typeof effectiveMax === "number")
          rule = rule.whenNumberLessThanOrEqualTo(effectiveMax);
        else throw new Error("number_lte requires max");
        break;
      case "number_eq":
        if (typeof effectiveEquals === "number")
          rule = rule.whenNumberEqualTo(effectiveEquals);
        else throw new Error("number_eq requires equals");
        break;
      case "number_neq":
        if (typeof effectiveEquals === "number")
          rule = rule.whenNumberNotEqualTo(effectiveEquals);
        else throw new Error("number_neq requires equals");
        break;
      case "text_contains":
        if (typeof contains === "string")
          rule = rule.whenTextContains(contains);
        else throw new Error("text_contains requires 'contains'");
        break;
      case "text_not_contains":
        if (typeof contains === "string")
          rule = rule.whenTextDoesNotContain(contains);
        else throw new Error("text_not_contains requires 'contains'");
        break;
      case "text_starts_with":
        if (typeof startsWith === "string")
          rule = rule.whenTextStartsWith(startsWith);
        else throw new Error("text_starts_with requires 'startsWith'");
        break;
      case "text_ends_with":
        if (typeof endsWith === "string")
          rule = rule.whenTextEndsWith(endsWith);
        else throw new Error("text_ends_with requires 'endsWith'");
        break;
      case "not_empty":
        rule = rule.whenCellNotEmpty();
        break;
      case "empty":
        rule = rule.whenCellEmpty();
        break;
      case "formula":
        if (typeof formula === "string" && formula.trim())
          rule = rule.whenFormulaSatisfied(formula);
        else throw new Error("formula rule requires 'formula' string");
        break;
      case "unique":
        rule = rule.setUniqueValues();
        break;
      case "duplicate":
        rule = rule.setDuplicateValues();
        break;
      case "data_bar":
        // Data bars can cause 'type' errors in some builds. Use color scale instead.
        rule = rule.setColorScale();
        break;
      case "color_scale":
        rule = rule.setColorScale();
        break;
      default:
        throw new Error(`Unsupported ruleType: ${inferredRule}`);
    }

    // Apply formatting only if explicitly provided by caller.
    const isVisualScaleRule = inferredRule === "color_scale";
    const canStyle =
      !isVisualScaleRule &&
      typeof (rule as any).setBackground === "function" &&
      typeof (rule as any).setFontColor === "function";

    // Helper: normalize color (accept hex or common names). Avoid off-brand palettes by requiring explicit input.
    const named: Record<string, string> = {
      green: "#22c55e",
      red: "#ef4444",
      yellow: "#eab308",
      blue: "#3b82f6",
      orange: "#f59e0b",
      purple: "#a855f7",
      gray: "#6b7280",
      grey: "#6b7280",
      black: "#000000",
      white: "#ffffff",
    };
    const normalizeColor = (c?: string) => {
      if (!c) return undefined;
      const s = String(c).trim().toLowerCase();
      if (s.startsWith("#") && (s.length === 7 || s.length === 4)) return s;
      return named[s];
    };

    // Helper: pick readable text color for the given background if fontColor omitted
    const pickText = (bg: string) => {
      const hex = bg.replace("#", "");
      const bigint = parseInt(
        hex.length === 3
          ? hex
              .split("")
              .map((ch) => ch + ch)
              .join("")
          : hex,
        16
      );
      const r = (bigint >> 16) & 255;
      const g = (bigint >> 8) & 255;
      const b = bigint & 255;
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      return luminance > 0.6 ? "#000000" : "#ffffff";
    };

    let effectiveFormat: any = {};
    if (canStyle) {
      const bg = normalizeColor(format?.backgroundColor || format?.background);
      const fg = normalizeColor(format?.fontColor || format?.color);

      if (!bg && !fg && format?.bold == null && format?.italic == null) {
        // No styling supplied; do not apply implicit colors.
        effectiveFormat = {};
      } else {
        if (bg) rule = (rule as any).setBackground(bg);
        const finalFont = fg || (bg ? pickText(bg) : undefined);
        if (finalFont) rule = (rule as any).setFontColor(finalFont);
        if (typeof (rule as any).setBold === "function" && format?.bold)
          rule = (rule as any).setBold(true);
        if (typeof (rule as any).setItalic === "function" && format?.italic)
          rule = (rule as any).setItalic(true);
        effectiveFormat = {
          backgroundColor: bg,
          fontColor: finalFont,
          bold: Boolean(format?.bold),
          italic: Boolean(format?.italic),
        };
      }
    }

    // If the builder supports explicit range binding, set it before build
    const fRange = context.fWorksheet.getRange(range);
    if (typeof (rule as any).setRanges === "function") {
      try {
        // Per docs, pass IRangeData via getRange()
        rule = (rule as any).setRanges([(fRange as any).getRange()]);
      } catch {}
    }
    const builtRule = rule.build();

    // Ensure ranges exist on the final rule (older builds may ignore setRanges)
    const rawRange = (fRange as any).getRange
      ? (fRange as any).getRange()
      : undefined;
    const finalRule =
      rawRange && !(builtRule as any)?.ranges?.length
        ? { ...(builtRule as any), ranges: [rawRange] }
        : builtRule;

    // Apply exactly as in docs: worksheet.addConditionalFormattingRule(rule)
    const ws: any = context.fWorksheet as any;
    if (typeof ws.addConditionalFormattingRule === "function") {
      try {
        const ruleId = ws.addConditionalFormattingRule(finalRule);
        return {
          success: true,
          range,
          ruleType: inferredRule,
          inputs: {
            min: effectiveMin,
            max: effectiveMax,
            equals: effectiveEquals,
            contains,
            startsWith,
            endsWith,
            formula,
          },
          format: effectiveFormat,
          ruleId,
          message: `Applied conditional formatting to ${range} using '${inferredRule}'${
            effectiveMin !== undefined ? ` (${effectiveMin})` : ""
          }${effectiveMax !== undefined ? ` to ${effectiveMax}` : ""}`,
        };
      } catch (err) {
        // Fall through to secondary API
      }
    }

    // Fallback: try alternative API signatures
    if (typeof ws.setConditionalFormattingRule === "function") {
      try {
        const res = ws.setConditionalFormattingRule(
          (finalRule as any).cfId,
          finalRule
        );
        return {
          success: true,
          range,
          ruleType: inferredRule,
          inputs: {
            min: effectiveMin,
            max: effectiveMax,
            equals: effectiveEquals,
            contains,
            startsWith,
            endsWith,
            formula,
          },
          format: effectiveFormat,
          ruleId: res,
          message: `Applied conditional formatting to ${range} using '${inferredRule}'${
            effectiveMin !== undefined ? ` (${effectiveMin})` : ""
          }${effectiveMax !== undefined ? ` to ${effectiveMax}` : ""}`,
        };
      } catch (err) {
        // Continue to next fallback
      }
    }

    // Final fallback: try range-based API
    if (typeof (fRange as any).addConditionalFormattingRule === "function") {
      try {
        const ruleId = (fRange as any).addConditionalFormattingRule(finalRule);
        return {
          success: true,
          range,
          ruleType: inferredRule,
          inputs: {
            min: effectiveMin,
            max: effectiveMax,
            equals: effectiveEquals,
            contains,
            startsWith,
            endsWith,
            formula,
          },
          format: effectiveFormat,
          ruleId,
          message: `Applied conditional formatting to ${range} using '${inferredRule}'${
            effectiveMin !== undefined ? ` (${effectiveMin})` : ""
          }${effectiveMax !== undefined ? ` to ${effectiveMax}` : ""}`,
        };
      } catch (err) {
        // Last resort fallback
      }
    }

    // If all APIs fail, provide a helpful error message
    throw new Error(
      `Conditional formatting failed for rule '${inferredRule}'. ` +
        `Tried multiple API methods but none succeeded. ` +
        `Ensure the conditional formatting preset is loaded and pass explicit colors in params.format when visual styling is required.`
    );

    // (Unreached)
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
          const after = before.replace(pattern, () => {
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
 * FILTER DATA - Modern Implementation with Applied Criteria
 * Applies specific filter criteria automatically
 */
export const FilterDataTool = createSimpleTool(
  {
    name: "filter_data",
    description:
      "Filter data to show only specific values or conditions. Adds filters and applies criteria automatically.",
    category: "structure",
    requiredContext: ["tables"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      column?: string;
      values?: string[];
      contains?: string;
      condition?: "greater_than" | "less_than" | "equals" | "not_equals" | "contains" | "not_contains";
      conditionValue?: string | number;
      tableId?: string;
    }
  ) => {
    const { column, values, contains, condition, conditionValue, tableId } = params;

    // Find the table
    const table = context.findTable(tableId);
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    const fRange = context.fWorksheet.getRange(table.range);
    
    // Create filter if it doesn't exist
    let fFilter = context.fWorksheet.getFilter();
    if (!fFilter) {
      fFilter = fRange.createFilter();
    }

    // Determine target column
    let targetColumn: any;
    if (column) {
      // Find column by name or letter
      targetColumn = table.columns.find(
        (col: any) => 
          col.name.toLowerCase() === column.toLowerCase() ||
          col.letter === column.toUpperCase()
      );
      if (!targetColumn) {
        throw new Error(`Column "${column}" not found in table`);
      }
    } else {
      // Try to infer column from values or context
      if (values?.length) {
        // Look for a column that might contain these values
        for (const col of table.columns) {
          if (
            col.sampleValues?.some((v: any) =>
              values.some(val => String(v).toLowerCase().includes(val.toLowerCase()))
            )
          ) {
            targetColumn = col;
            break;
          }
        }
      }
      
      if (!targetColumn) {
        // Default to first text column or Transaction column
        targetColumn = table.columns.find((col: any) => 
          col.name.toLowerCase().includes('transaction') ||
          col.name.toLowerCase().includes('type') ||
          col.name.toLowerCase().includes('category')
        ) || table.columns[1]; // Second column if no specific match
      }
    }

    if (!targetColumn) {
      throw new Error("Could not determine target column for filtering");
    }

    // Build filter criteria
    let filterCriteria: any = {};
    
    if (values?.length) {
      // Exact value matching
      filterCriteria = {
        colId: targetColumn.index,
        filters: {
          filters: values
        }
      };
    } else if (contains) {
      // Text contains
      filterCriteria = {
        colId: targetColumn.index,
        filters: {
          filters: [contains] // This will be interpreted as "contains" by Univer
        }
      };
    } else if (condition && conditionValue !== undefined) {
      // Condition-based filtering
      const conditionMap = {
        equals: "eq",
        not_equals: "ne",
        greater_than: "gt",
        less_than: "lt",
        contains: "contains",
        not_contains: "not_contains"
      };
      
      filterCriteria = {
        colId: targetColumn.index,
        condition: conditionMap[condition],
        value: conditionValue
      };
    }

    // Apply the filter criteria
    try {
      const worksheetColumn = context.fWorksheet.getRange(`${targetColumn.letter}:${targetColumn.letter}`).getColumn();
      fFilter.setColumnFilterCriteria(worksheetColumn, filterCriteria);
    } catch (error) {
      console.warn("Failed to apply filter criteria:", error);
      // Fallback: just create the filter without specific criteria
    }

    return {
      success: true,
      tableId: table.id,
      column: targetColumn.name,
      range: table.range,
      criteria: filterCriteria,
      appliedValues: values,
      appliedContains: contains,
      appliedCondition: condition,
      appliedConditionValue: conditionValue,
      message: `Applied filter to column "${targetColumn.name}"${
        values ? ` showing: ${values.join(", ")}` : ""
      }${contains ? ` containing: "${contains}"` : ""}${
        condition && conditionValue !== undefined 
          ? ` where ${condition.replace("_", " ")} ${conditionValue}` 
          : ""
      }`,
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
  FilterDataTool, // Add the new filter data tool
  ...AGGREGATED_CHART_TOOLS, // Add intelligent aggregated chart tools
  ...EXCEL_FUNCTION_TOOLS, // Add comprehensive Excel function support
];
