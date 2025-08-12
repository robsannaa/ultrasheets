/**
 * Migrated Legacy Tools - Modern Implementations
 *
 * This file contains all the legacy tools migrated to the new Universal Context framework.
 * Each tool is implemented using the modern patterns with intelligent context management.
 */

import { createSimpleTool } from "../tool-executor";
import type { UniversalToolContext } from "../universal-context";

/**
 * CALCULATE TOTAL - Modern Implementation
 * Migrated from: executeCalculateTotal
 */
export const CalculateTotalTool = createSimpleTool(
  {
    name: "calculate_total",
    description: "Calculate total for a specific column in a table",
    category: "data",
    requiredContext: ["tables", "columns"],
    invalidatesCache: true,
  },
  async (
    context: UniversalToolContext,
    params: {
      column: string;
      tableId?: string;
      aggregation?: "sum" | "average" | "count" | "max" | "min";
    }
  ) => {
    const { column, tableId, aggregation = "sum" } = params;

    const table = context.findTable(tableId);
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    const targetColumn = context.findColumn(column, tableId);
    if (!targetColumn) {
      throw new Error(
        `Column '${column}' not found. Available columns: ${table.columns
          .map((c: any) => c.name)
          .join(", ")}`
      );
    }

    // Determine the formula based on aggregation type
    const formulaMap = {
      sum: "SUM",
      average: "AVERAGE",
      count: "COUNT",
      max: "MAX",
      min: "MIN",
    };

    const formulaFunction = formulaMap[aggregation];
    const dataRange = context.getColumnRange(targetColumn.name, false, tableId);
    const formula = `=${formulaFunction}(${dataRange})`;

    // Place the result below the table
    const sumRow = table.position.endRow + 1;
    const sumCell = `${targetColumn.letter}${sumRow + 1}`;

    // Set the formula
    const target = context.fWorksheet.getRange(
      sumRow,
      targetColumn.index,
      1,
      1
    );
    if (typeof (target as any).setFormula === "function") {
      (target as any).setFormula(formula);
    } else {
      target.setValue(formula);
    }

    // Execute calculation
    try {
      const formulaService = context.univerAPI.getFormula();
      formulaService.executeCalculation();
    } catch (calcError) {
      console.warn("⚠️ Formula calculation failed:", calcError);
    }

    // Track for later formatting
    if (!(window as any).recentTotals) (window as any).recentTotals = [];
    (window as any).recentTotals.push({
      cell: sumCell,
      column: targetColumn.name,
      columnType: targetColumn.dataType,
      isCurrency: targetColumn.isCurrency,
      tableId: table.id,
      aggregation,
      timestamp: Date.now(),
    });

    return {
      column: targetColumn.name,
      cell: sumCell,
      formula,
      aggregation,
      dataRange,
      tableId: table.id,
      message: `Added ${aggregation} for column '${targetColumn.name}' at ${sumCell}`,
    };
  }
);

/**
 * FORMAT RECENT TOTALS - Modern Implementation
 * Migrated from: executeFormatRecentTotals
 */
export const FormatRecentTotalsTool = createSimpleTool(
  {
    name: "format_recent_totals",
    description: "Format recently added totals with currency formatting",
    category: "format",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      currency?: string;
      decimals?: number;
      columnPattern?: string;
      maxAge?: number;
    }
  ) => {
    const {
      currency = "USD",
      decimals = 2,
      columnPattern,
      maxAge = 300000, // 5 minutes
    } = params;

    const recentTotals = (window as any).recentTotals || [];
    if (recentTotals.length === 0) {
      throw new Error("No recent totals found to format");
    }

    // Filter recent totals based on criteria
    const cutoffTime = Date.now() - maxAge;
    let totalsToFormat = recentTotals.filter(
      (total: any) => total.timestamp > cutoffTime
    );

    // Apply column pattern filter if provided
    if (columnPattern) {
      totalsToFormat = totalsToFormat.filter((total: any) =>
        total.column.toLowerCase().includes(columnPattern.toLowerCase())
      );
    } else {
      // Auto-detect currency totals
      totalsToFormat = totalsToFormat.filter((total: any) => {
        const columnName = total.column.toLowerCase();
        return (
          total.isCurrency ||
          columnName.includes("price") ||
          columnName.includes("cost") ||
          columnName.includes("amount") ||
          columnName.includes("revenue") ||
          columnName.includes("value") ||
          columnName.includes("sales")
        );
      });
    }

    if (totalsToFormat.length === 0) {
      throw new Error("No matching currency totals found to format");
    }

    // Currency format patterns
    const currencyFormats: { [key: string]: string } = {
      USD: "$#,##0.00",
      EUR: "€#,##0.00",
      GBP: "£#,##0.00",
      JPY: "¥#,##0",
      PLN: "#,##0.00zł",
      CAD: "C$#,##0.00",
      AUD: "A$#,##0.00",
      CHF: "CHF#,##0.00",
      CNY: "¥#,##0.00",
      INR: "₹#,##0.00",
    };

    const formatPattern =
      currencyFormats[currency.toUpperCase()] ||
      `${currency}#,##0.${"0".repeat(decimals)}`;

    const formattedCells = [];

    // Format each total
    for (const total of totalsToFormat) {
      try {
        const fRange = context.fWorksheet.getRange(total.cell);
        fRange.setNumberFormat(formatPattern);
        formattedCells.push(total.cell);
      } catch (formatError) {
        console.warn(`⚠️ Failed to format cell ${total.cell}:`, formatError);
      }
    }

    return {
      currency,
      decimals,
      formatPattern,
      formattedCells,
      totalCount: formattedCells.length,
      message: `Formatted ${
        formattedCells.length
      } recent totals as ${currency}: ${formattedCells.join(", ")}`,
    };
  }
);

/**
 * CREATE PIVOT TABLE - Modern Implementation
 * Migrated from: executeCreatePivotTable
 */
export const CreatePivotTableTool = createSimpleTool(
  {
    name: "create_pivot_table",
    description:
      "Create a pivot table from table data with intelligent analysis",
    category: "analysis",
    requiredContext: ["tables", "spatial"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      groupBy: string;
      valueColumn: string;
      aggFunc?: "sum" | "count" | "average" | "max" | "min";
      tableId?: string;
      destination?: string;
    }
  ) => {
    const {
      groupBy,
      valueColumn,
      aggFunc = "sum",
      tableId,
      destination,
    } = params;

    const table = context.findTable(tableId);
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    const groupColumn = context.findColumn(groupBy, tableId);
    const dataColumn = context.findColumn(valueColumn, tableId);

    if (!groupColumn) {
      throw new Error(`Group column '${groupBy}' not found`);
    }
    if (!dataColumn) {
      throw new Error(`Value column '${valueColumn}' not found`);
    }

    // Get data from the source table
    const dataRange = table.range;
    const snapshot = context.activeSheetSnapshot;
    const cellData = snapshot.cellData || {};

    // Extract and process data
    const pivotData = new Map<string, number[]>();

    for (
      let row = table.position.startRow + 1;
      row <= table.position.endRow;
      row++
    ) {
      const rowData = cellData[row] || {};
      const groupValue = rowData[groupColumn.index]?.v?.toString() || "";
      const dataValue = parseFloat(rowData[dataColumn.index]?.v) || 0;

      if (groupValue && !isNaN(dataValue)) {
        if (!pivotData.has(groupValue)) {
          pivotData.set(groupValue, []);
        }
        pivotData.get(groupValue)!.push(dataValue);
      }
    }

    // Calculate aggregated values
    const aggregatedData = new Map<string, number>();

    for (const [group, values] of pivotData) {
      let result = 0;
      switch (aggFunc) {
        case "sum":
          result = values.reduce((sum, val) => sum + val, 0);
          break;
        case "count":
          result = values.length;
          break;
        case "average":
          result = values.reduce((sum, val) => sum + val, 0) / values.length;
          break;
        case "max":
          result = Math.max(...values);
          break;
        case "min":
          result = Math.min(...values);
          break;
      }
      aggregatedData.set(group, result);
    }

    // Determine destination for pivot table
    let destPosition;
    if (destination) {
      destPosition = destination;
    } else {
      // Use spatial analysis to find optimal placement
      const placement = context.findOptimalPlacement(
        3,
        aggregatedData.size + 2
      );
      destPosition = placement.range;
    }

    // Parse destination position
    const match = destPosition.match(/([A-Z]+)(\d+)/);
    if (!match) {
      throw new Error(`Invalid destination position: ${destPosition}`);
    }

    const [, startCol, startRow] = match;
    const startColIndex =
      startCol
        .split("")
        .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
    const startRowIndex = parseInt(startRow) - 1;

    // Write pivot table headers
    context.fWorksheet
      .getRange(startRowIndex, startColIndex, 1, 1)
      .setValue(groupBy);
    context.fWorksheet
      .getRange(startRowIndex, startColIndex + 1, 1, 1)
      .setValue(`${aggFunc.toUpperCase()}(${valueColumn})`);

    // Write pivot table data
    let currentRow = startRowIndex + 1;
    for (const [group, value] of aggregatedData) {
      context.fWorksheet
        .getRange(currentRow, startColIndex, 1, 1)
        .setValue(group);
      context.fWorksheet
        .getRange(currentRow, startColIndex + 1, 1, 1)
        .setValue(value);
      currentRow++;
    }

    const pivotRange = `${startCol}${startRow}:${String.fromCharCode(
      65 + startColIndex + 1
    )}${startRow + aggregatedData.size}`;

    return {
      groupBy,
      valueColumn,
      aggFunc,
      sourceTable: table.id,
      sourceRange: dataRange,
      pivotRange,
      destination: destPosition,
      groupCount: aggregatedData.size,
      message: `Created pivot table at ${pivotRange} grouping by '${groupBy}' with ${aggFunc} of '${valueColumn}'`,
    };
  }
);

/**
 * SWITCH SHEET - Modern Implementation with Intelligent Sheet Management
 * Automatically detects missing sheets and offers to create them
 */
export const SwitchSheetTool = createSimpleTool(
  {
    name: "switch_sheet",
    description: "Switch to a different sheet, analyze sheet structure, or intelligently handle missing sheets",
    category: "navigation",
    requiredContext: [],
    invalidatesCache: true, // Switching sheets invalidates current context
  },
  async (
    context: UniversalToolContext,
    params: {
      sheetName: string;
      action?: "switch" | "analyze" | "create_if_missing";
    }
  ) => {
    const { sheetName, action = "switch" } = params;
    const workbook = context.fWorkbook;

    // First, check if the sheet exists
    const sheets = workbook.getSheets?.() || [];
    const availableSheetNames = sheets.map(
      (s: any) =>
        s.getName?.() || s.getSheetName?.() || s.name || "Unnamed"
    );
    
    const targetSheet = sheets.find(
      (sheet: any) =>
        sheet.getName?.() === sheetName ||
        sheet.getSheetName?.() === sheetName ||
        sheet.name === sheetName
    );

    // Handle missing sheets intelligently
    if (!targetSheet) {
      return {
        action: "sheet_missing",
        sheetName,
        availableSheets: availableSheetNames,
        suggestion: "create_sheet",
        message: `The sheet '${sheetName}' doesn't exist. Available sheets: ${availableSheetNames.join(", ")}. Would you like me to create the '${sheetName}' sheet for you?`,
        userFriendlyMessage: `I couldn't find the '${sheetName}' sheet. The available sheets are: ${availableSheetNames.join(", ")}. Should I create the '${sheetName}' sheet for you?`,
      };
    }

    if (action === "analyze") {
      try {
        // Get basic info about the sheet
        const sheetSnapshot = targetSheet.getSnapshot?.() || {};

        return {
          action: "analyze",
          sheetName,
          exists: true,
          info: {
            name: sheetName,
            hasData: !!(
              sheetSnapshot.cellData &&
              Object.keys(sheetSnapshot.cellData).length > 0
            ),
            cellCount: sheetSnapshot.cellData
              ? Object.keys(sheetSnapshot.cellData).length
              : 0,
          },
          message: `Analyzed sheet '${sheetName}': ${sheetSnapshot.cellData ? "contains data" : "empty"}`,
        };
      } catch (error) {
        throw new Error(
          `Failed to analyze sheet '${sheetName}': ${
            error instanceof Error ? error.message : "Unknown error"
          }`
        );
      }
    }

    // Switch to the target sheet
    try {
      let switchSuccess = false;

      // Try different methods to switch sheets
      if (typeof workbook.setActiveSheet === "function") {
        switchSuccess = workbook.setActiveSheet(sheetName);
        
        // Also try with sheet ID if name fails
        if (!switchSuccess && typeof targetSheet.getSheetId === "function") {
          switchSuccess = workbook.setActiveSheet(targetSheet.getSheetId());
        }
      } else if (typeof workbook.switchSheet === "function") {
        await workbook.switchSheet(sheetName);
        switchSuccess = true;
      } else if (typeof targetSheet.activate === "function") {
        targetSheet.activate();
        switchSuccess = true;
      }

      if (!switchSuccess) {
        return {
          action: "switch_failed",
          sheetName,
          availableSheets: availableSheetNames,
          exists: true,
          message: `The sheet '${sheetName}' exists but I couldn't switch to it. This might be a technical issue with the sheet switching functionality.`,
          suggestion: "Try refreshing the page or creating a new sheet instead.",
        };
      }

      // Invalidate context since we switched sheets
      context.invalidateCache();

      return {
        action: "switch",
        sheetName,
        exists: true,
        previousSheet: context.activeSheetName,
        success: true,
        message: `Successfully switched to '${sheetName}'`,
      };
    } catch (error) {
      return {
        action: "switch_error",
        sheetName,
        exists: true,
        error: error instanceof Error ? error.message : "Unknown error",
        message: `Failed to switch to sheet '${sheetName}': ${
          error instanceof Error ? error.message : "Unknown error"
        }. The sheet exists but couldn't be accessed.`,
        suggestion: "Try creating a new sheet or refreshing the page.",
      };
    }
  }
);

/**
 * CREATE SHEET - Modern Implementation
 * Creates a new worksheet in the workbook
 */
export const CreateSheetTool = createSimpleTool(
  {
    name: "create_sheet",
    description: "Create a new worksheet in the workbook",
    category: "navigation",
    requiredContext: [],
    invalidatesCache: true, // Creating sheets invalidates current context
  },
  async (
    context: UniversalToolContext,
    params: {
      sheetName: string;
      switchToSheet?: boolean;
    }
  ) => {
    const { sheetName, switchToSheet = true } = params;

    try {
      const workbook = context.fWorkbook;

      // Check if sheet already exists
      const sheets = workbook.getSheets?.() || [];
      const existingSheet = sheets.find(
        (sheet: any) =>
          sheet.getName?.() === sheetName ||
          sheet.getSheetName?.() === sheetName ||
          sheet.name === sheetName
      );

      if (existingSheet) {
        // If sheet exists and we want to switch to it, do that
        if (switchToSheet) {
          if (typeof workbook.setActiveSheet === "function") {
            const success = workbook.setActiveSheet(sheetName);
            if (!success) {
              throw new Error(`Failed to switch to existing sheet '${sheetName}'`);
            }
            context.invalidateCache();
            return {
              created: false,
              switched: true,
              sheetName,
              message: `Sheet '${sheetName}' already exists. Switched to it.`,
            };
          }
        }
        return {
          created: false,
          switched: false,
          sheetName,
          message: `Sheet '${sheetName}' already exists.`,
        };
      }

      // Create the new sheet using the FWorkbook API
      if (typeof workbook.create === "function") {
        try {
          // FWorkbook.create(name, rows?, columns?) method
          const newSheet = workbook.create(sheetName, 1000, 26); // Default size
          
          if (!newSheet) {
            throw new Error(`Failed to create sheet '${sheetName}' - create() returned null`);
          }

          // Switch to the new sheet if requested
          if (switchToSheet) {
            if (typeof workbook.setActiveSheet === "function") {
              // Try with sheet name first
              let success = workbook.setActiveSheet(sheetName);
              
              // If that fails, try with the sheet object
              if (!success && typeof newSheet.getSheetId === "function") {
                success = workbook.setActiveSheet(newSheet.getSheetId());
              }
              
              // If still fails, try activating the sheet directly
              if (!success && typeof newSheet.activate === "function") {
                newSheet.activate();
                success = true;
              }
              
              if (!success) {
                console.warn(`Sheet '${sheetName}' created but could not switch to it`);
              }
            }
          }

          // Invalidate context since we created/switched sheets
          context.invalidateCache();

          return {
            created: true,
            switched: switchToSheet,
            sheetName,
            message: switchToSheet 
              ? `Created sheet '${sheetName}' and switched to it`
              : `Created sheet '${sheetName}'`,
          };
        } catch (createError) {
          console.error("Error creating sheet with workbook.create:", createError);
          // Fall through to other methods
        }
      }
      
      // Try using the lower-level Univer API
      if (context.univerAPI) {
        try {
          // Get the Univer instance and create a new worksheet
          const univerInstance = context.univerAPI._univer || context.univerAPI.getUniver?.();
          if (univerInstance) {
            // Create worksheet data
            const worksheetData = {
              id: `sheet_${Date.now()}`, // Generate unique ID
              name: sheetName,
              rowCount: 1000,
              columnCount: 26,
              cellData: {},
            };
            
            // Add the worksheet to the workbook
            const workbookData = univerInstance.getUniverSheet()?.getActiveWorkbook()?.getSnapshot();
            if (workbookData && workbookData.sheets) {
              workbookData.sheets[worksheetData.id] = worksheetData;
              
              // Trigger workbook update
              await univerInstance.getUniverSheet()?.getActiveWorkbook()?.getCommandService()?.syncExecuteCommand({
                id: 'sheet.command.insert-sheet',
                params: { worksheet: worksheetData }
              });
            }
          }
          
          // Switch to the new sheet if requested
          if (switchToSheet && typeof workbook.setActiveSheet === "function") {
            workbook.setActiveSheet(sheetName);
          }

          context.invalidateCache();
          
          return {
            created: true,
            switched: switchToSheet,
            sheetName,
            message: switchToSheet 
              ? `Created sheet '${sheetName}' using Univer core API and switched to it`
              : `Created sheet '${sheetName}' using Univer core API`,
          };
        } catch (coreError) {
          console.error("Error creating sheet with Univer core API:", coreError);
        }
      }
      
      throw new Error(`Sheet creation failed - tried multiple API methods. Available methods: ${Object.keys(workbook).filter(k => typeof workbook[k] === 'function').join(', ')}`);
    } catch (error) {
      throw new Error(
        `Failed to create sheet '${sheetName}': ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }
);

/**
 * FORMAT CELLS - Modern Implementation
 * Migrated from: executeFormatCells (part of legacy tools)
 */
export const FormatCellsTool = createSimpleTool(
  {
    name: "format_cells",
    description:
      "Apply comprehensive formatting to cells (bold, colors, alignment, etc.)",
    category: "format",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      range: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontSize?: number;
      fontColor?: string;
      backgroundColor?: string;
      textAlign?: "left" | "center" | "right";
      verticalAlign?: "top" | "middle" | "bottom";
      textRotation?: number;
      textWrap?: "overflow" | "truncate" | "wrap";
      numberFormat?: string;
    }
  ) => {
    const {
      range,
      bold,
      italic,
      underline,
      fontSize,
      fontColor,
      backgroundColor,
      textAlign,
      verticalAlign,
      textRotation,
      textWrap,
      numberFormat,
    } = params;

    if (!range) {
      throw new Error("Range is required for cell formatting");
    }

    const fRange = context.fWorksheet.getRange(range);
    const appliedFormats = [];

    // Apply text formatting
    if (bold !== undefined) {
      fRange.setBold(bold);
      appliedFormats.push(`bold: ${bold}`);
    }

    if (italic !== undefined) {
      fRange.setItalic(italic);
      appliedFormats.push(`italic: ${italic}`);
    }

    if (underline !== undefined) {
      fRange.setUnderline(underline);
      appliedFormats.push(`underline: ${underline}`);
    }

    if (fontSize !== undefined) {
      fRange.setFontSize(fontSize);
      appliedFormats.push(`fontSize: ${fontSize}`);
    }

    if (fontColor) {
      fRange.setFontColor(fontColor);
      appliedFormats.push(`fontColor: ${fontColor}`);
    }

    if (backgroundColor) {
      fRange.setBackgroundColor(backgroundColor);
      appliedFormats.push(`backgroundColor: ${backgroundColor}`);
    }

    // Apply alignment
    if (textAlign) {
      fRange.setHorizontalAlignment(textAlign);
      appliedFormats.push(`textAlign: ${textAlign}`);
    }

    if (verticalAlign) {
      fRange.setVerticalAlignment(verticalAlign);
      appliedFormats.push(`verticalAlign: ${verticalAlign}`);
    }

    if (textRotation !== undefined) {
      fRange.setTextRotation(textRotation);
      appliedFormats.push(`textRotation: ${textRotation}`);
    }

    if (textWrap) {
      const wrapMode =
        textWrap === "wrap"
          ? "WRAP"
          : textWrap === "truncate"
          ? "CLIP"
          : "OVERFLOW";
      fRange.setWrapStrategy(wrapMode);
      appliedFormats.push(`textWrap: ${textWrap}`);
    }

    // Apply number formatting
    if (numberFormat) {
      fRange.setNumberFormat(numberFormat);
      appliedFormats.push(`numberFormat: ${numberFormat}`);
    }

    return {
      range,
      appliedFormats,
      message: `Applied formatting to ${range}: ${appliedFormats.join(", ")}`,
    };
  }
);

/**
 * LIST COLUMNS - Modern Implementation
 * Migrated from: executeListColumns
 */
export const ListColumnsTool = createSimpleTool(
  {
    name: "list_columns",
    description: "List all columns in a table with their properties",
    category: "navigation",
    requiredContext: ["tables", "columns"],
    invalidatesCache: false,
  },
  async (context: UniversalToolContext, params: { tableId?: string }) => {
    const table = context.findTable(params.tableId);
    if (!table) {
      throw new Error(`Table ${params.tableId || "primary"} not found`);
    }

    const columnInfo = table.columns.map((column: any, index: number) => ({
      index,
      name: column.name,
      letter: column.letter,
      dataType: column.dataType,
      isNumeric: column.isNumeric,
      isCalculable: column.isCalculable,
      isCurrency: column.isCurrency,
      sampleValues: column.sampleValues || [],
    }));

    return {
      tableId: table.id,
      tableRange: table.range,
      columnCount: columnInfo.length,
      columns: columnInfo,
      calculableColumns: columnInfo
        .filter((col: any) => col.isCalculable)
        .map((col: any) => col.name),
      numericColumns: columnInfo
        .filter((col: any) => col.isNumeric)
        .map((col: any) => col.name),
      currencyColumns: columnInfo
        .filter((col: any) => col.isCurrency)
        .map((col: any) => col.name),
      message: `Found ${columnInfo.length} columns in table '${
        table.id
      }': ${columnInfo.map((c: any) => c.name).join(", ")}`,
    };
  }
);

/**
 * Export all migrated tools
 */
// Unified tool assembly happens in lib/tools/index.ts (domain modules). No grouped export here.
