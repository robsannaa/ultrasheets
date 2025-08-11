/**
 * Enhanced Context System for UltraSheets
 * Provides intelligent table detection, spatial awareness, and selection context
 */

export interface TableStructure {
  range: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  headers: string[];
  recordCount: number;
  numericColumns: string[];
  nextAvailableColumn: string; // Key addition!
  nextAvailableRow: number;
  spatialContext: {
    hasSpaceRight: boolean;
    hasSpaceBelow: boolean;
    emptyColumnsRight: string[];
    emptyRowsBelow: number[];
  };
}

export interface SelectionContext {
  hasSelection: boolean;
  activeRange: string | null;
  selectedTable: TableStructure | null;
  selectionIntent: "add_column" | "add_row" | "modify_data" | "unknown";
  currentCell: {
    row: number;
    column: number;
    a1: string;
  } | null;
}

export interface LiveSelectionData {
  values: any[][];
  formulas: string[][];
  displayValues: string[][];
  cellData: any[][];
  isTableLike: boolean;
  hasHeaders: boolean;
  suggestedOperations: string[];
}

export interface EnhancedSheetContext {
  name: string;
  isActive: boolean;
  tables: TableStructure[];
  selection: SelectionContext;
  liveSelection?: LiveSelectionData;
  spatialMap: {
    usedRange: string;
    freeColumns: string[];
    freeRows: number[];
    nextSafeColumn: string;
    nextSafeRow: number;
  };
  intelligentSuggestions: {
    recommendedActions: string[];
    contextualHints: string[];
    potentialIssues: string[];
  };
}

/**
 * Enhanced table detection with spatial awareness
 */
export function detectTablesWithSpatialContext(
  cellData: any
): TableStructure[] {
  const tables: TableStructure[] = [];
  if (!cellData) return tables;

  // Get all used rows and columns
  const rows = Object.keys(cellData)
    .map((k) => parseInt(k, 10))
    .sort((a, b) => a - b);
  const maxRow = rows.length ? rows[rows.length - 1] : 0;
  let maxCol = 0;

  for (const r of rows) {
    const cols = Object.keys(cellData[r] || {}).map((k) => parseInt(k, 10));
    if (cols.length) maxCol = Math.max(maxCol, cols[cols.length - 1]);
  }

  // Find header rows (rows with consecutive text strings)
  const headerCandidates: number[] = [];
  for (let r = 0; r <= Math.min(maxRow, 30); r++) {
    const rowData = cellData[r] || {};
    let consecutiveHeaders = 0;
    let firstHeaderCol = -1;
    let lastHeaderCol = -1;

    for (let c = 0; c <= Math.min(maxCol, 50); c++) {
      const cell = rowData[c];
      if (cell && typeof cell.v === "string" && cell.v.trim() && !cell.f) {
        if (firstHeaderCol === -1) firstHeaderCol = c;
        lastHeaderCol = c;
        consecutiveHeaders++;
      } else if (consecutiveHeaders > 0) {
        break; // Stop at first empty cell after headers
      }
    }

    if (consecutiveHeaders >= 2) {
      headerCandidates.push(r);
    }
  }

  // Process each detected table
  for (const headerRow of headerCandidates) {
    const rowData = cellData[headerRow] || {};
    const headers: string[] = [];
    let firstCol = -1;
    let lastCol = -1;

    // Extract headers
    for (let c = 0; c <= Math.min(maxCol, 50); c++) {
      const cell = rowData[c];
      if (cell && typeof cell.v === "string" && cell.v.trim() && !cell.f) {
        headers.push(String(cell.v).trim());
        if (firstCol === -1) firstCol = c;
        lastCol = c;
      } else if (headers.length > 0) {
        break; // Stop at first empty cell
      }
    }

    if (headers.length < 2) continue;

    // Count data rows
    let recordCount = 0;
    let actualEndRow = headerRow;

    for (let r = headerRow + 1; r <= Math.min(headerRow + 1000, maxRow); r++) {
      const rd = cellData[r] || {};
      let hasData = false;

      for (let c = firstCol; c <= lastCol; c++) {
        const cell = rd[c];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
          hasData = true;
          break;
        }
      }

      if (hasData) {
        recordCount++;
        actualEndRow = r;
      } else if (recordCount > 0) {
        // Stop at first completely empty row after data
        break;
      }
    }

    // Calculate spatial context
    const nextAvailableCol = lastCol + 1;
    const nextAvailableColumn = String.fromCharCode(65 + nextAvailableCol);
    const nextAvailableRow = actualEndRow + 1;

    // Check available space
    const hasSpaceRight = nextAvailableCol <= 25; // Up to column Z
    const hasSpaceBelow = nextAvailableRow <= maxRow + 100;

    // Find empty columns to the right
    const emptyColumnsRight: string[] = [];
    for (
      let c = nextAvailableCol;
      c <= Math.min(nextAvailableCol + 10, 25);
      c++
    ) {
      let isEmpty = true;
      for (let r = headerRow; r <= actualEndRow + 5; r++) {
        const cell = (cellData[r] || {})[c];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        emptyColumnsRight.push(String.fromCharCode(65 + c));
      }
    }

    // Detect numeric columns
    const numericColumns: string[] = [];
    for (let c = firstCol; c <= lastCol; c++) {
      let numericHits = 0;
      const sampleSize = Math.min(5, recordCount);

      for (let r = headerRow + 1; r <= headerRow + sampleSize; r++) {
        const cell = (cellData[r] || {})[c];
        if (!cell) continue;

        const v = cell.v;
        if (
          typeof v === "number" ||
          (typeof v === "string" && /^[£$€¥]?\s*[\d,.]+$/.test(v.toString()))
        ) {
          numericHits++;
        }
      }

      if (numericHits >= Math.max(2, sampleSize * 0.6)) {
        numericColumns.push(String.fromCharCode(65 + c));
      }
    }

    const startColLetter = String.fromCharCode(65 + firstCol);
    const endColLetter = String.fromCharCode(65 + lastCol);
    const range = `${startColLetter}${headerRow + 1}:${endColLetter}${
      actualEndRow + 1
    }`;

    tables.push({
      range,
      startRow: headerRow,
      endRow: actualEndRow,
      startCol: firstCol,
      endCol: lastCol,
      headers,
      recordCount,
      numericColumns,
      nextAvailableColumn,
      nextAvailableRow,
      spatialContext: {
        hasSpaceRight,
        hasSpaceBelow,
        emptyColumnsRight,
        emptyRowsBelow: [], // Can be populated if needed
      },
    });
  }

  return tables;
}

/**
 * Analyze live selection data for intelligent suggestions
 */
export function analyzeLiveSelectionData(
  univerAPI: any,
  activeRange: any
): LiveSelectionData | undefined {
  try {
    if (!activeRange) return undefined;

    const values = activeRange.getValues();
    const formulas = activeRange.getFormulas();
    const displayValues = activeRange.getDisplayValues();
    const cellData = activeRange.getCellDatas();

    // Analyze if selection looks like a table
    const isTableLike = values.length > 1 && values[0]?.length > 1;
    const hasHeaders =
      isTableLike &&
      values[0]?.every((v: any) => typeof v === "string" && v.trim());

    // Generate intelligent suggestions based on selection content
    const suggestedOperations: string[] = [];

    if (hasHeaders && values.length > 2) {
      suggestedOperations.push(
        "create_pivot_table",
        "generate_chart",
        "add_filter"
      );
    }

    // Check for numeric data patterns
    const hasNumericData = values
      .slice(1)
      .some((row: any[]) => row.some((cell) => typeof cell === "number"));
    if (hasNumericData) {
      suggestedOperations.push(
        "calculate_total",
        "format_currency",
        "conditional_formatting"
      );
    }

    // Check for potential lookup scenarios
    if (hasHeaders && values[0]?.length >= 2) {
      suggestedOperations.push("vlookup_setup", "index_match_setup");
    }

    return {
      values,
      formulas,
      displayValues,
      cellData,
      isTableLike,
      hasHeaders,
      suggestedOperations,
    };
  } catch (error) {
    console.error("Error analyzing live selection:", error);
    return undefined;
  }
}

/**
 * Generate intelligent suggestions based on current context
 */
export function generateIntelligentSuggestions(
  tables: TableStructure[],
  selection: SelectionContext,
  liveSelection: LiveSelectionData | undefined
): {
  recommendedActions: string[];
  contextualHints: string[];
  potentialIssues: string[];
} {
  const recommendedActions: string[] = [];
  const contextualHints: string[] = [];
  const potentialIssues: string[] = [];

  // Analyze table context
  if (tables.length === 0) {
    contextualHints.push(
      "No structured data tables detected. Consider organizing your data with headers."
    );
  } else if (tables.length > 1) {
    contextualHints.push(
      `${tables.length} data tables detected. Specify which table for operations.`
    );
  }

  // Analyze selection context
  if (selection.hasSelection && selection.selectedTable) {
    const table = selection.selectedTable;

    switch (selection.selectionIntent) {
      case "add_column":
        recommendedActions.push("smart_add_column");
        contextualHints.push(
          `Ready to add column to table ${table.range} in column ${table.nextAvailableColumn}`
        );
        break;
      case "add_row":
        recommendedActions.push("insert_row");
        contextualHints.push(
          `Ready to add row to table ${table.range} at row ${table.nextAvailableRow}`
        );
        break;
      case "modify_data":
        recommendedActions.push(
          "bulk_edit",
          "format_cells",
          "conditional_formatting"
        );
        break;
    }

    // Check for potential issues
    if (
      !table.spatialContext.hasSpaceRight &&
      selection.selectionIntent === "add_column"
    ) {
      potentialIssues.push(
        "Limited space to the right. Consider inserting column or using a different location."
      );
    }
  }

  // Analyze live selection data
  if (liveSelection) {
    recommendedActions.push(...liveSelection.suggestedOperations);

    if (liveSelection.isTableLike && !liveSelection.hasHeaders) {
      contextualHints.push(
        "Selected data appears tabular but lacks headers. Consider adding column headers."
      );
    }
  }

  return { recommendedActions, contextualHints, potentialIssues };
}

/**
 * Analyze current selection to understand user intent
 */
export function analyzeSelectionIntent(
  univerAPI: any,
  tables: TableStructure[]
): SelectionContext {
  try {
    const workbook = univerAPI.getActiveWorkbook();
    const worksheet = workbook?.getActiveSheet();
    if (!worksheet) {
      return {
        hasSelection: false,
        activeRange: null,
        selectedTable: null,
        selectionIntent: "unknown",
        currentCell: null,
      };
    }

    const selection = worksheet.getSelection();
    const activeRange = selection?.getActiveRange();
    const currentCell = selection?.getCurrentCell();

    if (!activeRange) {
      return {
        hasSelection: false,
        activeRange: null,
        selectedTable: null,
        selectionIntent: "unknown",
        currentCell: currentCell
          ? {
              row: currentCell.actualRow,
              column: currentCell.actualColumn,
              a1: worksheet
                .getRange(currentCell.actualRow, currentCell.actualColumn)
                .getA1Notation(),
            }
          : null,
      };
    }

    const activeRangeNotation = activeRange.getA1Notation();

    // Find which table (if any) the selection relates to
    let selectedTable: TableStructure | null = null;
    let selectionIntent: SelectionContext["selectionIntent"] = "unknown";

    for (const table of tables) {
      // Check if selection is within or adjacent to table
      const selectionStart = parseA1Notation(activeRangeNotation.split(":")[0]);
      const selectionEnd = parseA1Notation(
        activeRangeNotation.split(":")[1] || activeRangeNotation.split(":")[0]
      );

      if (!selectionStart || !selectionEnd) continue;

      // Selection within table
      if (
        selectionStart.row >= table.startRow &&
        selectionStart.row <= table.endRow &&
        selectionStart.column >= table.startCol &&
        selectionStart.column <= table.endCol
      ) {
        selectedTable = table;
        selectionIntent = "modify_data";
        break;
      }

      // Selection in next available column (likely wants to add column)
      if (
        selectionStart.column === table.endCol + 1 &&
        selectionStart.row >= table.startRow &&
        selectionStart.row <= table.endRow
      ) {
        selectedTable = table;
        selectionIntent = "add_column";
        break;
      }

      // Selection in next available row (likely wants to add row)
      if (
        selectionStart.row === table.endRow + 1 &&
        selectionStart.column >= table.startCol &&
        selectionStart.column <= table.endCol
      ) {
        selectedTable = table;
        selectionIntent = "add_row";
        break;
      }
    }

    return {
      hasSelection: true,
      activeRange: activeRangeNotation,
      selectedTable,
      selectionIntent,
      currentCell: currentCell
        ? {
            row: currentCell.actualRow,
            column: currentCell.actualColumn,
            a1: worksheet
              .getRange(currentCell.actualRow, currentCell.actualColumn)
              .getA1Notation(),
          }
        : null,
    };
  } catch (error) {
    console.error("Error analyzing selection:", error);
    return {
      hasSelection: false,
      activeRange: null,
      selectedTable: null,
      selectionIntent: "unknown",
      currentCell: null,
    };
  }
}

/**
 * Parse A1 notation to row/column numbers
 */
function parseA1Notation(a1: string): { row: number; column: number } | null {
  const match = a1.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  const [, colStr, rowStr] = match;
  let column = 0;

  for (let i = 0; i < colStr.length; i++) {
    column = column * 26 + (colStr.charCodeAt(i) - 64);
  }

  return {
    row: parseInt(rowStr, 10) - 1, // Convert to 0-based
    column: column - 1, // Convert to 0-based
  };
}

/**
 * Enhanced workbook context extraction
 */
export function extractEnhancedWorkbookData(): {
  sheets: EnhancedSheetContext[];
  recentActions: any[];
} | null {
  try {
    if (typeof window === "undefined" || !(window as any).univerAPI) {
      return null;
    }

    const univerAPI = (window as any).univerAPI;
    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook || typeof workbook.save !== "function") return null;

    const activeSheet = workbook.getActiveSheet();
    const activeSnapshot = activeSheet?.getSheet()?.getSnapshot();
    const activeName = activeSnapshot?.name;

    const wbData = workbook.save();
    const sheetOrder: string[] = wbData.sheetOrder || [];
    const sheetsData: Record<string, any> = wbData.sheets || {};

    const sheets: EnhancedSheetContext[] = sheetOrder.map((sid) => {
      const s = sheetsData[sid];
      const name = s?.name || "Sheet";
      const cellData = s?.cellData || {};
      const isActive = name === activeName;

      // Enhanced table detection
      const tables = detectTablesWithSpatialContext(cellData);

      // Selection analysis (only for active sheet)
      let selection: SelectionContext;
      let liveSelection: LiveSelectionData | undefined = undefined;

      if (isActive) {
        selection = analyzeSelectionIntent(univerAPI, tables);

        // Get live selection data if there's an active selection
        if (selection.hasSelection) {
          try {
            const activeSheet = univerAPI.getActiveWorkbook().getActiveSheet();
            const activeRange = activeSheet.getSelection()?.getActiveRange();
            liveSelection = analyzeLiveSelectionData(univerAPI, activeRange);
          } catch (error) {
            console.error("Error getting live selection data:", error);
          }
        }
      } else {
        selection = {
          hasSelection: false,
          activeRange: null,
          selectedTable: null,
          selectionIntent: "unknown" as const,
          currentCell: null,
        };
      }

      // Calculate spatial map
      let maxRowUsed = -1;
      let maxColUsed = -1;

      for (const r in cellData) {
        for (const c in cellData[r]) {
          const cell = cellData[r][c];
          if (
            cell &&
            cell.v !== undefined &&
            cell.v !== null &&
            cell.v !== ""
          ) {
            const ri = parseInt(r, 10);
            const ci = parseInt(c, 10);
            if (ri > maxRowUsed) maxRowUsed = ri;
            if (ci > maxColUsed) maxColUsed = ci;
          }
        }
      }

      const usedRange =
        maxRowUsed >= 0 && maxColUsed >= 0
          ? `A1:${String.fromCharCode(65 + maxColUsed)}${maxRowUsed + 1}`
          : "A1:A1";

      // Find globally safe positions
      const freeColumns: string[] = [];
      for (let c = maxColUsed + 1; c <= Math.min(maxColUsed + 10, 25); c++) {
        freeColumns.push(String.fromCharCode(65 + c));
      }

      const nextSafeColumn = freeColumns[0] || "Z";
      const nextSafeRow = maxRowUsed + 2;

      // Generate intelligent suggestions
      const intelligentSuggestions = generateIntelligentSuggestions(
        tables,
        selection,
        liveSelection
      );

      return {
        name,
        isActive,
        tables,
        selection,
        liveSelection,
        spatialMap: {
          usedRange,
          freeColumns,
          freeRows: [], // Can be calculated if needed
          nextSafeColumn,
          nextSafeRow,
        },
        intelligentSuggestions,
      };
    });

    // Get recent actions
    const recentActions = (window as any).ultraActionLog || [];

    return { sheets, recentActions };
  } catch (error) {
    console.error("Error extracting enhanced workbook data:", error);
    return null;
  }
}
