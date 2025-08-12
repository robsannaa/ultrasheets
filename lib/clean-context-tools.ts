/**
 * Clean Context Tools - No Hardcoding, Pure Intelligence
 *
 * This replaces all hardcoded heuristics with intelligent context discovery
 * LLM gets real data, makes real decisions
 */

export interface CleanSheetContext {
  // Raw Univer API data - no processing, no assumptions
  sheetSnapshot: any;
  selection: {
    activeRange: string | null;
    currentCell: string | null;
    hasSelection: boolean;
    selectionData?: {
      values: any[][];
      formulas: string[][];
      cellData: any[][];
    };
  };

  // Intelligent analysis - but no hardcoded ranges
  analysis: {
    dataRegions: Array<{
      range: string;
      type: "table" | "list" | "sparse";
      headers: string[];
      rowCount: number;
      columnTypes: Record<string, "text" | "number" | "date" | "formula">;
    }>;
    emptyAreas: {
      nextColumns: string[];
      nextRows: number[];
      largestEmptyArea: string;
    };
    formulas: Array<{
      cell: string;
      formula: string;
      dependencies: string[];
    }>;
  };
}

/**
 * Get pure sheet context - no assumptions, just facts
 */
export async function getCleanSheetContext(
  univerAPI: any
): Promise<CleanSheetContext> {
  const workbook = univerAPI.getActiveWorkbook();
  const worksheet = workbook.getActiveSheet();

  // Get raw snapshot - this is the source of truth
  const sheetSnapshot = worksheet.getSheet().getSnapshot();

  // Get current selection without assumptions
  const selection = worksheet.getSelection();
  const activeRange = selection?.getActiveRange();
  const currentCell = selection?.getCurrentCell();

  let selectionData:
    | { values: any[][]; formulas: string[][]; cellData: any[][] }
    | undefined = undefined;
  if (activeRange) {
    try {
      selectionData = {
        values: activeRange.getValues(),
        formulas: activeRange.getFormulas(),
        cellData: activeRange.getCellDatas(),
      };
    } catch (error) {
      console.log("Could not get selection data:", error);
    }
  }

  // Analyze the data intelligently - but based on actual content
  const analysis = analyzeSheetIntelligently(sheetSnapshot);

  return {
    sheetSnapshot,
    selection: {
      activeRange: activeRange?.getA1Notation() || null,
      currentCell: currentCell
        ? worksheet
            .getRange(currentCell.actualRow, currentCell.actualColumn)
            .getA1Notation()
        : null,
      hasSelection: !!activeRange,
      selectionData,
    },
    analysis,
  };
}

/**
 * Intelligent analysis without hardcoded assumptions
 */
function analyzeSheetIntelligently(sheetSnapshot: any) {
  const cellData = sheetSnapshot.cellData || {};

  // Find actual data boundaries
  const { maxRow, maxCol, usedCells } = findActualDataBoundaries(cellData);

  // Detect data regions by analyzing patterns, not assumptions
  const dataRegions = detectDataRegionsIntelligently(cellData, maxRow, maxCol);

  // Find truly empty areas
  const emptyAreas = findEmptyAreas(cellData, maxRow, maxCol, dataRegions);

  // Extract all formulas
  const formulas = extractAllFormulas(cellData);

  return {
    dataRegions,
    emptyAreas,
    formulas,
  };
}

/**
 * Find actual data boundaries - no assumptions
 */
function findActualDataBoundaries(cellData: any) {
  let maxRow = -1;
  let maxCol = -1;
  const usedCells: Array<{ row: number; col: number; value: any }> = [];

  for (const rowStr in cellData) {
    const row = parseInt(rowStr, 10);
    for (const colStr in cellData[row] || {}) {
      const col = parseInt(colStr, 10);
      const cell = cellData[row][col];

      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
        usedCells.push({ row, col, value: cell.v });
      }
    }
  }

  return { maxRow, maxCol, usedCells };
}

/**
 * Detect data regions by analyzing actual patterns
 */
function detectDataRegionsIntelligently(
  cellData: any,
  maxRow: number,
  maxCol: number
) {
  const regions: Array<{
    range: string;
    type: "table" | "list" | "sparse";
    headers: string[];
    rowCount: number;
    columnTypes: Record<string, "text" | "number" | "date" | "formula">;
  }> = [];

  // Look for header patterns (consecutive text cells in a row)
  for (let row = 0; row <= Math.min(maxRow, 20); row++) {
    const rowData = cellData[row] || {};
    const consecutiveHeaders = findConsecutiveHeaders(rowData, maxCol);

    if (consecutiveHeaders.length >= 2) {
      // Found a potential table - analyze its extent
      const tableRegion = analyzeTableRegion(
        cellData,
        row,
        consecutiveHeaders,
        maxRow
      );
      if (tableRegion) {
        regions.push(tableRegion);
      }
    }
  }

  return regions;
}

/**
 * Find consecutive header cells
 */
function findConsecutiveHeaders(rowData: any, maxCol: number) {
  const headers: Array<{ col: number; value: string }> = [];
  let consecutiveStart = -1;

  for (let col = 0; col <= maxCol; col++) {
    const cell = rowData[col];
    const isHeader =
      cell && typeof cell.v === "string" && cell.v.trim() && !cell.f;

    if (isHeader) {
      if (consecutiveStart === -1) consecutiveStart = col;
      headers.push({ col, value: cell.v.trim() });
    } else if (consecutiveStart !== -1) {
      // End of consecutive headers
      break;
    }
  }

  return headers;
}

/**
 * Analyze table region extent and types
 */
function analyzeTableRegion(
  cellData: any,
  headerRow: number,
  headers: any[],
  maxRow: number
) {
  const startCol = headers[0].col;
  const endCol = headers[headers.length - 1].col;

  // Count data rows and detect table boundaries (excluding totals/summary rows)
  let dataRows = 0;
  let lastDataRow = headerRow;

  for (let row = headerRow + 1; row <= maxRow; row++) {
    const rowData = cellData[row] || {};
    let hasData = false;
    let isLikelySummaryRow = false;

    // Check if this row contains data
    for (let col = startCol; col <= endCol; col++) {
      const cell = rowData[col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        hasData = true;

        // Detect if this is likely a summary/total row
        if (typeof cell.v === "string") {
          const cellValue = cell.v.toString().toLowerCase();
          if (
            cellValue.includes("total") ||
            cellValue.includes("sum") ||
            cellValue.includes("subtotal") ||
            cellValue.includes("grand")
          ) {
            isLikelySummaryRow = true;
          }
        }

        // Also check for formulas that span the entire data range (likely totals)
        if (cell.f && cell.f.includes("SUM(") && cell.f.includes(":")) {
          isLikelySummaryRow = true;
        }
      }
    }

    if (hasData) {
      if (isLikelySummaryRow && dataRows > 3) {
        // If we've already found substantial data and hit a summary row, stop here
        // This prevents totals from being included in the data range
        break;
      }
      dataRows++;
      lastDataRow = row;
    } else if (dataRows > 0) {
      // Empty row after data - likely end of table
      break;
    }
  }

  if (dataRows === 0) return null;

  // Analyze column types
  const columnTypes: Record<string, "text" | "number" | "date" | "formula"> =
    {};

  headers.forEach((header) => {
    columnTypes[header.value] = analyzeColumnType(
      cellData,
      header.col,
      headerRow + 1,
      lastDataRow
    );
  });


  const startColLetter = String.fromCharCode(65 + startCol);
  const endColLetter = String.fromCharCode(65 + endCol);
  const range = `${startColLetter}${headerRow + 1}:${endColLetter}${
    lastDataRow + 1
  }`;

  return {
    range,
    type: "table" as const,
    headers: headers.map((h) => h.value),
    rowCount: dataRows,
    columnTypes,
  };
}

/**
 * Analyze column type based on actual data
 */
function analyzeColumnType(
  cellData: any,
  col: number,
  startRow: number,
  endRow: number
): "text" | "number" | "date" | "formula" {
  let numberCount = 0;
  let textCount = 0;
  let formulaCount = 0;
  let dateCount = 0;
  let totalSamples = 0;

  for (let row = startRow; row <= Math.min(endRow, startRow + 10); row++) {
    const cell = cellData[row]?.[col];
    if (!cell || cell.v === undefined || cell.v === null || cell.v === "")
      continue;

    totalSamples++;

    if (cell.f) {
      formulaCount++;
    } else if (typeof cell.v === "number") {
      numberCount++;
    } else if (typeof cell.v === "string") {
      const str = cell.v.toString();
      if (
        /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str) ||
        /^\d{4}-\d{2}-\d{2}$/.test(str)
      ) {
        dateCount++;
      } else if (/^[£$€¥]?\s*[\d,.]+$/.test(str)) {
        numberCount++;
      } else {
        textCount++;
      }
    }
  }

  if (totalSamples === 0) return "text";

  const numberRatio = numberCount / totalSamples;
  const formulaRatio = formulaCount / totalSamples;
  const dateRatio = dateCount / totalSamples;

  if (formulaRatio >= 0.5) return "formula";
  if (dateRatio >= 0.5) return "date";
  if (numberRatio >= 0.5) return "number";
  return "text";
}

/**
 * Find empty areas for intelligent placement
 */
function findEmptyAreas(
  cellData: any,
  maxRow: number,
  maxCol: number,
  dataRegions: any[]
) {
  const nextColumns: string[] = [];
  const nextRows: number[] = [];

  // Find next available columns after each table
  dataRegions.forEach((region) => {
    const match = region.range.match(/([A-Z]+)\d+:([A-Z]+)\d+/);
    if (match) {
      const endCol = match[2];
      const endColIndex = endCol.charCodeAt(endCol.length - 1) - 65;
      const nextColIndex = endColIndex + 1;
      if (nextColIndex <= 25) {
        // Up to column Z
        nextColumns.push(String.fromCharCode(65 + nextColIndex));
      }
    }
  });

  // Find largest empty area
  let largestEmptyArea = `${String.fromCharCode(65 + maxCol + 1)}1`;

  return {
    nextColumns,
    nextRows,
    largestEmptyArea,
  };
}

/**
 * Extract all formulas with dependencies
 */
function extractAllFormulas(cellData: any) {
  const formulas: Array<{
    cell: string;
    formula: string;
    dependencies: string[];
  }> = [];

  for (const rowStr in cellData) {
    const row = parseInt(rowStr, 10);
    for (const colStr in cellData[row] || {}) {
      const col = parseInt(colStr, 10);
      const cell = cellData[row][col];

      if (cell && cell.f) {
        const cellA1 = `${String.fromCharCode(65 + col)}${row + 1}`;
        const dependencies = extractFormulaDependencies(cell.f);

        formulas.push({
          cell: cellA1,
          formula: cell.f,
          dependencies,
        });
      }
    }
  }

  return formulas;
}


/**
 * Extract cell dependencies from formula
 */
function extractFormulaDependencies(formula: string): string[] {
  const dependencies: string[] = [];
  const cellRefPattern = /[A-Z]+\d+/g;
  let match;

  while ((match = cellRefPattern.exec(formula)) !== null) {
    if (!dependencies.includes(match[0])) {
      dependencies.push(match[0]);
    }
  }

  return dependencies;
}
