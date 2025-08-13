/**
 * Univer Data Source - Production Grade
 * 
 * ZERO heuristics, ZERO guessing, ZERO assumptions
 * Uses ONLY Univer API as single source of truth
 * 
 * Designed for financial professionals who require precision
 */

export interface UniverSheetData {
  sheetName: string;
  cells: Array<{
    row: number;
    col: number;
    value: any;
    formula?: string;
    address: string; // e.g., "A1"
  }>;
  usedRange: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
    address: string; // e.g., "A1:F10"
  } | null;
}

export interface UniverWorkbookData {
  workbookId: string;
  activeSheetName: string;
  sheets: UniverSheetData[];
}

/**
 * Extract actual workbook data using ONLY Univer API
 * No assumptions, no guessing, no heuristics
 */
export function getWorkbookData(): UniverWorkbookData | null {
  try {
    const univerAPI = (window as any).univerAPI;
    if (!univerAPI) {
      console.warn("Univer API not available");
      return null;
    }

    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook) {
      console.warn("No active workbook");
      return null;
    }

    const activeSheet = workbook.getActiveSheet();
    if (!activeSheet) {
      console.warn("No active sheet");
      return null;
    }

    // Get workbook snapshot for complete data
    const workbookSnapshot = workbook.getSnapshot();
    const sheets: UniverSheetData[] = [];

    // Process each sheet using Univer API
    const allSheets = workbook.getSheets();
    for (const sheet of allSheets) {
      const sheetSnapshot = sheet.getSheet().getSnapshot();
      const sheetData = extractSheetData(sheet, sheetSnapshot);
      if (sheetData) {
        sheets.push(sheetData);
      }
    }

    return {
      workbookId: workbookSnapshot.id || 'unknown',
      activeSheetName: activeSheet.getSheet().getSnapshot().name || 'Sheet1',
      sheets
    };

  } catch (error) {
    console.error("Error accessing Univer workbook data:", error);
    return null;
  }
}

/**
 * Extract sheet data using ONLY Univer API
 * Returns actual cell values without any interpretation
 */
function extractSheetData(sheet: any, sheetSnapshot: any): UniverSheetData | null {
  try {
    const sheetName = sheetSnapshot.name || 'Unknown Sheet';
    const cells: Array<{
      row: number;
      col: number; 
      value: any;
      formula?: string;
      address: string;
    }> = [];

    const cellData = sheetSnapshot.cellData || {};
    let maxRow = -1;
    let maxCol = -1;
    let minRow = Infinity;
    let minCol = Infinity;

    // Extract actual cell data from Univer snapshot
    for (const rowStr in cellData) {
      const row = parseInt(rowStr, 10);
      const rowData = cellData[row];
      
      for (const colStr in rowData) {
        const col = parseInt(colStr, 10);
        const cellObj = rowData[col];
        
        if (cellObj && (cellObj.v !== undefined || cellObj.f !== undefined)) {
          const address = `${String.fromCharCode(65 + col)}${row + 1}`;
          
          cells.push({
            row,
            col,
            value: cellObj.v,
            formula: cellObj.f,
            address
          });

          // Track actual used range
          maxRow = Math.max(maxRow, row);
          maxCol = Math.max(maxCol, col);
          minRow = Math.min(minRow, row);
          minCol = Math.min(minCol, col);
        }
      }
    }

    // Calculate used range only from actual data
    const usedRange = cells.length > 0 ? {
      startRow: minRow,
      endRow: maxRow,
      startCol: minCol,
      endCol: maxCol,
      address: `${String.fromCharCode(65 + minCol)}${minRow + 1}:${String.fromCharCode(65 + maxCol)}${maxRow + 1}`
    } : null;

    return {
      sheetName,
      cells,
      usedRange
    };

  } catch (error) {
    console.error("Error extracting sheet data:", error);
    return null;
  }
}

/**
 * Get specific cell value using Univer API
 */
export function getCellValue(address: string): any {
  try {
    const univerAPI = (window as any).univerAPI;
    if (!univerAPI) return null;

    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook) return null;

    const sheet = workbook.getActiveSheet();
    if (!sheet) return null;

    const range = sheet.getRange(address);
    return range.getValue();

  } catch (error) {
    console.error(`Error getting cell value for ${address}:`, error);
    return null;
  }
}

/**
 * Get range values using Univer API  
 */
export function getRangeValues(rangeAddress: string): any[][] | null {
  try {
    const univerAPI = (window as any).univerAPI;
    if (!univerAPI) return null;

    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook) return null;

    const sheet = workbook.getActiveSheet();
    if (!sheet) return null;

    const range = sheet.getRange(rangeAddress);
    return range.getValues();

  } catch (error) {
    console.error(`Error getting range values for ${rangeAddress}:`, error);
    return null;
  }
}

/**
 * Get current selection using Univer API
 */
export function getCurrentSelection(): {
  address: string;
  values: any[][];
} | null {
  try {
    const univerAPI = (window as any).univerAPI;
    if (!univerAPI) return null;

    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook) return null;

    const sheet = workbook.getActiveSheet();
    if (!sheet) return null;

    // Get active range (current selection)
    const activeRange = sheet.getActiveRange();
    if (!activeRange) return null;

    const address = activeRange.getAddress();
    const values = activeRange.getValues();

    return { address, values };

  } catch (error) {
    console.error("Error getting current selection:", error);
    return null;
  }
}