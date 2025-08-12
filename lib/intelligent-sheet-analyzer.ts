/**
 * Intelligent Sheet Analyzer - Completely Agnostic System
 *
 * This system can analyze any spreadsheet structure without assumptions:
 * - Multiple tables per sheet
 * - Any column names, positions, data types
 * - Spatial awareness and relationships
 * - Business domain inference
 */

export interface IntelligentTable {
  id: string;
  range: string;
  position: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  };
  headers: string[];
  columnCount: number;
  rowCount: number;

  // COLUMN INTELLIGENCE
  columns: Array<{
    name: string;
    letter: string;
    index: number;
    dataType: "text" | "number" | "date" | "formula" | "currency" | "empty";
    sampleValues: any[];
    hasFormulas: boolean;
    isNumeric: boolean;
    isCurrency: boolean;
    isCalculable: boolean;
  }>;

  // SPATIAL AWARENESS
  spatial: {
    nextAvailableColumn: string;
    nextAvailableRow: number;
    emptySpaceRight: string[];
    emptySpaceBelow: number[];
    canExpandRight: boolean;
    canExpandDown: boolean;
  };

  // SEMANTIC ANALYSIS
  semantics: {
    tableType: "inventory" | "financial" | "customer" | "temporal" | "general";
    businessDomain: "commerce" | "physical" | "evaluation" | "general";
    keyColumns: string[];
    calculableColumns: string[];
    relationships: Array<{
      type: string;
      columns: string[];
    }>;
  };
}

export interface SheetAnalysis {
  tables: IntelligentTable[];
  spatialMap: Array<{
    usedRegions: string[];
    largestEmptyArea: string;
    optimalPlacementZones: Array<{
      type: string;
      tableId?: string;
      position: string;
      description: string;
    }>;
  }>;
  crossTableRelationships: Array<{
    type: string;
    table1: string;
    table2: string;
    commonColumns?: string[];
  }>;
}

export function analyzeSheetIntelligently(cellData: any): SheetAnalysis {
  console.log("🔍 Starting intelligent sheet analysis...");
  console.log("📊 Raw cellData structure:", {
    hasData: !!cellData,
    keys: Object.keys(cellData || {}),
    firstRowData: cellData?.[0] || null,
    sampleCells: Object.keys(cellData || {}).slice(0, 5).map(row => ({
      row,
      data: cellData[row]
    }))
  });

  // BULLETPROOF TABLE DETECTION
  // Layer 1: Standard intelligent detection
  const { maxRow, maxCol } = findDataBoundaries(cellData);
  console.log(`📐 Data boundaries: maxRow=${maxRow}, maxCol=${maxCol}`);
  
  let detectedTables = detectAllTableRegions(cellData, maxRow, maxCol);
  console.log(`🔍 Standard detection found ${detectedTables.length} tables`);

  // Layer 2: Emergency fallback detection for obvious table structures
  if (detectedTables.length === 0) {
    console.warn("⚠️ No tables found with standard detection, using emergency fallback");
    detectedTables = emergencyTableDetection(cellData, maxRow, maxCol);
    console.log(`🚨 Emergency detection found ${detectedTables.length} tables`);
  }

  // Layer 3: Desperate fallback - create table from any data
  if (detectedTables.length === 0 && maxRow >= 0 && maxCol >= 0) {
    console.warn("🆘 Using desperate fallback - creating table from all data");
    detectedTables = desperateFallbackTableCreation(cellData, maxRow, maxCol);
    console.log(`💀 Desperate fallback created ${detectedTables.length} tables`);
  }

  console.log(`✅ Final table count: ${detectedTables.length}`);
  detectedTables.forEach((table, idx) => {
    console.log(`📋 Table ${idx + 1}:`, {
      range: table.range,
      headers: table.headers,
      rowCount: table.rowCount
    });
  });

  const tables: IntelligentTable[] = detectedTables.map((table) => ({
    id: table.range, // Use range as ID to match chat route format (sheetIndex will be added later)
    range: table.range,
    position: {
      startRow: table.startRow,
      endRow: table.endRow,
      startCol: table.startCol,
      endCol: table.endCol,
    },
    headers: table.headers,
    columnCount: table.headers.length,
    rowCount: table.rowCount,

    columns: table.headers.map((header, idx) => {
      const colIndex = table.startCol + idx;
      const dataType = analyzeColumnDataType(
        cellData,
        colIndex,
        table.startRow + 1,
        table.endRow
      );
      const isNumeric = dataType === "number" || dataType === "currency";

      return {
        name: header,
        letter: String.fromCharCode(65 + colIndex),
        index: colIndex,
        dataType,
        sampleValues: extractColumnSample(
          cellData,
          colIndex,
          table.startRow + 1,
          table.endRow
        ),
        hasFormulas: hasFormulasInColumn(
          cellData,
          colIndex,
          table.startRow + 1,
          table.endRow
        ),
        isNumeric,
        isCurrency: isCurrencyColumn(
          cellData,
          colIndex,
          table.startRow + 1,
          table.endRow
        ),
        isCalculable: isNumeric || isCalculableByName(header),
      };
    }),

    spatial: {
      nextAvailableColumn: String.fromCharCode(65 + table.endCol + 1),
      nextAvailableRow: table.endRow + 1,
      emptySpaceRight: findEmptySpaceRight(
        cellData,
        table.endCol + 1,
        table.startRow,
        table.endRow
      ),
      emptySpaceBelow: findEmptySpaceBelow(
        cellData,
        table.startCol,
        table.endCol,
        table.endRow + 1
      ),
      canExpandRight: canExpandRight(
        cellData,
        table.endCol + 1,
        table.startRow,
        table.endRow
      ),
      canExpandDown: canExpandDown(
        cellData,
        table.startCol,
        table.endCol,
        table.endRow + 1
      ),
    },

    semantics: {
      tableType: inferTableType(table.headers),
      businessDomain: inferBusinessDomain(table.headers),
      keyColumns: identifyKeyColumns(table.headers),
      calculableColumns: findCalculableColumns(table.headers),
      relationships: findColumnRelationships(table.headers),
    },
  }));

  const spatialMap = [
    {
      usedRegions: tables.map((t) => t.range),
      largestEmptyArea: findLargestEmptyArea(cellData, maxRow, maxCol, tables),
      optimalPlacementZones: findOptimalPlacementZones(
        cellData,
        maxRow,
        maxCol,
        tables
      ),
    },
  ];

  const crossTableRelationships =
    tables.length > 1 ? findCrossTableRelationships(tables) : [];

  return { tables, spatialMap, crossTableRelationships };
}

// Helper functions
function findDataBoundaries(cellData: any) {
  let maxRow = -1,
    maxCol = -1;
  for (const rowStr in cellData) {
    const row = parseInt(rowStr, 10);
    for (const colStr in cellData[row] || {}) {
      const col = parseInt(colStr, 10);
      const cell = cellData[row][col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
      }
    }
  }
  return { maxRow, maxCol };
}

function detectAllTableRegions(cellData: any, maxRow: number, maxCol: number) {
  const tables = [];
  const processedCells = new Set();

  // Intelligent approach: Find all text cells first (potential headers)
  const textCells = [];
  for (let row = 0; row <= maxRow; row++) {
    const rowData = cellData[row] || {};
    for (let col = 0; col <= maxCol; col++) {
      const cell = rowData[col];
      if (cell && typeof cell.v === "string" && cell.v.trim() && !cell.f) {
        textCells.push({ row, col, value: cell.v.trim() });
      }
    }
  }

  // Group text cells into potential header rows
  const headerRowCandidates = new Map();
  for (const textCell of textCells) {
    if (!headerRowCandidates.has(textCell.row)) {
      headerRowCandidates.set(textCell.row, []);
    }
    headerRowCandidates.get(textCell.row).push(textCell);
  }

  // Analyze each potential header row
  for (const [row, cells] of headerRowCandidates) {
    if (cells.length < 2) continue; // Need at least 2 headers

    // Sort by column to find consecutive headers
    cells.sort((a: any, b: any) => a.col - b.col);

    // Find groups of consecutive headers
    const headerGroups = [];
    let currentGroup = [cells[0]];

    for (let i = 1; i < cells.length; i++) {
      if (cells[i].col === cells[i - 1].col + 1) {
        // Consecutive
        currentGroup.push(cells[i]);
      } else {
        // Gap found
        if (currentGroup.length >= 2) {
          headerGroups.push(currentGroup);
        }
        currentGroup = [cells[i]];
      }
    }
    if (currentGroup.length >= 2) {
      headerGroups.push(currentGroup);
    }

    // Analyze each header group as a potential table
    for (const headerGroup of headerGroups) {
      const startCol = headerGroup[0].col;
      const endCol = headerGroup[headerGroup.length - 1].col;

      // Check if this region is already processed
      const regionKey = `${row}_${startCol}_${endCol}`;
      if (processedCells.has(regionKey)) continue;

      const table = analyzeTableRegion(cellData, row, headerGroup, maxRow);
      if (table && table.rowCount > 0) {
        tables.push(table);
        processedCells.add(regionKey);
      }
    }
  }

  return tables;
}

/**
 * EMERGENCY TABLE DETECTION - Bulletproof Fallback
 * 
 * This function uses aggressive pattern matching to find tables
 * even when the standard detection fails.
 */
function emergencyTableDetection(cellData: any, maxRow: number, maxCol: number) {
  console.log("🚨 Emergency table detection activated");
  const tables = [];
  
  // Strategy 1: Look for any row with multiple consecutive non-empty text cells
  for (let row = 0; row <= Math.min(maxRow, 20); row++) { // Check first 20 rows
    const rowData = cellData[row] || {};
    const textCells = [];
    
    for (let col = 0; col <= maxCol; col++) {
      const cell = rowData[col];
      if (cell && cell.v && typeof cell.v === 'string' && cell.v.trim()) {
        textCells.push({ col, value: cell.v.trim() });
      }
    }
    
    if (textCells.length >= 2) {
      // Found potential header row - check for consecutive cells
      textCells.sort((a, b) => a.col - b.col);
      
      let consecutiveGroup = [textCells[0]];
      for (let i = 1; i < textCells.length; i++) {
        if (textCells[i].col === textCells[i-1].col + 1) {
          consecutiveGroup.push(textCells[i]);
        } else {
          if (consecutiveGroup.length >= 2) {
            // Found a valid header group
            const table = buildEmergencyTable(cellData, row, consecutiveGroup, maxRow);
            if (table) {
              tables.push(table);
              console.log(`✅ Emergency detection found table at row ${row}:`, table.headers);
            }
          }
          consecutiveGroup = [textCells[i]];
        }
      }
      
      // Check final group
      if (consecutiveGroup.length >= 2) {
        const table = buildEmergencyTable(cellData, row, consecutiveGroup, maxRow);
        if (table) {
          tables.push(table);
          console.log(`✅ Emergency detection found table at row ${row}:`, table.headers);
        }
      }
    }
  }
  
  return tables;
}

/**
 * Build table from emergency detection
 */
function buildEmergencyTable(cellData: any, headerRow: number, headers: any[], maxRow: number) {
  const startCol = headers[0].col;
  const endCol = headers[headers.length - 1].col;
  
  // Count data rows below header
  let dataRows = 0;
  let lastDataRow = headerRow;
  
  for (let row = headerRow + 1; row <= maxRow; row++) {
    const rowData = cellData[row] || {};
    let hasAnyData = false;
    
    // More lenient - any data in any column counts
    for (let col = startCol; col <= endCol; col++) {
      const cell = rowData[col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
        hasAnyData = true;
        break;
      }
    }
    
    if (hasAnyData) {
      dataRows++;
      lastDataRow = row;
    } else if (dataRows > 2) { // Need at least 2 data rows to stop
      break;
    }
  }
  
  if (dataRows === 0) return null;
  
  const startColLetter = String.fromCharCode(65 + startCol);
  const endColLetter = String.fromCharCode(65 + endCol);
  const range = `${startColLetter}${headerRow + 1}:${endColLetter}${lastDataRow + 1}`;
  
  return {
    range,
    startRow: headerRow,
    endRow: lastDataRow,
    startCol,
    endCol,
    headers: headers.map(h => h.value),
    rowCount: dataRows,
  };
}

/**
 * DESPERATE FALLBACK - Create table from any available data
 * 
 * This creates a basic table structure even if no clear headers are found.
 */
function desperateFallbackTableCreation(cellData: any, maxRow: number, maxCol: number) {
  console.log("💀 Desperate fallback table creation activated");
  
  // Find the first row with data and use it as headers
  for (let row = 0; row <= Math.min(maxRow, 10); row++) {
    const rowData = cellData[row] || {};
    const cellsInRow = [];
    
    for (let col = 0; col <= maxCol; col++) {
      const cell = rowData[col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
        cellsInRow.push({
          col,
          value: cell.v.toString().trim() || `Column ${String.fromCharCode(65 + col)}`
        });
      }
    }
    
    if (cellsInRow.length >= 2) {
      // Use this as a header row and count data below
      let dataRows = 0;
      let lastDataRow = row;
      
      for (let dataRow = row + 1; dataRow <= maxRow; dataRow++) {
        const dataRowData = cellData[dataRow] || {};
        let hasData = false;
        
        for (const headerCell of cellsInRow) {
          const cell = dataRowData[headerCell.col];
          if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
            hasData = true;
            break;
          }
        }
        
        if (hasData) {
          dataRows++;
          lastDataRow = dataRow;
        }
      }
      
      if (dataRows > 0) {
        const startCol = Math.min(...cellsInRow.map(c => c.col));
        const endCol = Math.max(...cellsInRow.map(c => c.col));
        const startColLetter = String.fromCharCode(65 + startCol);
        const endColLetter = String.fromCharCode(65 + endCol);
        const range = `${startColLetter}${row + 1}:${endColLetter}${lastDataRow + 1}`;
        
        console.log(`💀 Created desperate fallback table: ${range}`);
        
        return [{
          range,
          startRow: row,
          endRow: lastDataRow,
          startCol,
          endCol,
          headers: cellsInRow.map(c => c.value),
          rowCount: dataRows,
        }];
      }
    }
  }
  
  // Ultimate fallback - create generic table if ANY data exists
  if (maxRow >= 0 && maxCol >= 0) {
    const range = `A1:${String.fromCharCode(65 + Math.min(maxCol, 25))}${Math.min(maxRow + 1, 100)}`;
    const headers = [];
    for (let col = 0; col <= Math.min(maxCol, 25); col++) {
      headers.push(`Column ${String.fromCharCode(65 + col)}`);
    }
    
    console.log(`💀 Created ultimate fallback table: ${range}`);
    
    return [{
      range,
      startRow: 0,
      endRow: Math.min(maxRow, 99),
      startCol: 0,
      endCol: Math.min(maxCol, 25),
      headers,
      rowCount: Math.min(maxRow + 1, 100),
    }];
  }
  
  return [];
}

// findConsecutiveHeaders removed - now using intelligent detection above

function analyzeTableRegion(
  cellData: any,
  headerRow: number,
  headers: any[],
  maxRow: number
) {
  const startCol = headers[0].col;
  const endCol = headers[headers.length - 1].col;

  // Intelligently find the actual data extent below headers
  let dataRows = 0;
  let lastDataRow = headerRow;

  for (let row = headerRow + 1; row <= maxRow; row++) {
    const rowData = cellData[row] || {};
    let hasData = false;

    // Check if this row has data in any of the table columns
    for (let col = startCol; col <= endCol; col++) {
      const cell = rowData[col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        hasData = true;
        break;
      }
    }

    if (hasData) {
      dataRows++;
      lastDataRow = row;
    } else if (dataRows > 0) {
      // Found empty row after data - check if it's really the end
      // Look ahead a few rows to see if there's more data (handles sparse tables)
      let hasDataAhead = false;
      for (
        let lookAhead = row + 1;
        lookAhead <= Math.min(row + 3, maxRow);
        lookAhead++
      ) {
        const lookAheadRowData = cellData[lookAhead] || {};
        for (let col = startCol; col <= endCol; col++) {
          const cell = lookAheadRowData[col];
          if (
            cell &&
            cell.v !== undefined &&
            cell.v !== null &&
            cell.v !== ""
          ) {
            hasDataAhead = true;
            break;
          }
        }
        if (hasDataAhead) break;
      }

      if (!hasDataAhead) {
        break; // Truly end of table
      }
    }
  }

  if (dataRows === 0) return null;

  const startColLetter = String.fromCharCode(65 + startCol);
  const endColLetter = String.fromCharCode(65 + endCol);
  const range = `${startColLetter}${headerRow + 1}:${endColLetter}${
    lastDataRow + 1
  }`;

  return {
    range,
    startRow: headerRow,
    endRow: lastDataRow,
    startCol,
    endCol,
    headers: headers.map((h) => h.value),
    rowCount: dataRows,
  };
}

function analyzeColumnDataType(
  cellData: any,
  col: number,
  startRow: number,
  endRow: number
) {
  const samples = [];
  for (let row = startRow; row <= Math.min(endRow, startRow + 10); row++) {
    const cell = cellData[row]?.[col];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
      samples.push(cell);
    }
  }

  if (samples.length === 0) return "empty";

  let numberCount = 0,
    dateCount = 0,
    formulaCount = 0,
    currencyCount = 0;
  for (const cell of samples) {
    if (cell.f) formulaCount++;
    else if (typeof cell.v === "number") numberCount++;
    else if (typeof cell.v === "string") {
      const str = cell.v.toString();
      if (
        /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str) ||
        /^\d{4}-\d{2}-\d{2}$/.test(str)
      ) {
        dateCount++;
      } else if (/^[£$€¥]\s*[\d,.]+$/.test(str)) {
        currencyCount++;
      } else if (/^[\d,.]+$/.test(str)) {
        numberCount++;
      }
    }
  }

  const total = samples.length;
  if (formulaCount / total >= 0.5) return "formula";
  if (currencyCount / total >= 0.3) return "currency";
  if (dateCount / total >= 0.5) return "date";
  if (numberCount / total >= 0.5) return "number";
  return "text";
}

function extractColumnSample(
  cellData: any,
  col: number,
  startRow: number,
  endRow: number
) {
  const samples = [];
  for (let row = startRow; row <= Math.min(endRow, startRow + 3); row++) {
    const cell = cellData[row]?.[col];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
      samples.push(cell.v);
    }
  }
  return samples;
}

function hasFormulasInColumn(
  cellData: any,
  col: number,
  startRow: number,
  endRow: number
) {
  for (let row = startRow; row <= Math.min(endRow, startRow + 5); row++) {
    const cell = cellData[row]?.[col];
    if (cell?.f) return true;
  }
  return false;
}

function isCurrencyColumn(
  cellData: any,
  col: number,
  startRow: number,
  endRow: number
) {
  const samples = extractColumnSample(cellData, col, startRow, endRow);
  return samples.some(
    (v) => typeof v === "string" && /[£$€¥]/.test(v.toString())
  );
}

function isCalculableByName(header: string) {
  const lower = header.toLowerCase();
  return (
    lower.includes("price") ||
    lower.includes("cost") ||
    lower.includes("amount") ||
    lower.includes("weight") ||
    lower.includes("quantity") ||
    lower.includes("total") ||
    lower.includes("sum") ||
    lower.includes("value")
  );
}

function findEmptySpaceRight(
  cellData: any,
  startCol: number,
  startRow: number,
  endRow: number
) {
  const emptyColumns = [];
  for (let col = startCol; col <= startCol + 5; col++) {
    let isEmpty = true;
    for (let row = startRow; row <= endRow; row++) {
      const cell = cellData[row]?.[col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        isEmpty = false;
        break;
      }
    }
    if (isEmpty) emptyColumns.push(String.fromCharCode(65 + col));
    else break;
  }
  return emptyColumns;
}

function findEmptySpaceBelow(
  cellData: any,
  startCol: number,
  endCol: number,
  startRow: number
) {
  const emptyRows = [];
  for (let row = startRow; row <= startRow + 10; row++) {
    let isEmpty = true;
    for (let col = startCol; col <= endCol; col++) {
      const cell = cellData[row]?.[col];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        isEmpty = false;
        break;
      }
    }
    if (isEmpty) emptyRows.push(row + 1);
    else break;
  }
  return emptyRows;
}

function canExpandRight(
  cellData: any,
  col: number,
  startRow: number,
  endRow: number
) {
  for (let row = startRow; row <= endRow; row++) {
    const cell = cellData[row]?.[col];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
      return false;
    }
  }
  return true;
}

function canExpandDown(
  cellData: any,
  startCol: number,
  endCol: number,
  row: number
) {
  for (let col = startCol; col <= endCol; col++) {
    const cell = cellData[row]?.[col];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
      return false;
    }
  }
  return true;
}

function inferTableType(headers: string[]) {
  const h = headers.map((h) => h.toLowerCase());
  if (h.some((h) => h.includes("product") || h.includes("item")))
    return "inventory";
  if (h.some((h) => h.includes("sales") || h.includes("revenue")))
    return "financial";
  if (h.some((h) => h.includes("customer") || h.includes("client")))
    return "customer";
  if (h.some((h) => h.includes("date") || h.includes("time")))
    return "temporal";
  return "general";
}

function inferBusinessDomain(headers: string[]) {
  const h = headers.map((h) => h.toLowerCase());
  if (
    h.some(
      (h) => h.includes("price") || h.includes("cost") || h.includes("revenue")
    )
  )
    return "commerce";
  if (h.some((h) => h.includes("weight") || h.includes("size")))
    return "physical";
  if (h.some((h) => h.includes("rating") || h.includes("score")))
    return "evaluation";
  return "general";
}

function identifyKeyColumns(headers: string[]) {
  return headers.filter((h, i) => {
    const lower = h.toLowerCase();
    return (
      lower.includes("id") ||
      lower.includes("name") ||
      lower.includes("product") ||
      i === 0
    );
  });
}

function findCalculableColumns(headers: string[]) {
  return headers.filter((h) => isCalculableByName(h));
}

function findColumnRelationships(headers: string[]) {
  const relationships = [];
  const h = headers.map((h) => h.toLowerCase());

  const priceIdx = h.findIndex((h) => h.includes("price"));
  const costIdx = h.findIndex((h) => h.includes("cost"));
  if (priceIdx >= 0 && costIdx >= 0) {
    relationships.push({
      type: "profit_margin",
      columns: [headers[priceIdx], headers[costIdx]],
    });
  }

  const qtyIdx = h.findIndex(
    (h) => h.includes("quantity") || h.includes("qty")
  );
  const weightIdx = h.findIndex((h) => h.includes("weight"));
  if (qtyIdx >= 0 && weightIdx >= 0) {
    relationships.push({
      type: "per_unit_weight",
      columns: [headers[weightIdx], headers[qtyIdx]],
    });
  }

  return relationships;
}

function findLargestEmptyArea(
  cellData: any,
  maxRow: number,
  maxCol: number,
  tables: any[]
) {
  let rightmostCol = 0,
    bottommostRow = 0;

  for (const table of tables) {
    rightmostCol = Math.max(rightmostCol, table.position.endCol);
    bottommostRow = Math.max(bottommostRow, table.position.endRow);
  }

  const nextCol = String.fromCharCode(65 + rightmostCol + 2);
  const nextRow = bottommostRow + 3;

  return `${nextCol}${nextRow}`;
}

function findOptimalPlacementZones(
  cellData: any,
  maxRow: number,
  maxCol: number,
  tables: any[]
) {
  const zones = [];

  for (const table of tables) {
    if (table.spatial.canExpandRight) {
      zones.push({
        type: "right_of_table",
        tableId: table.id,
        position:
          table.spatial.nextAvailableColumn + (table.position.startRow + 1),
        description: `Right of ${table.id}`,
      });
    }

    if (table.spatial.canExpandDown) {
      zones.push({
        type: "below_table",
        tableId: table.id,
        position:
          String.fromCharCode(65 + table.position.startCol) +
          (table.position.endRow + 2),
        description: `Below ${table.id}`,
      });
    }
  }

  return zones;
}

function findCrossTableRelationships(tables: any[]) {
  const relationships = [];

  for (let i = 0; i < tables.length; i++) {
    for (let j = i + 1; j < tables.length; j++) {
      const table1 = tables[i];
      const table2 = tables[j];

      const commonColumns = table1.headers.filter((h1: string) =>
        table2.headers.some(
          (h2: string) => h1.toLowerCase() === h2.toLowerCase()
        )
      );

      if (commonColumns.length > 0) {
        relationships.push({
          type: "shared_columns",
          table1: table1.id,
          table2: table2.id,
          commonColumns,
        });
      }
    }
  }

  return relationships;
}
