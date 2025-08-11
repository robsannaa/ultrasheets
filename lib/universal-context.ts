/**
 * Universal Context Manager for Univer AI
 *
 * Provides centralized, cached, and intelligent context for all spreadsheet operations.
 * This eliminates code duplication and ensures consistent state across all tools.
 */

import {
  analyzeSheetIntelligently,
  type SheetAnalysis,
} from "./intelligent-sheet-analyzer";

export interface UniversalWorkbookContext {
  // Raw Univer API objects
  univerAPI: any;
  fWorkbook: any;
  fWorksheet: any;

  // Current state
  activeSheetName: string;
  activeSheetSnapshot: any;

  // Intelligent analysis (cached)
  intelligence: SheetAnalysis;

  // Quick accessors
  tables: any[];
  primaryTable: any | null;
  calculableColumns: any[];
  numericColumns: any[];
  spatialMap: any;

  // Metadata
  lastUpdated: number;
  cacheValid: boolean;
}

export interface UniversalToolContext extends UniversalWorkbookContext {
  // Tool-specific helpers
  findTable: (tableId?: string) => any | null;
  findColumn: (columnName: string, tableId?: string) => any | null;
  getTableRange: (tableId?: string) => string;
  getColumnRange: (
    columnName: string,
    includeHeader?: boolean,
    tableId?: string
  ) => string;

  // Smart range builders
  buildSumFormula: (columnName: string, tableId?: string) => string;
  findOptimalPlacement: (
    width: number,
    height: number
  ) => { row: number; col: number; range: string };

  // State helpers
  invalidateCache: () => void;
  refresh: () => Promise<UniversalToolContext>;
}

class UniversalContextManager {
  private static instance: UniversalContextManager;
  private cachedContext: UniversalWorkbookContext | null = null;
  private cacheTimeout = 5000; // 5 seconds

  private constructor() {}

  public static getInstance(): UniversalContextManager {
    if (!UniversalContextManager.instance) {
      UniversalContextManager.instance = new UniversalContextManager();
    }
    return UniversalContextManager.instance;
  }

  /**
   * Get universal context with intelligent caching
   */
  public async getContext(forceRefresh = false): Promise<UniversalToolContext> {
    // Check cache validity
    if (!forceRefresh && this.cachedContext && this.isCacheValid()) {
      return this.enrichContext(this.cachedContext);
    }

    // Build fresh context
    const baseContext = await this.buildBaseContext();
    this.cachedContext = baseContext;

    return this.enrichContext(baseContext);
  }

  /**
   * Brute force data extraction using every possible method
   */
  private async bruteForceDataExtraction(fWorksheet: any): Promise<any> {
    const extractedData: any = {};
    let cellsFound = 0;

    console.log("🔨 Starting brute force data extraction...");

    // Method 1: Large range scan
    try {
      const largeRange = fWorksheet.getRange("A1:Z200");
      const values = largeRange.getValues();
      if (values && Array.isArray(values)) {
        values.forEach((row: any[], rowIndex: number) => {
          if (Array.isArray(row)) {
            row.forEach((cellValue: any, colIndex: number) => {
              if (cellValue != null && cellValue !== '') {
                if (!extractedData[rowIndex]) {
                  extractedData[rowIndex] = {};
                }
                extractedData[rowIndex][colIndex] = {
                  v: cellValue,
                  t: typeof cellValue === 'number' ? 1 : 0
                };
                cellsFound++;
              }
            });
          }
        });
      }
    } catch (error) {
      console.log("🔨 Large range method failed:", error);
    }

    // Method 2: Grid-based extraction
    if (cellsFound === 0) {
      console.log("🔨 Trying grid-based extraction...");
      for (let row = 0; row < 100; row++) {
        for (let col = 0; col < 50; col++) {
          try {
            const colLetter = this.columnToLetter(col);
            const cellRef = `${colLetter}${row + 1}`;
            const cell = fWorksheet.getRange(cellRef);
            const value = cell.getValue();
            
            if (value != null && value !== '') {
              if (!extractedData[row]) {
                extractedData[row] = {};
              }
              extractedData[row][col] = {
                v: value,
                t: typeof value === 'number' ? 1 : 0
              };
              cellsFound++;
            }
          } catch {} // Ignore individual cell errors
        }
      }
    }

    // Method 3: Enhanced Univer API approaches
    if (cellsFound === 0) {
      console.log("🔨 Trying enhanced Univer API methods...");
      
      // Try getting worksheet snapshot and extracting data differently
      try {
        const sheetSnapshot = fWorksheet.getSheet().getSnapshot();
        if (sheetSnapshot && sheetSnapshot.cellData) {
          const snapshotCellData = sheetSnapshot.cellData;
          console.log("🔨 Found snapshot cellData:", Object.keys(snapshotCellData).length, "rows");
          
          // Process snapshot data
          Object.keys(snapshotCellData).forEach(rowKey => {
            const row = parseInt(rowKey);
            const rowData = snapshotCellData[row];
            if (rowData && typeof rowData === 'object') {
              Object.keys(rowData).forEach(colKey => {
                const col = parseInt(colKey);
                const cellData = rowData[col];
                if (cellData && cellData.v != null && cellData.v !== '') {
                  if (!extractedData[row]) {
                    extractedData[row] = {};
                  }
                  extractedData[row][col] = {
                    v: cellData.v,
                    t: cellData.t || (typeof cellData.v === 'number' ? 1 : 0)
                  };
                  cellsFound++;
                }
              });
            }
          });
        }
      } catch (error) {
        console.log("🔨 Enhanced snapshot extraction failed:", error);
      }
      
      // Try alternative range methods
      if (cellsFound === 0) {
        console.log("🔨 Trying alternative range access methods...");
        try {
          // Try different range sizes and methods
          const rangesToTry = ['A1:Z50', 'A1:AA100', 'A1:IV200'];
          
          for (const rangeStr of rangesToTry) {
            try {
              const range = fWorksheet.getRange(rangeStr);
              
              // Try different value extraction methods
              const methods = ['getValues', 'getDisplayValues', 'getFormattedValues'];
              
              for (const method of methods) {
                if (typeof range[method] === 'function') {
                  try {
                    const values = range[method]();
                    if (values && Array.isArray(values) && values.length > 0) {
                      console.log(`🔨 ${method} found data:`, values.length, "rows");
                      
                      values.forEach((row: any[], rowIndex: number) => {
                        if (Array.isArray(row)) {
                          row.forEach((cellValue: any, colIndex: number) => {
                            if (cellValue != null && cellValue !== '') {
                              if (!extractedData[rowIndex]) {
                                extractedData[rowIndex] = {};
                              }
                              extractedData[rowIndex][colIndex] = {
                                v: cellValue,
                                t: typeof cellValue === 'number' ? 1 : 0
                              };
                              cellsFound++;
                            }
                          });
                        }
                      });
                      
                      if (cellsFound > 0) break;
                    }
                  } catch (methodError) {
                    console.log(`🔨 ${method} failed:`, methodError);
                  }
                }
              }
              
              if (cellsFound > 0) break;
            } catch (rangeError) {
              console.log(`🔨 Range ${rangeStr} failed:`, rangeError);
            }
          }
        } catch (error) {
          console.log("🔨 Alternative range methods failed:", error);
        }
      }
    }

    console.log(`🔨 Brute force extraction completed: ${cellsFound} cells found`);
    return extractedData;
  }

  /**
   * Convert column index to letter (0=A, 1=B, etc.)
   */
  private columnToLetter(col: number): string {
    let result = '';
    while (col >= 0) {
      result = String.fromCharCode((col % 26) + 65) + result;
      col = Math.floor(col / 26) - 1;
    }
    return result;
  }

  /**
   * Build the base context from Univer API
   */
  private async buildBaseContext(): Promise<UniversalWorkbookContext> {
    const univerAPI = (window as any).univerAPI;
    if (!univerAPI) {
      throw new Error("Univer API not available");
    }

    const fWorkbook = univerAPI.getActiveWorkbook();
    if (!fWorkbook) {
      throw new Error("No active workbook");
    }

    const fWorksheet = fWorkbook.getActiveSheet();
    if (!fWorksheet) {
      throw new Error("No active worksheet");
    }

    // Force synchronization of any pending changes to the snapshot
    try {
      // First try to get the workbook to flush any pending operations
      fWorkbook.save(); // This should flush any pending changes
    } catch (e) {
      console.warn("Could not flush workbook before snapshot:", e);
    }
    
    const activeSheetSnapshot = fWorksheet.getSheet().getSnapshot();
    if (!activeSheetSnapshot) {
      throw new Error("Could not get sheet snapshot");
    }

    console.log(`📊 Raw snapshot cellData keys: ${Object.keys(activeSheetSnapshot.cellData || {}).length}`);
    
    // Enhanced data detection with multiple fallback methods
    let cellData = activeSheetSnapshot.cellData || {};
    
    // Fallback 1: Enhanced range-based data access with multiple methods
    if (Object.keys(cellData).length === 0) {
      console.log("🔄 Snapshot cellData empty, trying enhanced range-based fallback");
      
      const rangesToTry = ['A1:Z100', 'A1:AA200', 'A1:IV500'];
      const methodsToTry = ['getValues', 'getDisplayValues', 'getFormattedValues'];
      
      for (const rangeStr of rangesToTry) {
        try {
          const range = fWorksheet.getRange(rangeStr);
          
          for (const methodName of methodsToTry) {
            if (typeof range[methodName] === 'function') {
              try {
                console.log(`🔄 Trying ${methodName} on range ${rangeStr}`);
                const values = range[methodName]();
                
                // Convert range values to cellData format
                const convertedCellData: any = {};
                if (values && Array.isArray(values)) {
                  values.forEach((row, rowIndex) => {
                    if (Array.isArray(row)) {
                      row.forEach((cellValue, colIndex) => {
                        if (cellValue != null && cellValue !== '') {
                          if (!convertedCellData[rowIndex]) {
                            convertedCellData[rowIndex] = {};
                          }
                          convertedCellData[rowIndex][colIndex] = {
                            v: cellValue,
                            t: typeof cellValue === 'number' ? 1 : 0
                          };
                        }
                      });
                    }
                  });
                }
                
                if (Object.keys(convertedCellData).length > 0) {
                  console.log(`✅ ${methodName} on ${rangeStr} found data: ${Object.keys(convertedCellData).length} rows`);
                  cellData = convertedCellData;
                  activeSheetSnapshot.cellData = cellData;
                  break;
                }
              } catch (methodError) {
                console.log(`⚠️ ${methodName} on ${rangeStr} failed:`, methodError);
              }
            }
          }
          
          if (Object.keys(cellData).length > 0) break;
        } catch (rangeError) {
          console.warn(`⚠️ Range ${rangeStr} access failed:`, rangeError);
        }
      }
    }

    // Fallback 2: Direct cell scanning if still empty
    if (Object.keys(cellData).length === 0) {
      console.log("🔄 Range fallback empty, trying direct cell scanning");
      try {
        const scannedData: any = {};
        for (let row = 0; row < 50; row++) { // Scan first 50 rows
          for (let col = 0; col < 26; col++) { // Scan A-Z columns
            try {
              const colLetter = String.fromCharCode(65 + col);
              const cellRef = `${colLetter}${row + 1}`;
              const cell = fWorksheet.getRange(cellRef);
              const value = cell.getValue();
              
              if (value != null && value !== '') {
                if (!scannedData[row]) {
                  scannedData[row] = {};
                }
                scannedData[row][col] = {
                  v: value,
                  t: typeof value === 'number' ? 1 : 0
                };
              }
            } catch {} // Ignore individual cell errors
          }
        }
        
        if (Object.keys(scannedData).length > 0) {
          console.log(`✅ Direct cell scan found data: ${Object.keys(scannedData).length} rows`);
          cellData = scannedData;
          activeSheetSnapshot.cellData = cellData;
        }
      } catch (error) {
        console.warn("⚠️ Direct cell scanning failed:", error);
      }
    }

    console.log(`📊 Final cellData: ${Object.keys(cellData).length} rows detected`);

    // Focus on Univer API methods only - no DOM extraction

    console.log(`📊 Final cellData after all fallbacks: ${Object.keys(cellData).length} rows detected`);

    // Force additional data extraction if still empty - aggressive approach
    if (Object.keys(cellData).length === 0) {
      console.log("🚨 AGGRESSIVE: No data found through any method, trying brute force extraction");
      cellData = await this.bruteForceDataExtraction(fWorksheet);
      if (Object.keys(cellData).length > 0) {
        console.log(`🎯 Brute force extraction successful: ${Object.keys(cellData).length} rows found`);
        activeSheetSnapshot.cellData = cellData;
      }
    }

    console.log(`📊 FINAL RESULT: ${Object.keys(cellData).length} rows will be used for analysis`);

    // Run intelligent analysis
    const intelligence = analyzeSheetIntelligently(cellData);

    return {
      univerAPI,
      fWorkbook,
      fWorksheet,
      activeSheetName: activeSheetSnapshot.name || "Sheet1",
      activeSheetSnapshot,
      intelligence,
      tables: intelligence.tables,
      primaryTable: intelligence.tables[0] || null,
      calculableColumns: intelligence.tables.flatMap((table: any) =>
        table.columns
          .filter((col: any) => col.isCalculable)
          .map((col: any) => col.name)
      ),
      numericColumns: intelligence.tables.flatMap((table: any) =>
        table.columns
          .filter((col: any) => col.isNumeric)
          .map((col: any) => col.name)
      ),
      spatialMap: intelligence.spatialMap,
      lastUpdated: Date.now(),
      cacheValid: true,
    };
  }

  /**
   * Enrich base context with tool helpers
   */
  private enrichContext(
    baseContext: UniversalWorkbookContext
  ): UniversalToolContext {
    const enriched = baseContext as UniversalToolContext;

    // Tool-specific helpers
    enriched.findTable = (tableId?: string) => {
      if (!tableId) return enriched.primaryTable;
      return enriched.tables.find((t) => t.id === tableId) || null;
    };

    enriched.findColumn = (columnName: string, tableId?: string) => {
      const table = enriched.findTable(tableId);
      if (!table) return null;

      return (
        table.columns.find(
          (c: any) =>
            c.name.toLowerCase().includes(columnName.toLowerCase()) ||
            c.letter === columnName.toUpperCase()
        ) || null
      );
    };

    enriched.getTableRange = (tableId?: string) => {
      const table = enriched.findTable(tableId);
      return table?.range || "";
    };

    enriched.getColumnRange = (
      columnName: string,
      includeHeader = false,
      tableId?: string
    ) => {
      const table = enriched.findTable(tableId);
      const column = enriched.findColumn(columnName, tableId);

      if (!table || !column) return "";

      const startRow = includeHeader
        ? table.position.startRow
        : table.position.startRow + 1;
      const endRow = table.position.endRow;

      return `${column.letter}${startRow + 1}:${column.letter}${endRow + 1}`;
    };

    enriched.buildSumFormula = (columnName: string, tableId?: string) => {
      const dataRange = enriched.getColumnRange(columnName, false, tableId);
      return dataRange ? `=SUM(${dataRange})` : "";
    };

    enriched.findOptimalPlacement = (width: number, height: number) => {
      // Use spatial map to find best placement
      const spatialMap = enriched.spatialMap;

      // Default to right of primary table if no spatial analysis
      if (!spatialMap && enriched.primaryTable) {
        const table = enriched.primaryTable;
        return {
          row: table.position.startRow,
          col: table.position.endCol + 2,
          range: `${String.fromCharCode(65 + table.position.endCol + 2)}${
            table.position.startRow + 1
          }`,
        };
      }

      // Use spatial analysis for optimal placement
      const optimalZone = spatialMap?.optimalPlacementZones?.[0];
      if (optimalZone) {
        return {
          row: optimalZone.startRow,
          col: optimalZone.startCol,
          range: `${String.fromCharCode(65 + optimalZone.startCol)}${
            optimalZone.startRow + 1
          }`,
        };
      }

      // Fallback
      return { row: 0, col: 8, range: "I1" };
    };

    // State management
    enriched.invalidateCache = () => {
      this.cachedContext = null;
    };

    enriched.refresh = async () => {
      return await this.getContext(true);
    };

    return enriched;
  }

  /**
   * Check if cache is still valid
   */
  private isCacheValid(): boolean {
    if (!this.cachedContext) return false;

    const age = Date.now() - this.cachedContext.lastUpdated;
    return age < this.cacheTimeout && this.cachedContext.cacheValid;
  }

  /**
   * Invalidate cache (call when sheet changes)
   */
  public invalidateCache(): void {
    this.cachedContext = null;
  }

  /**
   * Update cache timeout
   */
  public setCacheTimeout(ms: number): void {
    this.cacheTimeout = ms;
  }
}

/**
 * Universal context factory - main entry point for all tools
 */
export async function getUniversalContext(
  forceRefresh = false
): Promise<UniversalToolContext> {
  const manager = UniversalContextManager.getInstance();
  return await manager.getContext(forceRefresh);
}

/**
 * Invalidate global context cache
 */
export function invalidateUniversalContext(): void {
  const manager = UniversalContextManager.getInstance();
  manager.invalidateCache();
}

/**
 * Decorator for tool functions to automatically provide context
 */
export function withUniversalContext<T extends any[], R>(
  toolFunction: (context: UniversalToolContext, ...args: T) => Promise<R>
) {
  return async (...args: T): Promise<R> => {
    const context = await getUniversalContext();
    return await toolFunction(context, ...args);
  };
}

/**
 * Debug helper to inspect current context
 */
export async function debugUniversalContext(): Promise<void> {
  try {
    const context = await getUniversalContext();

    console.log("🔍 UNIVERSAL CONTEXT DEBUG:", {
      activeSheet: context.activeSheetName,
      tablesCount: context.tables.length,
      tables: context.tables.map((t) => ({
        id: t.id,
        range: t.range,
        headers: t.headers,
        calculableColumns: t.columns
          .filter((c: any) => c.isCalculable)
          .map((c: any) => c.name),
      })),
      primaryTable: context.primaryTable
        ? {
            id: context.primaryTable.id,
            range: context.primaryTable.range,
            headers: context.primaryTable.headers,
          }
        : null,
      calculableColumns: context.calculableColumns,
      cacheAge: Date.now() - context.lastUpdated,
      spatialZones: context.spatialMap?.optimalPlacementZones?.length || 0,
    });
  } catch (error) {
    console.error("❌ Context debug failed:", error);
  }
}
