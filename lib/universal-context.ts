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

    const activeSheetSnapshot = fWorksheet.getSheet().getSnapshot();
    if (!activeSheetSnapshot) {
      throw new Error("Could not get sheet snapshot");
    }

    // Run intelligent analysis
    const intelligence = analyzeSheetIntelligently(
      activeSheetSnapshot.cellData || {}
    );

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
