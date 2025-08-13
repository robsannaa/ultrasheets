/**
 * Universal Tool Executor Framework
 *
 * Standardizes tool execution with automatic context injection,
 * error handling, logging, and result formatting.
 */

import {
  getWorkbookData,
  getCellValue,
  getRangeValues,
  getCurrentSelection,
} from "./univer-data-source";

// Production-grade context interface using only Univer API
export interface UniversalToolContext {
  // Univer API references
  univerAPI: any;
  fWorkbook: any;
  fWorksheet: any;
  
  // Helper methods
  findTable: (tableId?: string) => any | null;
  findColumn: (columnName: string, tableId?: string) => any | null;
  findOptimalPlacement: (width: number, height: number) => { range: string };
}

export interface ToolResult {
  success: boolean;
  data?: any;
  message: string;
  error?: string;
  metadata?: {
    executionTime: number;
    contextCacheHit: boolean;
    tablesAnalyzed: number;
    operationType: string;
    commandExecuted?: boolean;
    retryAttempts?: number;
  };
}

export interface ToolDefinition {
  name: string;
  description: string;
  category: "data" | "format" | "analysis" | "structure" | "navigation";
  requiredContext: ("tables" | "columns")[];
  invalidatesCache: boolean;
}

/**
 * Base class for all Univer tools
 */
export abstract class UniversalTool {
  protected definition: ToolDefinition;

  constructor(definition: ToolDefinition) {
    this.definition = definition;
  }

  /**
   * Get tool definition
   */
  public getDefinition(): ToolDefinition {
    return this.definition;
  }

  /**
   * Execute the tool with automatic context management and enhanced error handling
   */
  public async execute(params: any): Promise<ToolResult> {
    const startTime = Date.now();
    let context: UniversalToolContext;
    let retryAttempts = 0;
    const maxRetries = 2;

    while (retryAttempts <= maxRetries) {
      try {
        // Get context using direct Univer API
        context = await createUniversalContext();

        // Validate required context
        this.validateContext(context);

        // Execute the tool logic with command system monitoring
        const result = await this.executeWithCommandMonitoring(context, params);

        // No cache to invalidate with direct Univer API access
        if (this.definition.invalidatesCache) {
          console.log('Tool marked for cache invalidation, but using direct API (no cache)');
        }

        const executionTime = Date.now() - startTime;

        return {
          success: true,
          data: result,
          message: `${this.definition.name} completed successfully`,
          metadata: {
            executionTime,
            contextCacheHit: false, // Direct API access, no cache
            tablesAnalyzed: 0, // Simplified for direct API
            operationType: this.definition.category,
            commandExecuted: true,
            retryAttempts,
          },
        };

      } catch (error) {
        retryAttempts++;
        
        // Check if this is a retryable error
        if (retryAttempts <= maxRetries && this.isRetryableError(error)) {
          console.warn(`⚠️ Tool ${this.definition.name} attempt ${retryAttempts} failed, retrying...`, error);
          
          // No cache to clear with direct Univer API access
          console.log('Retry attempt - using fresh Univer API data');
          
          // Wait a bit before retry
          await new Promise(resolve => setTimeout(resolve, 100 * retryAttempts));
          continue;
        }

        // Final failure
        const executionTime = Date.now() - startTime;
        console.error(`❌ Tool ${this.definition.name} failed after ${retryAttempts} attempts:`, error);

        return {
          success: false,
          message: `${this.definition.name} failed`,
          error: error instanceof Error ? error.message : "Unknown error",
          metadata: {
            executionTime,
            contextCacheHit: false,
            tablesAnalyzed: 0, // Simplified for direct API
            operationType: this.definition.category,
            commandExecuted: false,
            retryAttempts: retryAttempts - 1,
          },
        };
      }
    }

    // Should never reach here, but TypeScript needs this
    throw new Error('Unexpected execution path');
  }

  /**
   * Execute tool with command system monitoring
   */
  private async executeWithCommandMonitoring<T>(
    context: UniversalToolContext,
    params: any
  ): Promise<T> {
    const univerAPI = context.univerAPI;
    const toolName = this.definition.name;
    
    return new Promise(async (resolve, reject) => {
      let operationResult: T;
      let commandMonitoringActive = false;
      let commandExecutionTimer: NodeJS.Timeout | null = null;
      
      try {
        // Set up command monitoring if available
        if (univerAPI?.onBeforeCommandExecute && univerAPI?.onCommandExecuted) {
          commandMonitoringActive = true;
          
          const onBeforeExecute = (command: any) => {
            if (this.isRelevantCommand(command, toolName)) {
              console.log(`🔧 Tool ${toolName} executing command:`, command.id || command.type);
            }
          };
          
          const onCommandExecuted = (command: any) => {
            if (this.isRelevantCommand(command, toolName)) {
              console.log(`✅ Tool ${toolName} command completed:`, command.id || command.type);
            }
          };
          
          const onCommandFailed = (error: any, command: any) => {
            if (this.isRelevantCommand(command, toolName)) {
              console.error(`❌ Tool ${toolName} command failed:`, command?.id || 'unknown', error);
              if (commandExecutionTimer) clearTimeout(commandExecutionTimer);
              reject(new Error(`Command failed: ${error.message || error}`));
              return;
            }
          };

          univerAPI.addEventListener?.('onBeforeCommandExecute', onBeforeExecute);
          univerAPI.addEventListener?.('onCommandExecuted', onCommandExecuted);
          univerAPI.addEventListener?.('onCommandFailed', onCommandFailed);
          
          // Cleanup function
          const cleanup = () => {
            univerAPI.removeEventListener?.('onBeforeCommandExecute', onBeforeExecute);
            univerAPI.removeEventListener?.('onCommandExecuted', onCommandExecuted);
            univerAPI.removeEventListener?.('onCommandFailed', onCommandFailed);
            if (commandExecutionTimer) clearTimeout(commandExecutionTimer);
          };

          // Set timeout for command execution
          commandExecutionTimer = setTimeout(() => {
            cleanup();
            // Don't reject on timeout, just log warning
            console.warn(`⚠️ Tool ${toolName} command monitoring timed out, proceeding anyway`);
          }, 10000);

          try {
            // Execute the core tool logic
            operationResult = await this.executeCore(context, params);
            cleanup();
            resolve(operationResult);
          } catch (error) {
            cleanup();
            reject(error);
          }

        } else {
          // No command monitoring available, execute directly
          operationResult = await this.executeCore(context, params);
          resolve(operationResult);
        }
        
      } catch (error) {
        if (commandExecutionTimer) clearTimeout(commandExecutionTimer);
        reject(error);
      }
    });
  }

  /**
   * Check if a command is relevant to this tool
   */
  private isRelevantCommand(command: any, toolName: string): boolean {
    if (!command) return false;
    
    const commandId = (command.id || command.type || '').toLowerCase();
    const lowerToolName = toolName.toLowerCase();
    
    // General relevance checks
    const relevantPatterns = [
      lowerToolName,
      lowerToolName.replace(/_/g, '-'),
      lowerToolName.replace(/_/g, ''),
    ];
    
    return relevantPatterns.some(pattern => 
      commandId.includes(pattern) || pattern.includes(commandId.replace(/[^a-z]/g, ''))
    );
  }

  /**
   * Check if an error is retryable
   */
  private isRetryableError(error: any): boolean {
    const errorMessage = (error?.message || error || '').toLowerCase();
    
    const retryablePatterns = [
      'network',
      'timeout', 
      'connection',
      'temporary',
      'not ready',
      'loading',
      'busy',
      'locked'
    ];
    
    return retryablePatterns.some(pattern => errorMessage.includes(pattern));
  }

  /**
   * Validate that context has required elements
   */
  private validateContext(context: UniversalToolContext): void {
    // Simplified validation for direct Univer API access
    if (!context.univerAPI) {
      throw new Error("Univer API not available");
    }
    
    if (!context.fWorkbook) {
      throw new Error("No active workbook");
    }
    
    if (!context.fWorksheet) {
      throw new Error("No active worksheet");
    }
    
    // For tools requiring specific context, we check using Univer API
    for (const requirement of this.definition.requiredContext) {
      switch (requirement) {
        case "tables":
        case "columns":
          const workbookData = getWorkbookData();
          if (!workbookData || workbookData.sheets.length === 0) {
            throw new Error("No data found in the workbook");
          }
          const activeSheet = workbookData.sheets.find(s => s.sheetName === workbookData.activeSheetName);
          if (!activeSheet || !activeSheet.usedRange) {
            throw new Error("No data range found in the active sheet");
          }
          break;
      }
    }
  }

  /**
   * Core tool logic - implemented by each tool
   */
  protected abstract executeCore(
    context: UniversalToolContext,
    params: any
  ): Promise<any>;
}

/**
 * Tool Registry - manages all available tools
 */
class ToolRegistry {
  private static instance: ToolRegistry;
  private tools = new Map<string, UniversalTool>();

  private constructor() {}

  public static getInstance(): ToolRegistry {
    if (!ToolRegistry.instance) {
      ToolRegistry.instance = new ToolRegistry();
    }
    return ToolRegistry.instance;
  }

  public register(tool: UniversalTool): void {
    this.tools.set(tool.getDefinition().name, tool);
  }

  public get(name: string): UniversalTool | undefined {
    return this.tools.get(name);
  }

  public list(): ToolDefinition[] {
    return Array.from(this.tools.values()).map((tool) => tool.getDefinition());
  }

  public async execute(toolName: string, params: any): Promise<ToolResult> {
    const tool = this.get(toolName);
    if (!tool) {
      return {
        success: false,
        message: `Tool '${toolName}' not found`,
        error: `Available tools: ${Array.from(this.tools.keys()).join(", ")}`,
      };
    }

    return await tool.execute(params);
  }
}

/**
 * Global tool executor function
 */
export async function executeUniversalTool(
  toolName: string,
  params: any = {}
): Promise<ToolResult> {
  const registry = ToolRegistry.getInstance();
  return await registry.execute(toolName, params);
}

/**
 * Register a tool with the global registry
 */
export function registerUniversalTool(tool: UniversalTool): void {
  const registry = ToolRegistry.getInstance();
  registry.register(tool);
}

/**
 * Get list of all registered tools
 */
export function getRegisteredTools(): ToolDefinition[] {
  const registry = ToolRegistry.getInstance();
  return registry.list();
}

/**
 * Helper function to create simple tools without classes
 */
export function createSimpleTool(
  definition: ToolDefinition,
  executeFunction: (context: UniversalToolContext, params: any) => Promise<any>
): UniversalTool {
  return new (class extends UniversalTool {
    protected async executeCore(
      context: UniversalToolContext,
      params: any
    ): Promise<any> {
      return await executeFunction(context, params);
    }
  })(definition);
}

/**
 * Create Universal Context using direct Univer API access
 * No caching, no heuristics - production grade
 */
async function createUniversalContext(): Promise<UniversalToolContext> {
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

  return {
    univerAPI,
    fWorkbook,
    fWorksheet,
    
    // Simple helper methods using direct Univer API
    findTable: (tableId?: string) => {
      // For now, return a basic table structure from workbook data
      const workbookData = getWorkbookData();
      if (!workbookData || workbookData.sheets.length === 0) return null;
      
      const sheet = workbookData.sheets[0]; // Use first sheet
      if (!sheet.usedRange) return null;
      
      return {
        id: tableId || 'main',
        range: sheet.usedRange?.address || 'A1:A1',
        position: sheet.usedRange ? {
          startRow: sheet.usedRange.startRow,
          endRow: sheet.usedRange.endRow,
          startCol: sheet.usedRange.startCol,
          endCol: sheet.usedRange.endCol
        } : { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        headers: sheet.usedRange ? sheet.cells
          .filter(cell => cell.row === sheet.usedRange!.startRow)
          .sort((a, b) => a.col - b.col)
          .map(cell => cell.value?.toString() || `Col${cell.col}`) : [],
        columns: sheet.usedRange ? sheet.cells
          .filter(cell => cell.row === sheet.usedRange!.startRow)
          .sort((a, b) => a.col - b.col)
          .map((cell, index) => ({
            name: cell.value?.toString() || `Col${cell.col}`,
            letter: String.fromCharCode(65 + cell.col),
            index: cell.col,
            dataType: 'text' // Simplified
          })) : []
      };
    },
    
    findColumn: (columnName: string, tableId?: string) => {
      const workbookData = getWorkbookData();
      if (!workbookData || workbookData.sheets.length === 0) return null;
      
      const sheet = workbookData.sheets[0];
      if (!sheet.usedRange) return null;
      
      const headers = sheet.usedRange ? sheet.cells
        .filter(cell => cell.row === sheet.usedRange!.startRow)
        .sort((a, b) => a.col - b.col) : [];
        
      const headerCell = headers.find(cell => 
        cell.value?.toString().toLowerCase().includes(columnName.toLowerCase())
      );
      
      if (!headerCell) return null;
      
      return {
        name: headerCell.value?.toString() || `Col${headerCell.col}`,
        letter: String.fromCharCode(65 + headerCell.col),
        index: headerCell.col,
        dataType: 'text'
      };
    },
    
    findOptimalPlacement: (width: number, height: number) => {
      // Simple placement logic - just use column H by default
      return { range: "H1" };
    }
  };
}

/**
 * Debug helper to show tool registry status
 */
export function debugToolRegistry(): void {
  const tools = getRegisteredTools();
}
