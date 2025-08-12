/**
 * Universal Tool Executor Framework
 *
 * Standardizes tool execution with automatic context injection,
 * error handling, logging, and result formatting.
 */

import {
  getUniversalContext,
  invalidateUniversalContext,
  type UniversalToolContext,
} from "./universal-context";

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
   * Execute the tool with automatic context management
   */
  public async execute(params: any): Promise<ToolResult> {
    const startTime = Date.now();
    let context: UniversalToolContext;

    try {
      // Get universal context
      context = await getUniversalContext();

      // Validate required context
      this.validateContext(context);

      // Execute the tool logic
      const result = await this.executeCore(context, params);

      // Invalidate cache if needed
      if (this.definition.invalidatesCache) {
        invalidateUniversalContext();
      }

      const executionTime = Date.now() - startTime;

      return {
        success: true,
        data: result,
        message: `${this.definition.name} completed successfully`,
        metadata: {
          executionTime,
          contextCacheHit: true, // TODO: track cache hits
          tablesAnalyzed: context.tables.length,
          operationType: this.definition.category,
        },
      };
    } catch (error) {
      const executionTime = Date.now() - startTime;
      console.error(`❌ Tool ${this.definition.name} failed:`, error);

      return {
        success: false,
        message: `${this.definition.name} failed`,
        error: error instanceof Error ? error.message : "Unknown error",
        metadata: {
          executionTime,
          contextCacheHit: false,
          tablesAnalyzed: 0,
          operationType: this.definition.category,
        },
      };
    }
  }

  /**
   * Validate that context has required elements
   */
  private validateContext(context: UniversalToolContext): void {
    for (const requirement of this.definition.requiredContext) {
      switch (requirement) {
        case "tables":
          if (context.tables.length === 0) {
            throw new Error(
              "No tables found in the sheet. This tool requires at least one data table."
            );
          }
          break;
        case "columns":
          if (
            !context.primaryTable ||
            context.primaryTable.columns.length === 0
          ) {
            throw new Error("No columns found in the primary table.");
          }
          break;
        // spatial requirement removed
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
 * Debug helper to show tool registry status
 */
export function debugToolRegistry(): void {
  const tools = getRegisteredTools();
}
