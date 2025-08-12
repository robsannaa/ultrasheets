/**
 * Modern Univer Bridge
 *
 * Bridges the new Universal Context system with the existing Univer component,
 * allowing gradual migration while maintaining backward compatibility.
 */

import {
  executeUniversalTool,
  registerUniversalTool,
  debugToolRegistry,
  getRegisteredTools,
  type ToolResult,
} from "./tool-executor";
import {
  getUniversalContext,
  invalidateUniversalContext,
  debugUniversalContext,
} from "./universal-context";
import { MODERN_TOOLS } from "./tools/modern-tools";
import { MIGRATED_TOOLS } from "./tools/migrated-tools";
import { ADDITIONAL_TOOLS } from "./tools/additional-tools";

/**
 * Initialize the modern tool system
 */
export function initializeModernTools(): void {
  // Register all modern tools
  MODERN_TOOLS.forEach((tool) => {
    registerUniversalTool(tool);
  });

  // Register all migrated tools
  MIGRATED_TOOLS.forEach((tool) => {
    registerUniversalTool(tool);
  });

  // Register all additional tools
  ADDITIONAL_TOOLS.forEach((tool) => {
    registerUniversalTool(tool);
  });

  // Debug info in development
  if (process.env.NODE_ENV === "development") {
    debugToolRegistry();
  }
}

/**
 * Execute modern tools only - no legacy fallback
 */
export async function executeModernUniverTool(
  toolName: string,
  params: any = {}
): Promise<ToolResult> {
  const result = await executeUniversalTool(toolName, params);

  if (result.success) {
    console.log(`✅ Modern tool '${toolName}' executed successfully:`, result);
  } else {
    console.warn(`⚠️ Modern tool '${toolName}' failed:`, result.error);
  }

  return result;
}

/**
 * Context synchronization utilities
 */
export class ContextSync {
  /**
   * Invalidate context when sheet changes
   */
  static onSheetChange(): void {
    invalidateUniversalContext();
    console.log("🔄 Context invalidated due to sheet change");
  }

  /**
   * Refresh context manually
   */
  static async refreshContext(): Promise<void> {
    await getUniversalContext(true);
    console.log("🔄 Context manually refreshed");
  }

  /**
   * Debug current context state
   */
  static async debugContext(): Promise<void> {
    await debugUniversalContext();
  }
}

/**
 * Performance monitoring
 */
export class PerformanceMonitor {
  private static metrics: Array<{
    toolName: string;
    executionTime: number;
    success: boolean;
    timestamp: number;
    isModern: boolean;
  }> = [];

  static record(
    toolName: string,
    executionTime: number,
    success: boolean,
    isModern: boolean
  ): void {
    this.metrics.push({
      toolName,
      executionTime,
      success,
      timestamp: Date.now(),
      isModern,
    });

    // Keep only last 100 metrics
    if (this.metrics.length > 100) {
      this.metrics = this.metrics.slice(-100);
    }
  }

  static getReport(): {
    totalExecutions: number;
    averageExecutionTime: number;
    successRate: number;
    modernToolUsage: number;
    slowestTools: Array<{ tool: string; avgTime: number }>;
  } {
    if (this.metrics.length === 0) {
      return {
        totalExecutions: 0,
        averageExecutionTime: 0,
        successRate: 0,
        modernToolUsage: 0,
        slowestTools: [],
      };
    }

    const totalExecutions = this.metrics.length;
    const averageExecutionTime =
      this.metrics.reduce((sum, m) => sum + m.executionTime, 0) /
      totalExecutions;
    const successfulExecutions = this.metrics.filter((m) => m.success).length;
    const successRate = (successfulExecutions / totalExecutions) * 100;
    const modernExecutions = this.metrics.filter((m) => m.isModern).length;
    const modernToolUsage = (modernExecutions / totalExecutions) * 100;

    // Calculate average time per tool
    const toolTimes = this.metrics.reduce((acc, metric) => {
      if (!acc[metric.toolName]) {
        acc[metric.toolName] = { total: 0, count: 0 };
      }
      acc[metric.toolName].total += metric.executionTime;
      acc[metric.toolName].count += 1;
      return acc;
    }, {} as Record<string, { total: number; count: number }>);

    const slowestTools = Object.entries(toolTimes)
      .map(([tool, data]) => ({
        tool,
        avgTime: data.total / data.count,
      }))
      .sort((a, b) => b.avgTime - a.avgTime)
      .slice(0, 5);

    return {
      totalExecutions,
      averageExecutionTime,
      successRate,
      modernToolUsage,
      slowestTools,
    };
  }

  static logReport(): void {
    const report = this.getReport();
    console.log("📊 PERFORMANCE REPORT:", report);
  }
}

/**
 * Migration utilities
 */
export class MigrationHelper {
  /**
   * Check which tools need migration
   */
  static async analyzeToolUsage(): Promise<{
    modernTools: string[];
    legacyTools: string[];
    recommendations: string[];
  }> {
    const modernTools = getRegisteredTools().map((tool) => tool.name);
    const legacyFunctions = (window as any).__ultraToolFns || {};
    const legacyTools = Object.keys(legacyFunctions)
      .map((key) => key.replace("execute_", ""))
      .filter((tool) => !modernTools.includes(tool));

    const recommendations = legacyTools.map(
      (tool) => `Consider migrating '${tool}' to modern tool framework`
    );

    return {
      modernTools,
      legacyTools,
      recommendations,
    };
  }

  /**
   * Generate migration template for a legacy tool
   */
  static generateMigrationTemplate(toolName: string): string {
    return `
// Migration template for ${toolName}
export const ${
      toolName.charAt(0).toUpperCase() + toolName.slice(1)
    }Tool = createSimpleTool(
  {
    name: '${toolName}',
    description: 'TODO: Add description',
    category: 'data', // TODO: Choose correct category
    requiredContext: ['tables'], // TODO: Specify requirements
    invalidatesCache: false, // TODO: Set if tool modifies data
  },
  async (context: UniversalToolContext, params: any) => {
    // TODO: Implement using context helpers:
    // - context.findTable(tableId)
    // - context.findColumn(columnName, tableId)
    // - context.getTableRange(tableId)
    // - context.buildSumFormula(columnName, tableId)
    // - context.findOptimalPlacement(width, height)
    
    throw new Error('Migration not implemented yet');
  }
);
    `.trim();
  }
}

/**
 * Global setup function to be called from Univer component
 */
export function setupModernUniverBridge(): void {
  // Initialize modern tools
  initializeModernTools();

  // Replace the global executeUniverTool function
  (window as any).executeUniverTool = executeModernUniverTool;

  // Add context utilities to global scope for debugging
  (window as any).__modernContext = {
    getContext: getUniversalContext,
    debugContext: debugUniversalContext,
    invalidateContext: invalidateUniversalContext,
    syncContext: ContextSync,
    performance: PerformanceMonitor,
    migration: MigrationHelper,
  };
}
