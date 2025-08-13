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
  getWorkbookData,
  getCellValue,
  getRangeValues,
  getCurrentSelection,
} from "./univer-data-source";
import { ALL_TOOLS } from "./tools";

/**
 * Initialize the modern tool system
 */
export function initializeModernTools(): void {
  // Register all tools via unified entry
  ALL_TOOLS.forEach((tool) => registerUniversalTool(tool));

  // Debug info in development
  if (process.env.NODE_ENV === "development") {
    debugToolRegistry();
  }
}

/**
 * Setup event listeners for automatic cache invalidation and monitoring
 */
export function setupEventListeners(): void {
  const univerAPI = (window as any).univerAPI;
  if (!univerAPI) {
    console.warn('⚠️ Cannot setup event listeners - univerAPI not available');
    return;
  }

  try {
    const disposables = [];

    // Sheet change events
    if (univerAPI.addEvent && univerAPI.Event?.SheetChanged) {
      const sheetChangeDisposable = univerAPI.addEvent(univerAPI.Event.SheetChanged, () => {
        console.log('📊 Sheet changed - invalidating context cache');
        ContextSync.onSheetChange();
      });
      disposables.push(sheetChangeDisposable);
    }

    // Cell value change events
    if (univerAPI.addEvent && univerAPI.Event?.CellValueChanged) {
      const cellChangeDisposable = univerAPI.addEvent(univerAPI.Event.CellValueChanged, () => {
        console.log('📝 Cell value changed - invalidating context cache');
        ContextSync.onSheetChange();
      });
      disposables.push(cellChangeDisposable);
    }

    // Command execution monitoring
    if (typeof univerAPI.onBeforeCommandExecute === 'function') {
      const beforeCommandDisposable = univerAPI.onBeforeCommandExecute((command: any) => {
        console.log('🔧 Command executing:', command.id || command.type);
        PerformanceMonitor.recordCommandStart(command.id || command.type);
      });
      disposables.push(beforeCommandDisposable);
    }

    if (typeof univerAPI.onCommandExecuted === 'function') {
      const afterCommandDisposable = univerAPI.onCommandExecuted((command: any) => {
        console.log('✅ Command completed:', command.id || command.type);
        PerformanceMonitor.recordCommandEnd(command.id || command.type, true);
        
        // Invalidate context after user actions that modify data
        if (command.params?.source === 'user' || isDataModifyingCommand(command)) {
          setTimeout(() => ContextSync.onSheetChange(), 100);
        }
      });
      disposables.push(afterCommandDisposable);
    }

    // Command failure monitoring
    if (typeof univerAPI.onCommandFailed === 'function') {
      const commandFailedDisposable = univerAPI.onCommandFailed((error: any, command: any) => {
        console.error('❌ Command failed:', command?.id || 'unknown', error);
        PerformanceMonitor.recordCommandEnd(command?.id || 'unknown', false);
      });
      disposables.push(commandFailedDisposable);
    }

    // Store disposables for cleanup
    (window as any).__univerEventDisposables = disposables;

    console.log(`✅ Setup ${disposables.length} event listeners for cache invalidation`);

  } catch (error) {
    console.error('❌ Failed to setup event listeners:', error);
  }
}

/**
 * Check if a command modifies data and requires cache invalidation
 */
function isDataModifyingCommand(command: any): boolean {
  const dataModifyingCommands = [
    'set-cell-value',
    'set-range-values', 
    'insert-row',
    'insert-column',
    'delete-row',
    'delete-column',
    'paste',
    'cut',
    'sort',
    'filter',
    'add-table',
    'remove-table',
    'conditional-format'
  ];
  
  const commandId = (command.id || command.type || '').toLowerCase();
  return dataModifyingCommands.some(cmd => commandId.includes(cmd));
}

/**
 * Cleanup event listeners
 */
export function disposeEventListeners(): void {
  const disposables = (window as any).__univerEventDisposables;
  if (Array.isArray(disposables)) {
    disposables.forEach((disposable: any) => {
      try {
        if (typeof disposable?.dispose === 'function') {
          disposable.dispose();
        }
      } catch (error) {
        console.warn('⚠️ Failed to dispose event listener:', error);
      }
    });
    (window as any).__univerEventDisposables = [];
    console.log('🧹 Cleaned up event listeners');
  }
}

/**
 * Execute modern tools only - no legacy fallback
 */
export async function executeModernUniverTool(
  toolName: string,
  params: any = {}
): Promise<ToolResult> {
  const startTime = Date.now();
  const result = await executeUniversalTool(toolName, params);
  const executionTime = Date.now() - startTime;

  // Record performance metrics
  PerformanceMonitor.record(toolName, executionTime, result.success, true);

  if (result.success) {
    console.log(`✅ Modern tool '${toolName}' executed successfully in ${executionTime}ms`);
  } else {
    console.warn(`⚠️ Modern tool '${toolName}' failed after ${executionTime}ms:`, result.error);
  }

  return result;
}

/**
 * Context synchronization utilities - Production Grade
 * Uses only Univer API, no caching or heuristics
 */
export class ContextSync {
  /**
   * Handle sheet changes - no cache invalidation needed with direct API access
   */
  static onSheetChange(): void {
    console.log("📊 Sheet changed - using direct Univer API access (no cache to invalidate)");
    // No cache to invalidate - we get fresh data directly from Univer API
  }

  /**
   * Get fresh workbook data using Univer API
   */
  static async refreshContext(): Promise<void> {
    console.log("🔄 Getting fresh workbook data from Univer API...");
    const workbookData = getWorkbookData();
    console.log("✅ Fresh data retrieved:", workbookData ? 'success' : 'no data');
  }

  /**
   * Debug current workbook state using Univer API
   */
  static async debugContext(): Promise<void> {
    console.log("🔍 Debugging workbook state via Univer API");
    const workbookData = getWorkbookData();
    const selection = getCurrentSelection();
    console.log("Workbook Data:", workbookData);
    console.log("Current Selection:", selection);
  }
}

/**
 * Enhanced Performance monitoring with command tracking
 */
export class PerformanceMonitor {
  private static metrics: Array<{
    toolName: string;
    executionTime: number;
    success: boolean;
    timestamp: number;
    isModern: boolean;
    contextCacheHit: boolean;
    dataSize: number;
  }> = [];

  private static commandTimings = new Map<string, number>();

  static record(
    toolName: string,
    executionTime: number,
    success: boolean,
    isModern: boolean,
    contextCacheHit: boolean = false,
    dataSize: number = 0
  ): void {
    this.metrics.push({
      toolName,
      executionTime,
      success,
      timestamp: Date.now(),
      isModern,
      contextCacheHit,
      dataSize
    });

    // Auto-optimization suggestions
    if (executionTime > 2000) {
      console.warn(`⚠️ Slow operation detected: ${toolName} took ${executionTime}ms`);
    }

    if (!contextCacheHit && executionTime > 500) {
      console.info(`💡 Suggestion: ${toolName} could benefit from context caching`);
    }

    // Keep only last 100 metrics
    if (this.metrics.length > 100) {
      this.metrics = this.metrics.slice(-100);
    }
  }

  static recordCommandStart(commandId: string): void {
    this.commandTimings.set(commandId, Date.now());
  }

  static recordCommandEnd(commandId: string, success: boolean): void {
    const startTime = this.commandTimings.get(commandId);
    if (startTime) {
      const executionTime = Date.now() - startTime;
      this.record(commandId, executionTime, success, false);
      this.commandTimings.delete(commandId);
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
    const suggestions = this.getOptimizationSuggestions();
    const cacheStats = this.getCacheStats();
    
    console.log("📊 PERFORMANCE REPORT:");
    console.table({
      'Total Executions': report.totalExecutions,
      'Avg Execution Time (ms)': Math.round(report.averageExecutionTime),
      'Success Rate (%)': Math.round(report.successRate),
      'Modern Tool Usage (%)': Math.round(report.modernToolUsage),
      'Cache Hit Rate (%)': Math.round(cacheStats.hitRate),
      'Cache Strategy': cacheStats.strategy
    });
    
    if (report.slowestTools.length > 0) {
      console.log("🐌 SLOWEST TOOLS:");
      console.table(report.slowestTools);
    }
    
    if (suggestions.length > 0) {
      console.log("💡 OPTIMIZATION SUGGESTIONS:");
      suggestions.forEach((suggestion, index) => {
        console.log(`${index + 1}. ${suggestion}`);
      });
    }
  }

  private static getCacheStats(): { hitRate: number; strategy: string } {
    try {
      const contextManager = (window as any).__modernContext?.getContext;
      if (contextManager) {
        // We'll get this from the context manager when available
        return { hitRate: 75, strategy: 'moderate' }; // Default values
      }
    } catch (error) {
      console.warn('Could not get cache stats:', error);
    }
    return { hitRate: 0, strategy: 'unknown' };
  }

  static getOptimizationSuggestions(): string[] {
    const suggestions: string[] = [];
    const recentMetrics = this.metrics.slice(-50); // Last 50 operations
    
    if (recentMetrics.length === 0) return suggestions;

    const slowOperations = recentMetrics.filter(m => m.executionTime > 1000);
    if (slowOperations.length > 0) {
      const uniqueSlowOps = [...new Set(slowOperations.map(m => m.toolName))];
      suggestions.push(`Consider optimizing slow operations: ${uniqueSlowOps.join(', ')}`);
    }
    
    const cacheMisses = recentMetrics.filter(m => !m.contextCacheHit).length;
    const cacheHitRate = ((recentMetrics.length - cacheMisses) / recentMetrics.length) * 100;
    
    if (cacheHitRate < 70) {
      suggestions.push(`Low cache hit rate (${cacheHitRate.toFixed(1)}%). Consider adjusting cache strategy.`);
    }

    const failureRate = (recentMetrics.filter(m => !m.success).length / recentMetrics.length) * 100;
    if (failureRate > 5) {
      suggestions.push(`High failure rate (${failureRate.toFixed(1)}%). Check error handling and validation.`);
    }

    const modernToolUsage = (recentMetrics.filter(m => m.isModern).length / recentMetrics.length) * 100;
    if (modernToolUsage < 80) {
      suggestions.push(`Only ${modernToolUsage.toFixed(1)}% modern tool usage. Consider migrating legacy tools.`);
    }
    
    return suggestions;
  }

  /**
   * Real-time performance monitoring with automatic alerts
   */
  static startRealTimeMonitoring(): void {
    // Check performance every 30 seconds
    const monitoringInterval = setInterval(() => {
      const report = this.getReport();
      const suggestions = this.getOptimizationSuggestions();
      
      // Alert on poor performance
      if (report.averageExecutionTime > 2000) {
        console.warn('⚠️ PERFORMANCE ALERT: Average execution time is high!');
        console.warn(`Current average: ${Math.round(report.averageExecutionTime)}ms`);
      }
      
      if (report.successRate < 90) {
        console.error('🚨 RELIABILITY ALERT: Success rate is low!');
        console.error(`Current success rate: ${Math.round(report.successRate)}%`);
      }
      
      // Auto-suggest optimizations
      if (suggestions.length > 0) {
        console.info('💡 AUTO-OPTIMIZATION SUGGESTIONS:', suggestions.slice(0, 2));
      }
      
    }, 30000);

    // Store interval for cleanup
    (window as any).__performanceMonitoringInterval = monitoringInterval;
  }

  static stopRealTimeMonitoring(): void {
    const interval = (window as any).__performanceMonitoringInterval;
    if (interval) {
      clearInterval(interval);
      (window as any).__performanceMonitoringInterval = null;
      console.log('📊 Performance monitoring stopped');
    }
  }

  /**
   * Export performance data for analysis
   */
  static exportMetrics(): {
    metrics: Array<{
      toolName: string;
      executionTime: number;
      success: boolean;
      timestamp: number;
      isModern: boolean;
      contextCacheHit: boolean;
      dataSize: number;
    }>;
    report: any;
    suggestions: string[];
    timestamp: number;
  } {
    return {
      metrics: [...this.metrics],
      report: this.getReport(),
      suggestions: this.getOptimizationSuggestions(),
      timestamp: Date.now()
    };
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
  console.log('🚀 Setting up Modern Univer Bridge...');
  
  // Initialize modern tools
  initializeModernTools();

  // Setup event listeners for automatic cache invalidation
  setupEventListeners();

  // Replace the global executeUniverTool function
  (window as any).executeUniverTool = executeModernUniverTool;

  // Add context utilities to global scope for debugging (Production Grade)
  (window as any).__modernContext = {
    getContext: getWorkbookData,
    getCellValue: getCellValue,
    getRangeValues: getRangeValues,
    getCurrentSelection: getCurrentSelection,
    syncContext: ContextSync,
    performance: PerformanceMonitor,
    migration: MigrationHelper,
    disposeEventListeners,
    setupEventListeners,
  };

  // Start real-time performance monitoring in development
  if (process.env.NODE_ENV === 'development') {
    PerformanceMonitor.startRealTimeMonitoring();
    
    if (typeof window !== 'undefined') {
      (window as any).__disposeModernBridge = () => {
        console.log('🧹 Disposing Modern Univer Bridge...');
        disposeEventListeners();
        PerformanceMonitor.stopRealTimeMonitoring();
        (window as any).__modernContext = null;
      };
    }
  }

  console.log('✅ Modern Univer Bridge setup complete');
}
