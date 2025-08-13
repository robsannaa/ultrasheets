/**
 * Pivot Table Tool using Univer's official Pivot Table API
 *
 * This tool follows the official Univer API documentation for pivot tables:
 * - Uses proper builder pattern for pivot table creation
 * - Supports various aggregation functions
 * - Handles row/column grouping
 * - Provides intelligent data source detection
 */

import { createSimpleTool } from "../tool-executor";
import type { UniversalToolContext } from "../tool-executor";

export const CreatePivotTableTool = createSimpleTool(
  {
    name: "create_pivot_table",
    description:
      "Create pivot tables using Univer's pivot table API with intelligent data detection",
    category: "analysis",
    requiredContext: ["tables"],
    invalidatesCache: true,
  },
  async (
    context: UniversalToolContext,
    params: {
      sourceRange?: string;
      tableId?: string;
      destination?: string;
      rows?: string[];
      columns?: string[];
      values?: Array<{
        field: string;
        aggregation:
          | "sum"
          | "average"
          | "count"
          | "max"
          | "min"
          | "product"
          | "stddev";
      }>;
      // Legacy single-field support for backward compatibility
      groupBy?: string;
      valueColumn?: string;
      aggFunc?: "sum" | "average" | "count" | "max" | "min";
    }
  ) => {
    const {
      sourceRange,
      tableId,
      destination,
      rows = [],
      columns = [],
      values = [],
      // Legacy parameters
      groupBy,
      valueColumn,
      aggFunc = "sum",
    } = params;

    console.log(`📊 CreatePivotTableTool: Creating pivot table`);

    try {
      // Determine source range
      let finalSourceRange: string;
      let sourceTable: any = null;

      if (sourceRange) {
        finalSourceRange = sourceRange;
      } else {
        sourceTable = context.findTable(tableId);
        if (!sourceTable) {
          throw new Error(
            `Table ${tableId || "primary"} not found for pivot table data`
          );
        }
        finalSourceRange = sourceTable.range;
      }

      // Handle legacy parameters by converting to new format
      let finalRows = [...rows];
      let finalColumns = [...columns];
      let finalValues = [...values];

      if (groupBy && !rows.includes(groupBy)) {
        finalRows.push(groupBy);
      }

      if (valueColumn && !values.some((v) => v.field === valueColumn)) {
        finalValues.push({
          field: valueColumn,
          aggregation: aggFunc,
        });
      }

      // Validate we have at least one row or column and one value
      if (finalRows.length === 0 && finalColumns.length === 0) {
        throw new Error(
          "At least one row or column field is required for pivot table"
        );
      }

      if (finalValues.length === 0) {
        throw new Error("At least one value field is required for pivot table");
      }

      console.log(`📊 Pivot configuration:`, {
        sourceRange: finalSourceRange,
        rows: finalRows,
        columns: finalColumns,
        values: finalValues,
      });

      // Determine destination
      const destRange =
        destination || context.findOptimalPlacement(8, 10).range;

      // Check if Univer has pivot table support
      const pivotTableSupport = checkPivotTableSupport(context.univerAPI);

      if (pivotTableSupport.hasNativePivotTable) {
        // Use native Univer pivot table API
        return await createNativePivotTable(context, {
          sourceRange: finalSourceRange,
          destination: destRange,
          rows: finalRows,
          columns: finalColumns,
          values: finalValues,
        });
      } else if (pivotTableSupport.hasDataPivot) {
        // Use Univer's data pivot functionality
        return await createDataPivot(context, {
          sourceRange: finalSourceRange,
          destination: destRange,
          rows: finalRows,
          columns: finalColumns,
          values: finalValues,
        });
      }
    } catch (error) {
      console.error(`❌ CreatePivotTableTool failed:`, error);
      throw error;
    }
  }
);

/**
 * Check what pivot table capabilities are available in Univer
 */
function checkPivotTableSupport(univerAPI: any): {
  hasNativePivotTable: boolean;
  hasDataPivot: boolean;
  hasPivotBuilder: boolean;
} {
  return {
    hasNativePivotTable:
      typeof univerAPI?.newPivotTable === "function" ||
      typeof univerAPI?.getActiveWorksheet?.().newPivotTable === "function",
    hasDataPivot: typeof univerAPI?.createDataPivot === "function",
    hasPivotBuilder: typeof univerAPI?.PivotTableBuilder !== "undefined",
  };
}

/**
 * Create pivot table using native Univer API
 */
async function createNativePivotTable(
  context: UniversalToolContext,
  config: {
    sourceRange: string;
    destination: string;
    rows: string[];
    columns: string[];
    values: Array<{ field: string; aggregation: string }>;
  }
): Promise<any> {
  console.log(`📊 Using native Univer pivot table API`);

  try {
    // Use Univer's native pivot table builder
    let pivotBuilder;

    if (typeof context.fWorksheet.newPivotTable === "function") {
      pivotBuilder = context.fWorksheet.newPivotTable();
    } else if (typeof context.univerAPI.newPivotTable === "function") {
      pivotBuilder = context.univerAPI.newPivotTable();
    } else {
      throw new Error("Native pivot table API not available");
    }

    // Set source range
    pivotBuilder = pivotBuilder.setSourceRange(config.sourceRange);

    // Add row fields
    for (const rowField of config.rows) {
      pivotBuilder = pivotBuilder.addRow(rowField);
    }

    // Add column fields
    for (const columnField of config.columns) {
      pivotBuilder = pivotBuilder.addColumn(columnField);
    }

    // Add value fields with aggregation
    for (const valueConfig of config.values) {
      const aggregation = getPivotAggregation(
        context.univerAPI,
        valueConfig.aggregation
      );
      pivotBuilder = pivotBuilder.addValue(valueConfig.field, aggregation);
    }

    // Set destination
    pivotBuilder = pivotBuilder.setDestination(config.destination);

    // Build and insert
    const pivotTable = pivotBuilder.build();
    const result = await context.fWorksheet.insertPivotTable(pivotTable);

    if (!result) {
      throw new Error("Failed to insert pivot table");
    }

    return {
      sourceRange: config.sourceRange,
      destination: config.destination,
      rows: config.rows,
      columns: config.columns,
      values: config.values,
      method: "native",
      message: `Created pivot table at ${config.destination} using native Univer API`,
      success: true,
    };
  } catch (error) {
    console.error("❌ Native pivot table creation failed:", error);
    throw error;
  }
}

/**
 * Create data pivot using Univer's data pivot functionality
 */
async function createDataPivot(
  context: UniversalToolContext,
  config: {
    sourceRange: string;
    destination: string;
    rows: string[];
    columns: string[];
    values: Array<{ field: string; aggregation: string }>;
  }
): Promise<any> {
  console.log(`📊 Using Univer data pivot functionality`);

  try {
    const pivotConfig = {
      dataRange: config.sourceRange,
      pivotRange: config.destination,
      rowFields: config.rows,
      columnFields: config.columns,
      dataFields: config.values.map((v) => ({
        field: v.field,
        aggregateFunction: v.aggregation.toUpperCase(),
      })),
    };

    const result = await context.univerAPI.createDataPivot(pivotConfig);

    if (!result) {
      throw new Error("Failed to create data pivot");
    }

    return {
      sourceRange: config.sourceRange,
      destination: config.destination,
      rows: config.rows,
      columns: config.columns,
      values: config.values,
      method: "data_pivot",
      message: `Created data pivot at ${config.destination}`,
      success: true,
    };
  } catch (error) {
    console.error("❌ Data pivot creation failed:", error);
    throw error;
  }
}

/**
 * Create manual pivot table using basic aggregation (fallback)
 */
async function createManualPivotTable(
  context: UniversalToolContext,
  config: {
    sourceRange: string;
    destination: string;
    rows: string[];
    columns: string[];
    values: Array<{ field: string; aggregation: string }>;
    sourceTable: any;
  }
): Promise<any> {
  console.log(`📊 Using manual pivot table creation (fallback)`);

  try {
    // This is a simplified pivot table creation
    // In a real implementation, this would analyze the source data and create aggregated results

    const destRange = context.fWorksheet.getRange(config.destination);

    // Create basic pivot table structure
    const headers = [
      ...config.rows.map((r) => `Group by ${r}`),
      ...config.values.map((v) => `${v.aggregation.toUpperCase()}(${v.field})`),
    ];

    // Set headers
    for (let i = 0; i < headers.length; i++) {
      const headerCell = destRange.offset(0, i, 1, 1);
      headerCell.setValue(headers[i]);
      headerCell.setFontWeight("bold");
    }

    // Add a note about the manual pivot
    const noteCell = destRange.offset(1, 0, 1, headers.length);
    noteCell.setValue(
      "Manual Pivot Table - Advanced pivot features require Univer Pro"
    );
    noteCell.setFontStyle("italic");

    return {
      sourceRange: config.sourceRange,
      destination: config.destination,
      rows: config.rows,
      columns: config.columns,
      values: config.values,
      method: "manual",
      message: `Created manual pivot table structure at ${config.destination} - Advanced features require Univer Pro`,
      success: true,
      warning:
        "This is a basic pivot table structure. Full pivot table functionality requires Univer Pro.",
    };
  } catch (error) {
    console.error("❌ Manual pivot table creation failed:", error);
    throw error;
  }
}

/**
 * Get Univer pivot aggregation enum
 */
function getPivotAggregation(univerAPI: any, aggFunc: string): any {
  const PivotAggregation = univerAPI?.Enum?.PivotAggregation;

  if (!PivotAggregation) {
    console.warn("⚠️ PivotAggregation enum not available, using string");
    return aggFunc.toLowerCase();
  }

  const aggMap: { [key: string]: any } = {
    sum: PivotAggregation.Sum || PivotAggregation.SUM,
    average: PivotAggregation.Average || PivotAggregation.AVERAGE,
    count: PivotAggregation.Count || PivotAggregation.COUNT,
    max: PivotAggregation.Max || PivotAggregation.MAX,
    min: PivotAggregation.Min || PivotAggregation.MIN,
    product: PivotAggregation.Product || PivotAggregation.PRODUCT,
    stddev: PivotAggregation.StdDev || PivotAggregation.STDDEV,
  };

  const mappedAgg = aggMap[aggFunc.toLowerCase()];
  if (!mappedAgg) {
    console.warn(`⚠️ Unknown aggregation function: ${aggFunc}, using Sum`);
    return PivotAggregation.Sum || PivotAggregation.SUM || "sum";
  }

  return mappedAgg;
}
