/**
 * Modern Tool Implementations using Universal Context Framework
 *
 * These tools demonstrate the new architecture and can serve as templates
 * for converting existing tools and creating new ones.
 */

import {
  UniversalTool,
  createSimpleTool,
  type ToolDefinition,
} from "../tool-executor";
import type { UniversalToolContext } from "../universal-context";

/**
 * ADD SMART TOTALS - Modern Implementation
 */
export const AddSmartTotalsTool = createSimpleTool(
  {
    name: "add_smart_totals",
    description:
      "Automatically add totals to all calculable columns in a table",
    category: "data",
    requiredContext: ["tables", "columns"],
    invalidatesCache: true,
  },
  async (
    context: UniversalToolContext,
    params: { columns?: string[]; tableId?: string }
  ) => {
    const table = context.findTable(params.tableId);
    if (!table) {
      throw new Error(`Table ${params.tableId || "primary"} not found`);
    }

    // Determine which columns to total
    let columnsToTotal =
      params.columns && params.columns.length > 0
        ? table.columns.filter((c: any) =>
            params.columns!.some(
              (reqCol) =>
                c.name.toLowerCase() === reqCol.toLowerCase() ||
                c.letter === reqCol.toUpperCase()
            )
          )
        : table.columns.filter((c: any) => c.isCalculable || c.isNumeric);

    if (columnsToTotal.length === 0) {
      throw new Error("No calculable columns found to total");
    }

    const results = [];
    const sumRow = table.position.endRow + 1;

    // Add totals for each column
    for (const column of columnsToTotal) {
      const sumFormula = context.buildSumFormula(column.name, params.tableId);
      const sumCellA1 = `${column.letter}${sumRow + 1}`;

      // Set the formula
      const target = context.fWorksheet.getRange(sumRow, column.index, 1, 1);
      if (typeof (target as any).setFormula === "function") {
        (target as any).setFormula(sumFormula);
      } else {
        target.setValue(sumFormula);
      }

      results.push({
        column: column.name,
        cell: sumCellA1,
        formula: sumFormula,
        columnType: column.dataType,
        isCurrency: column.isCurrency,
      });

      // Track for later formatting
      if (!(window as any).recentTotals) (window as any).recentTotals = [];
      (window as any).recentTotals.push({
        cell: sumCellA1,
        column: column.name,
        columnType: column.dataType,
        isCurrency: column.isCurrency,
        tableId: table.id,
        timestamp: Date.now(),
      });
    }

    // Execute calculation
    try {
      const formula = context.univerAPI.getFormula();
      formula.executeCalculation();
    } catch (calcError) {
      console.warn("⚠️ Formula calculation failed:", calcError);
    }

    return {
      tableId: table.id,
      totalsAdded: results.length,
      results,
      message: `Added ${results.length} totals to table "${table.id}": ${results
        .map((r) => r.column)
        .join(", ")}`,
    };
  }
);

/**
 * ADD FILTER - Modern Implementation
 */
export const AddFilterTool = createSimpleTool(
  {
    name: "add_filter",
    description: "Add Excel-like filter dropdowns to table headers",
    category: "structure",
    requiredContext: ["tables"],
    invalidatesCache: false,
  },
  async (context: UniversalToolContext, params: { tableId?: string }) => {
    const table = context.findTable(params.tableId);
    if (!table) {
      throw new Error(`Table ${params.tableId || "primary"} not found`);
    }

    const tableRange = table.range;
    const fRange = context.fWorksheet.getRange(tableRange);

    // Remove existing filter if any
    let fFilter = fRange.getFilter();
    if (fFilter) {
      fFilter.remove();
    }

    // Create new filter
    fFilter = fRange.createFilter();
    if (!fFilter) {
      throw new Error("Failed to create filter on the table range");
    }

    return {
      tableId: table.id,
      range: tableRange,
      headers: table.headers,
      message: `Filter added to table "${table.id}" (${tableRange}). Filter dropdowns are now available in the header row.`,
    };
  }
);

/**
 * FORMAT CURRENCY COLUMN - Modern Implementation
 */
export const FormatCurrencyColumnTool = createSimpleTool(
  {
    name: "format_currency_column",
    description: "Format a currency column with proper currency formatting",
    category: "format",
    requiredContext: ["tables", "columns"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      currency?: string;
      decimals?: number;
      columnName?: string;
      tableId?: string;
    }
  ) => {
    const { currency = "USD", decimals = 2, columnName, tableId } = params;

    const table = context.findTable(tableId);
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    // Find target column
    let targetColumn;
    if (columnName) {
      targetColumn = context.findColumn(columnName, tableId);
    } else {
      // FIRST: prefer the most recently added column if it looks like currency-related
      try {
        const w: any = window as any;
        const recent = Array.isArray(w.ultraActionLog)
          ? [...w.ultraActionLog].reverse()
          : [];
        const lastAdded = recent.find(
          (a: any) => a?.tool === "smart_add_column" && a?.params?.columnName
        );
        if (lastAdded?.params?.columnName) {
          const candidate = context.findColumn(
            lastAdded.params.columnName,
            tableId
          );
          if (candidate) {
            targetColumn = candidate;
          }
        }
      } catch {}

      // Auto-detect currency column as fallback
      if (!targetColumn) {
        targetColumn = table.columns.find(
          (c: any) =>
            c.isCurrency ||
            c.dataType === "currency" ||
            ["price", "cost", "amount", "revenue", "value", "sales"].some(
              (keyword) => c.name.toLowerCase().includes(keyword)
            )
        );
      }
    }

    if (!targetColumn) {
      throw new Error(
        `No currency column found. Available columns: ${table.columns
          .map((c: any) => c.name)
          .join(", ")}`
      );
    }

    // Get data range (excluding header)
    const dataRange = context.getColumnRange(targetColumn.name, false, tableId);

    // Currency format patterns
    const currencyFormats: { [key: string]: string } = {
      USD: "$#,##0.00",
      EUR: "€#,##0.00",
      GBP: "£#,##0.00",
      JPY: "¥#,##0",
      PLN: "#,##0.00zł",
      CAD: "C$#,##0.00",
      AUD: "A$#,##0.00",
      CHF: "CHF#,##0.00",
      CNY: "¥#,##0.00",
      INR: "₹#,##0.00",
    };

    const formatPattern =
      currencyFormats[currency.toUpperCase()] ||
      `${currency}#,##0.${"0".repeat(decimals)}`;

    // Apply formatting
    const fRange = context.fWorksheet.getRange(dataRange);
    fRange.setNumberFormat(formatPattern);

    return {
      column: targetColumn.name,
      range: dataRange,
      currency,
      decimals,
      formatPattern,
      message: `Formatted column "${targetColumn.name}" (${dataRange}) as ${currency} currency`,
    };
  }
);

/**
 * GENERATE CHART - Modern Implementation
 */
export const GenerateChartTool = createSimpleTool(
  {
    name: "generate_chart",
    description: "Generate charts using intelligent table detection",
    category: "analysis",
    requiredContext: ["tables", "spatial"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      chart_type?: string;
      title?: string;
      x_column?: string;
      y_columns?: string[];
      data_range?: string;
      position?: string;
      width?: number;
      height?: number;
      tableId?: string;
    }
  ) => {
    const {
      chart_type = "column",
      title = "Chart",
      width = 400,
      height = 300,
      tableId,
    } = params;

    // Determine data range
    let finalDataRange;
    if (params.data_range) {
      finalDataRange = params.data_range;
    } else {
      const table = context.findTable(tableId);
      if (!table) {
        throw new Error(
          `Table ${tableId || "primary"} not found for chart data`
        );
      }
      finalDataRange = table.range;
    }

    // Determine chart position using spatial analysis
    const position =
      params.position ||
      (() => {
        const placement = context.findOptimalPlacement(width / 60, height / 20); // Rough cell conversion
        return placement.range;
      })();

    // Chart type mapping
    const univerChartTypeMap: { [key: string]: string } = {
      column: "Column",
      line: "Line",
      pie: "Pie",
      bar: "Bar",
      scatter: "Scatter",
    };

    const univerChartType = univerChartTypeMap[chart_type] || "Column";
    const enumType = (context.univerAPI as any).Enum?.ChartType || {};
    const chartTypeEnum = enumType[univerChartType] || enumType.Column;

    // Parse position
    const posCol = (position.match(/[A-Z]+/i)?.[0] || "H").toUpperCase();
    const posRowNum = parseInt(position.replace(/\D+/g, ""), 10) || 2;
    const anchorRow = posRowNum - 1;
    const anchorCol = posCol.charCodeAt(0) - 65;

    // Create chart
    const builder = context.fWorksheet
      .newChart()
      .setChartType(chartTypeEnum)
      .addRange(finalDataRange)
      .setPosition(anchorRow, anchorCol, 0, 0)
      .setWidth(width)
      .setHeight(height);

    if (title) builder.setOptions("title.text", title);

    const chartInfo = builder.build();
    await context.fWorksheet.insertChart(chartInfo);

    return {
      chartType: univerChartType,
      dataRange: finalDataRange,
      position,
      title,
      width,
      height,
      message: `Created ${chart_type} chart "${title}" using data range ${finalDataRange} at ${position}`,
    };
  }
);

/**
 * GET WORKBOOK SNAPSHOT - Modern Implementation
 */
export const GetWorkbookSnapshotTool = createSimpleTool(
  {
    name: "get_workbook_snapshot",
    description: "Get complete workbook context with intelligent analysis",
    category: "navigation",
    requiredContext: [], // This tool builds the context
    invalidatesCache: false,
  },
  async (context: UniversalToolContext, params: {}) => {
    // Force a fresh context to ensure latest data
    const freshContext = await context.refresh();

    return {
      activeSheet: freshContext.activeSheetName,
      snapshot: freshContext.activeSheetSnapshot,
      intelligence: freshContext.intelligence,
      summary: {
        totalTables: freshContext.tables.length,
        totalColumns: freshContext.tables.reduce(
          (sum, table) => sum + table.columns.length,
          0
        ),
        calculableColumns: freshContext.calculableColumns,
        numericColumns: freshContext.numericColumns,
        spatialZones:
          freshContext.spatialMap?.optimalPlacementZones?.length || 0,
      },
      message: `Retrieved complete workbook context for ${freshContext.activeSheetName}`,
    };
  }
);

/**
 * Export all modern tools
 */
export const MODERN_TOOLS = [
  AddSmartTotalsTool,
  AddFilterTool,
  FormatCurrencyColumnTool,
  GenerateChartTool,
  GetWorkbookSnapshotTool,
];
