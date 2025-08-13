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
import type { UniversalToolContext } from "../tool-executor";
import { getWorkbookData } from "../univer-data-source";

/**
 * ADD SMART TOTALS - Modern Implementation
 */
export const AddSmartTotalsTool = createSimpleTool(
  {
    name: "add_totals",
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
    // Resolve table robustly using the context's findTable method
    let table = context.findTable(params.tableId);
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
      // Build SUM formula directly using column range
      const colLetter = String.fromCharCode(65 + column.index);
      const startRow = table.position.startRow + 2; // Skip header
      const endRow = table.position.endRow + 1;
      const dataRange = `${colLetter}${startRow}:${colLetter}${endRow}`;
      const sumFormula = `=SUM(${dataRange})`;
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
    console.log(
      `🔍 AddFilterTool: Looking for table with ID: "${params.tableId}"`
    );

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
    description:
      "Format numeric monetary columns (not dates or text) with proper currency formatting. Intelligently detects currency columns by name and data type.",
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
          (a: any) => a?.tool === "add_column" && a?.params?.columnName
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

      // Auto-detect currency column as fallback with enhanced intelligence
      if (!targetColumn) {
        // Enhanced currency column detection
        const currencyColumns = table.columns.filter((c: any) => {
          const columnName = c.name.toLowerCase();
          const hasMoneyKeywords = [
            "price",
            "cost",
            "amount",
            "revenue",
            "value",
            "sales",
            "total",
            "income",
            "expense",
            "profit",
            "fee",
            "charge",
            "bill",
            "payment",
            "money",
            "dollar",
            "euro",
            "currency",
            "cash",
            "budget",
            "financial",
          ].some((keyword) => columnName.includes(keyword));

          const hasMoneySymbols = /[\$€£¥₹]/.test(columnName);

          // Check if column has predominantly numeric data that looks like money
          const hasNumericData =
            c.dataType === "number" ||
            (c.sampleValues &&
              c.sampleValues.some(
                (v: any) => typeof v === "number" && v > 0 && v < 1000000
              ));

          // Exclude obvious non-currency columns
          const isExcluded = [
            "date",
            "time",
            "id",
            "index",
            "count",
            "quantity",
            "number",
            "year",
            "month",
            "day",
            "week",
            "serial",
            "code",
            "zip",
            "phone",
          ].some((keyword) => columnName.includes(keyword));

          return (
            (hasMoneyKeywords ||
              hasMoneySymbols ||
              c.isCurrency ||
              c.dataType === "currency") &&
            !isExcluded &&
            hasNumericData
          );
        });

        // Prefer the first currency column found
        targetColumn = currencyColumns[0];
      }
    }

    if (!targetColumn) {
      throw new Error(
        `No currency column found. Available columns: ${table.columns
          .map((c: any) => c.name)
          .join(", ")}`
      );
    }

    // Validate that the selected column is appropriate for currency formatting
    const targetColumnName = targetColumn.name.toLowerCase();
    const isDateColumn = [
      "date",
      "time",
      "created",
      "updated",
      "timestamp",
    ].some((keyword) => targetColumnName.includes(keyword));

    if (isDateColumn) {
      throw new Error(
        `Cannot format "${targetColumn.name}" as currency - this appears to be a date/time column. ` +
          `Available numeric columns: ${table.columns
            .filter(
              (c: any) =>
                c.dataType === "number" ||
                !["date", "time", "created", "updated", "timestamp"].some(
                  (keyword) => c.name.toLowerCase().includes(keyword)
                )
            )
            .map((c: any) => c.name)
            .join(", ")}`
      );
    }

    // Get data range (excluding header) - build directly
    const colLetter = String.fromCharCode(65 + targetColumn.index);
    const startRow = table.position.startRow + 2; // Skip header
    const endRow = table.position.endRow + 1;
    const dataRange = `${colLetter}${startRow}:${colLetter}${endRow}`;

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
 * GENERATE CHART - Enhanced Implementation with proper Univer builder pattern
 */
export const GenerateChartTool = createSimpleTool(
  {
    name: "generate_chart",
    description: "Generate charts using proper Univer builder pattern with intelligent data detection",
    category: "analysis",
    requiredContext: ["tables"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      chart_type?: string;
      title?: string;
      x_column?: string;
      y_columns?: string[] | string;
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
      x_column,
      y_columns
    } = params;

    console.log(`📊 GenerateChartTool: Creating ${chart_type} chart "${title}"`);

    try {
      // Determine data range with intelligent filtering
      let finalDataRange: string;
      let sourceTable: any = null;

      if (params.data_range) {
        finalDataRange = params.data_range;
        console.log(`📊 Using explicit data range: ${finalDataRange}`);
      } else {
        sourceTable = context.findTable(tableId);
        if (!sourceTable) {
          throw new Error(`Table ${tableId || "primary"} not found for chart data`);
        }
        
        // If specific columns are requested, create a filtered range
        if (x_column || y_columns) {
          finalDataRange = await createFilteredDataRange(context, sourceTable, x_column, y_columns);
          console.log(`📊 Created filtered data range: ${finalDataRange}`);
        } else {
          finalDataRange = sourceTable.range;
          console.log(`📊 Using full table range: ${finalDataRange}`);
        }
      }

      // Determine optimal chart position
      const position = params.position
        ? params.position
        : context.findOptimalPlacement(Math.ceil(width / 64), Math.ceil(height / 20)).range;

      console.log(`📊 Chart position: ${position}`);

      // Get proper Univer chart type enum
      const univerChartType = getUniverChartType(context.univerAPI, chart_type);
      console.log(`📊 Chart type enum:`, univerChartType);

      // Parse position coordinates
      const { anchorRow, anchorCol } = parseChartPosition(position);

      // Create chart using proper builder pattern
      const chartBuilder = context.fWorksheet.newChart();
      
      if (!chartBuilder) {
        throw new Error("Failed to create chart builder - newChart() returned null");
      }

      // Configure chart with method chaining
      let configuredBuilder = chartBuilder
        .setChartType(univerChartType)
        .addRange(finalDataRange)
        .setPosition(anchorRow, anchorCol, 0, 0)
        .setWidth(width)
        .setHeight(height);

      // Set chart options if title provided
      if (title) {
        // Try different methods to set title based on Univer version
        try {
          configuredBuilder = configuredBuilder.setOptions({ title: { text: title } });
        } catch (e1) {
          try {
            configuredBuilder = configuredBuilder.setOptions("title.text", title);
          } catch (e2) {
            try {
              configuredBuilder = configuredBuilder.setTitle(title);
            } catch (e3) {
              console.warn('⚠️ Could not set chart title, continuing without it');
            }
          }
        }
      }

      // Build the chart configuration
      const chartInfo = configuredBuilder.build();
      
      if (!chartInfo) {
        throw new Error("Chart builder returned null - build() failed");
      }

      console.log(`📊 Built chart config:`, chartInfo);

      // Insert chart into worksheet
      const insertResult = await context.fWorksheet.insertChart(chartInfo);
      
      if (insertResult === false) {
        throw new Error("Failed to insert chart into worksheet");
      }

      console.log(`✅ GenerateChartTool: Successfully created ${chart_type} chart`);

      return {
        chartType: chart_type,
        univerChartType: univerChartType,
        dataRange: finalDataRange,
        position,
        anchorRow,
        anchorCol,
        title,
        width,
        height,
        x_column,
        y_columns,
        tableId: sourceTable?.id,
        message: `Created ${chart_type} chart "${title}" using data range ${finalDataRange} at ${position}`,
        success: true
      };

    } catch (error) {
      console.error(`❌ GenerateChartTool failed:`, error);
      throw error;
    }
  }
);

/**
 * Get proper Univer chart type enum
 */
function getUniverChartType(univerAPI: any, chartType: string): any {
  const ChartType = univerAPI?.Enum?.ChartType;
  
  if (!ChartType) {
    console.warn('⚠️ ChartType enum not available, using fallback');
    return chartType; // Fallback to string
  }

  const typeMapping: { [key: string]: any } = {
    column: ChartType.Column || ChartType.COLUMN,
    line: ChartType.Line || ChartType.LINE,
    pie: ChartType.Pie || ChartType.PIE,
    bar: ChartType.Bar || ChartType.BAR,
    scatter: ChartType.Scatter || ChartType.SCATTER,
  };

  const mappedType = typeMapping[chartType.toLowerCase()];
  
  if (!mappedType) {
    console.warn(`⚠️ Unknown chart type: ${chartType}, using Column`);
    return ChartType.Column || ChartType.COLUMN || 'Column';
  }

  return mappedType;
}

/**
 * Parse chart position string to row/column coordinates
 */
function parseChartPosition(position: string): { anchorRow: number; anchorCol: number } {
  const colMatch = position.match(/[A-Z]+/i);
  const rowMatch = position.match(/\d+/);
  
  const posCol = (colMatch?.[0] || "H").toUpperCase();
  const posRowNum = parseInt(rowMatch?.[0] || "2", 10);
  
  // Convert to 0-based indices
  const anchorRow = Math.max(0, posRowNum - 1);
  const anchorCol = Math.max(0, posCol.charCodeAt(0) - 65);
  
  return { anchorRow, anchorCol };
}

/**
 * Create filtered data range based on specified columns
 */
async function createFilteredDataRange(
  _context: UniversalToolContext,
  table: any,
  x_column?: string,
  y_columns?: string[] | string
): Promise<string> {
  // For now, return the full table range
  // In a more advanced implementation, this would create a new range with only specified columns
  let targetColumns: string[] = [];
  
  if (x_column) {
    targetColumns.push(x_column);
  }
  
  if (y_columns) {
    const yColArray = Array.isArray(y_columns) ? y_columns : [y_columns];
    targetColumns.push(...yColArray);
  }
  
  if (targetColumns.length === 0) {
    return table.range;
  }
  
  // Find the columns in the table
  const matchedColumns = targetColumns
    .map(colName => table.columns.find((c: any) => 
      c.name.toLowerCase().includes(colName.toLowerCase()) || 
      c.letter === colName.toUpperCase()
    ))
    .filter(Boolean);
  
  if (matchedColumns.length === 0) {
    console.warn('⚠️ No matching columns found, using full table range');
    return table.range;
  }
  
  // For advanced filtering, we would construct a range with only these columns
  // For now, return the full table range
  console.log(`📊 Matched columns for chart: ${matchedColumns.map(c => c.name).join(', ')}`);
  return table.range;
}

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
  async (context: UniversalToolContext, _params: {}) => {
    // Get fresh data directly from Univer API
    const workbookData = getWorkbookData();
    console.log("🔍 GetWorkbookSnapshotTool: Fresh workbook data:", workbookData);

    return {
      activeSheet: workbookData?.activeSheetName || 'unknown',
      workbookData,
      summary: {
        totalSheets: workbookData?.sheets?.length || 0,
        activeSheet: workbookData?.activeSheetName || 'unknown'
      },
      message: `Retrieved workbook data for ${workbookData?.activeSheetName || 'unknown sheet'}`,
    };
  }
);

/**
 * SORTING TOOL - Modern Implementation using Univer Sheets Sort API
 * Using the official Univer sorting facade API: FRange.sort() and FWorksheet.sort()
 */
export const SortTool = createSimpleTool(
  {
    name: "sort_table",
    description:
      "Sort table data by specified column in ascending or descending order",
    category: "structure",
    requiredContext: ["tables"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      tableId?: string;
      column: string | number;
      ascending?: boolean;
      range?: string;
    }
  ) => {
    const { tableId, column, ascending = true, range } = params;

    console.log(
      `🔍 SortTool: Sorting by column: ${column}, ascending: ${ascending}`
    );

    let targetRange: string;
    let table: any = null;

    if (range) {
      // Use explicit range
      targetRange = range;
    } else {
      // Find table and use its range
      table = context.findTable(tableId);
      if (!table) {
        console.log(`❌ SortTool: Table not found`);
        throw new Error(`Table ${tableId || "primary"} not found`);
      }
      targetRange = table.range;
    }

    console.log(`🔍 SortTool: Using range: ${targetRange}`);

    // Convert column to index for sorting
    let colIndex: number;
    let columnName = column;

    if (typeof column === "string" && column.match(/^[A-Z]+$/)) {
      // Convert column letter to index (A=0, B=1, etc.)
      colIndex =
        column
          .split("")
          .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
    } else if (typeof column === "string") {
      // Try to find column by header name
      if (!table) table = context.findTable(tableId);
      const columnObj = table ? context.findColumn(column, tableId) : null;
      if (columnObj) {
        colIndex = columnObj.index; // Use the column index directly
        columnName = columnObj.name;
        console.log(
          `🔍 SortTool: Found column "${columnName}" at index ${colIndex}`
        );
      } else {
        throw new Error(
          `Column "${column}" not found in table. Available columns: ${
            table?.columns?.map((c: any) => c.name).join(", ") || "none"
          }`
        );
      }
    } else {
      colIndex = column;
    }

    console.log(
      `🔍 SortTool: Sorting by column index: ${colIndex}, ascending: ${ascending}`
    );

    // Use Univer's official FRange.sort() API - following docs religiously, no fallbacks
    const fRange = context.fWorksheet.getRange(targetRange);
    const sortResult = fRange.sort({ column: colIndex, ascending });

    console.log(`✅ SortTool: Successfully sorted ${targetRange}`);

    return {
      success: true,
      range: targetRange,
      column: columnName,
      columnIndex: colIndex,
      ascending,
      method: "range",
      message: `Sorted ${targetRange} by column ${columnName} in ${
        ascending ? "ascending" : "descending"
      } order`,
      result: sortResult,
    };
  }
);

// Unified tool assembly happens in lib/tools/index.ts (domain modules). No grouped export here.
