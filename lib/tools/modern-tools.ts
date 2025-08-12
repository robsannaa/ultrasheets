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
    console.log(
      `🔍 AddFilterTool: Looking for table with ID: "${params.tableId}"`
    );
    console.log(
      `🔍 Available tables:`,
      context.tables.map((t) => ({ id: t.id, range: t.range }))
    );
    console.log(
      `🔍 Primary table:`,
      context.primaryTable?.id,
      context.primaryTable?.range
    );

    const table = context.findTable(params.tableId);
    if (!table) {
      console.log(
        `❌ AddFilterTool: Table not found. Available table IDs: ${context.tables
          .map((t) => t.id)
          .join(", ")}`
      );

      // If no specific table ID and no primary table, try to use the first available table
      if (!params.tableId && context.tables.length > 0) {
        console.log(
          `🔄 AddFilterTool: No primary table found, using first available table: ${context.tables[0].id}`
        );
        const fallbackTable = context.tables[0];

        const fallbackTableRange = fallbackTable.range;
        const fallbackFRange = context.fWorksheet.getRange(fallbackTableRange);

        // Remove existing filter if any
        let fallbackFFilter = fallbackFRange.getFilter();
        if (fallbackFFilter) {
          fallbackFFilter.remove();
        }

        // Create new filter
        fallbackFFilter = fallbackFRange.createFilter();

        return {
          success: true,
          tableId: fallbackTable.id,
          range: fallbackTableRange,
          message: `Added filter to table ${fallbackTable.id} at range ${fallbackTableRange} (fallback)`,
          filterApplied: true,
        };
      }

      throw new Error(
        `Table ${
          params.tableId || "primary"
        } not found. Available tables: ${context.tables
          .map((t) => t.id)
          .join(", ")}`
      );
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
        console.log(
          `❌ SortTool: Table not found. Available table IDs: ${context.tables
            .map((t) => t.id)
            .join(", ")}`
        );
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

/**
 * Export all modern tools
 */
export const MODERN_TOOLS = [
  AddSmartTotalsTool,
  AddFilterTool,
  FormatCurrencyColumnTool,
  GenerateChartTool,
  GetWorkbookSnapshotTool,
  SortTool,
];
