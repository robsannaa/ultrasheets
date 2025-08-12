/**
 * Aggregated Chart Tools - Business Intelligence Chart Creation
 * 
 * These tools understand business prompts like "total sales by product" 
 * and automatically perform the necessary aggregation steps.
 */

import {
  createSimpleTool,
  type UniversalTool,
} from "../tool-executor";
import type { UniversalToolContext } from "../universal-context";

/**
 * CREATE AGGREGATED CHART - Business Intelligence Chart Creation
 * 
 * Handles prompts like:
 * - "chart of total sales by product" → GROUP BY product, SUM sales
 * - "average rating by category" → GROUP BY category, AVG rating
 * - "revenue by month" → GROUP BY month, SUM revenue
 */
export const CreateAggregatedChartTool = createSimpleTool(
  {
    name: "create_aggregated_chart",
    description: "Create charts with intelligent data aggregation for business analysis",
    category: "analysis",
    requiredContext: ["tables", "columns", "spatial"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      chart_type?: string;
      title?: string;
      group_by_column: string; // The dimension to group by (e.g., "product", "category")
      value_column: string; // The metric to aggregate (e.g., "sales", "price")
      aggregation_type?: "sum" | "average" | "count" | "max" | "min"; // How to aggregate
      tableId?: string;
      position?: string;
      width?: number;
      height?: number;
    }
  ) => {
    const {
      chart_type = "column",
      title,
      group_by_column,
      value_column,
      aggregation_type = "sum",
      width = 500,
      height = 350,
      tableId,
    } = params;

    console.log("📊 Creating aggregated chart:", {
      group_by_column,
      value_column,
      aggregation_type,
      chart_type,
    });

    // Find the source table
    const table = context.findTable(tableId);
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    // Find the columns
    const groupByColumn = context.findColumn(group_by_column, tableId);
    const valueColumn = context.findColumn(value_column, tableId);

    if (!groupByColumn) {
      throw new Error(
        `Group by column "${group_by_column}" not found. Available: ${table.columns.map((c: any) => c.name).join(", ")}`
      );
    }
    if (!valueColumn) {
      throw new Error(
        `Value column "${value_column}" not found. Available: ${table.columns.map((c: any) => c.name).join(", ")}`
      );
    }

    console.log("✅ Columns found:", {
      groupBy: groupByColumn.name,
      value: valueColumn.name,
      groupByLetter: groupByColumn.letter,
      valueLetter: valueColumn.letter,
    });

    // Create the aggregated data using a pivot table approach
    const pivotResult = await createInMemoryPivotTable(
      context,
      table,
      groupByColumn,
      valueColumn,
      aggregation_type
    );

    // Find optimal placement for the chart
    const position =
      params.position ||
      (() => {
        const placement = context.findOptimalPlacement(
          Math.ceil(width / 60), 
          Math.ceil(height / 20)
        );
        return placement.range;
      })();

    // Generate chart title if not provided
    const chartTitle = title || generateChartTitle(
      aggregation_type,
      valueColumn.name,
      groupByColumn.name
    );

    // Place the aggregated data in the spreadsheet first
    const dataPlacement = await placeAggregatedData(
      context,
      pivotResult,
      position,
      chartTitle
    );

    // Create chart from the aggregated data
    const chartResult = await createChartFromAggregatedData(
      context,
      dataPlacement,
      chart_type,
      chartTitle,
      width,
      height
    );

    return {
      aggregation: {
        groupByColumn: groupByColumn.name,
        valueColumn: valueColumn.name,
        aggregationType: aggregation_type,
        uniqueGroups: pivotResult.groups.length,
      },
      data: {
        range: dataPlacement.dataRange,
        groups: pivotResult.groups,
        values: pivotResult.values,
      },
      chart: chartResult,
      message: `Created ${chart_type} chart "${chartTitle}" showing ${aggregation_type} of ${valueColumn.name} by ${groupByColumn.name} (${pivotResult.groups.length} categories)`,
    };
  }
);

/**
 * Create in-memory pivot table for aggregation
 */
async function createInMemoryPivotTable(
  context: UniversalToolContext,
  table: any,
  groupByColumn: any,
  valueColumn: any,
  aggregationType: string
) {
  console.log("🔄 Creating in-memory pivot table...");

  // Get the raw data from the table
  const dataStartRow = table.position.startRow + 1; // Skip header
  const dataEndRow = table.position.endRow;
  
  const aggregationMap = new Map<string, number[]>();
  
  // Process each data row
  for (let row = dataStartRow; row <= dataEndRow; row++) {
    const groupCell = context.fWorksheet.getRange(row, groupByColumn.index, 1, 1);
    const valueCell = context.fWorksheet.getRange(row, valueColumn.index, 1, 1);
    
    const groupValue = String(groupCell.getValue() || "").trim();
    const numericValue = parseFloat(String(valueCell.getValue())) || 0;
    
    if (groupValue && !isNaN(numericValue)) {
      if (!aggregationMap.has(groupValue)) {
        aggregationMap.set(groupValue, []);
      }
      aggregationMap.get(groupValue)!.push(numericValue);
    }
  }

  // Perform aggregation
  const result = {
    groups: [] as string[],
    values: [] as number[],
  };

  for (const [group, values] of aggregationMap.entries()) {
    let aggregatedValue: number;
    
    switch (aggregationType) {
      case "sum":
        aggregatedValue = values.reduce((sum, val) => sum + val, 0);
        break;
      case "average":
        aggregatedValue = values.reduce((sum, val) => sum + val, 0) / values.length;
        break;
      case "count":
        aggregatedValue = values.length;
        break;
      case "max":
        aggregatedValue = Math.max(...values);
        break;
      case "min":
        aggregatedValue = Math.min(...values);
        break;
      default:
        aggregatedValue = values.reduce((sum, val) => sum + val, 0);
    }

    result.groups.push(group);
    result.values.push(aggregatedValue);
  }

  // Sort by group name for consistency
  const sortedIndices = result.groups
    .map((_, index) => index)
    .sort((a, b) => result.groups[a].localeCompare(result.groups[b]));

  const sortedGroups = sortedIndices.map(i => result.groups[i]);
  const sortedValues = sortedIndices.map(i => result.values[i]);

  console.log("✅ Aggregation complete:", {
    totalGroups: sortedGroups.length,
    groups: sortedGroups.slice(0, 5), // Log first 5
    values: sortedValues.slice(0, 5),
  });

  return {
    groups: sortedGroups,
    values: sortedValues,
  };
}

/**
 * Place aggregated data in the spreadsheet
 */
async function placeAggregatedData(
  context: UniversalToolContext,
  pivotResult: { groups: string[], values: number[] },
  position: string,
  title: string
) {
  console.log("📍 Placing aggregated data at:", position);

  // Parse position
  const posCol = (position.match(/[A-Z]+/i)?.[0] || "H").toUpperCase();
  const posRowNum = parseInt(position.replace(/\D+/g, ""), 10) || 2;
  const startRow = posRowNum - 1; // Convert to 0-based
  const startCol = posCol.charCodeAt(0) - 65; // Convert to 0-based

  // Place title
  context.fWorksheet.getRange(startRow, startCol, 1, 1).setValue(`📊 ${title}`);

  // Place headers
  context.fWorksheet.getRange(startRow + 2, startCol, 1, 1).setValue("Category");
  context.fWorksheet.getRange(startRow + 2, startCol + 1, 1, 1).setValue("Value");

  // Place data
  for (let i = 0; i < pivotResult.groups.length; i++) {
    const dataRow = startRow + 3 + i;
    context.fWorksheet.getRange(dataRow, startCol, 1, 1).setValue(pivotResult.groups[i]);
    context.fWorksheet.getRange(dataRow, startCol + 1, 1, 1).setValue(pivotResult.values[i]);
  }

  const dataStartRow = startRow + 2; // Headers start here
  const dataEndRow = startRow + 2 + pivotResult.groups.length; // Last data row
  const dataRange = `${posCol}${dataStartRow + 1}:${String.fromCharCode(65 + startCol + 1)}${dataEndRow + 1}`;

  console.log("✅ Data placed in range:", dataRange);

  return {
    titleCell: `${posCol}${startRow + 1}`,
    dataRange,
    headerRow: dataStartRow,
    dataRows: pivotResult.groups.length,
  };
}

/**
 * Create chart from aggregated data
 */
async function createChartFromAggregatedData(
  context: UniversalToolContext,
  dataPlacement: any,
  chartType: string,
  title: string,
  width: number,
  height: number
) {
  console.log("📊 Creating chart from aggregated data...");

  // Chart position (offset from data)
  const dataCol = dataPlacement.dataRange.match(/[A-Z]+/i)?.[0] || "H";
  const dataColIndex = dataCol.charCodeAt(0) - 65;
  const chartColIndex = dataColIndex + 4; // Place chart 4 columns to the right
  const chartCol = String.fromCharCode(65 + chartColIndex);
  const chartRow = dataPlacement.headerRow;

  const chartPosition = `${chartCol}${chartRow + 1}`;

  // Map chart types
  const univerChartTypeMap: { [key: string]: string } = {
    column: "Column",
    bar: "Bar", 
    line: "Line",
    pie: "Pie",
    scatter: "Scatter",
  };

  const univerChartType = univerChartTypeMap[chartType] || "Column";
  const enumType = (context.univerAPI as any).Enum?.ChartType || {};
  const chartTypeEnum = enumType[univerChartType] || enumType.Column;

  // Parse chart position
  const anchorRow = chartRow;
  const anchorCol = chartColIndex;

  try {
    // Create chart using Univer API
    const builder = context.fWorksheet
      .newChart()
      .setChartType(chartTypeEnum)
      .addRange(dataPlacement.dataRange)
      .setPosition(anchorRow, anchorCol, 0, 0)
      .setWidth(width)
      .setHeight(height);

    if (title) {
      builder.setOptions("title.text", title);
    }

    const chartInfo = builder.build();
    await context.fWorksheet.insertChart(chartInfo);

    console.log("✅ Chart created successfully");

    return {
      type: univerChartType,
      position: chartPosition,
      dataRange: dataPlacement.dataRange,
      title,
      width,
      height,
    };
  } catch (error) {
    console.warn("⚠️ Native chart creation failed, using fallback");
    
    // Fallback: Place a chart placeholder
    context.fWorksheet.getRange(anchorRow, anchorCol, 1, 1).setValue(`📊 ${title} Chart`);
    context.fWorksheet.getRange(anchorRow + 1, anchorCol, 1, 1).setValue(`Data: ${dataPlacement.dataRange}`);

    return {
      type: univerChartType,
      position: chartPosition,
      dataRange: dataPlacement.dataRange,
      title,
      width,
      height,
      fallback: true,
    };
  }
}

/**
 * Generate appropriate chart title based on aggregation
 */
function generateChartTitle(aggregationType: string, valueColumn: string, groupColumn: string): string {
  const aggregationLabels: { [key: string]: string } = {
    sum: "Total",
    average: "Average", 
    count: "Count of",
    max: "Maximum",
    min: "Minimum",
  };

  const label = aggregationLabels[aggregationType] || "Total";
  return `${label} ${valueColumn} by ${groupColumn}`;
}

/**
 * SMART CHART ANALYZER - Automatically determines the best aggregation
 * 
 * For prompts like "chart sales by product", this tool intelligently:
 * 1. Identifies the group column (product)
 * 2. Identifies the value column (sales)  
 * 3. Determines the appropriate aggregation (usually sum for sales)
 */
export const SmartChartAnalyzerTool = createSimpleTool(
  {
    name: "smart_chart_analyzer",
    description: "Automatically analyze user chart requests and determine optimal aggregation strategy",
    category: "analysis",
    requiredContext: ["tables", "columns"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      user_prompt: string; // e.g., "chart of total sales by product"
      tableId?: string;
    }
  ) => {
    const { user_prompt, tableId } = params;
    const table = context.findTable(tableId);
    
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    console.log("🔍 Analyzing chart prompt:", user_prompt);

    // Extract chart intentions from natural language
    const analysis = analyzeChartPrompt(user_prompt, table);

    return {
      interpretation: analysis,
      recommendedTool: "create_aggregated_chart",
      recommendedParams: {
        chart_type: analysis.chartType,
        title: analysis.title,
        group_by_column: analysis.groupByColumn,
        value_column: analysis.valueColumn,
        aggregation_type: analysis.aggregationType,
        tableId: tableId,
      },
      message: `Analyzed: "${user_prompt}" → ${analysis.aggregationType} of ${analysis.valueColumn} by ${analysis.groupByColumn}`,
    };
  }
);

/**
 * Analyze chart prompt using business intelligence patterns
 */
function analyzeChartPrompt(prompt: string, table: any) {
  const lower = prompt.toLowerCase();
  const columns = table.columns.map((c: any) => ({ name: c.name, lower: c.name.toLowerCase() }));

  // Detect aggregation type
  let aggregationType = "sum"; // default
  if (lower.includes("average") || lower.includes("avg")) aggregationType = "average";
  else if (lower.includes("count")) aggregationType = "count";
  else if (lower.includes("max") || lower.includes("maximum")) aggregationType = "max";
  else if (lower.includes("min") || lower.includes("minimum")) aggregationType = "min";
  else if (lower.includes("total") || lower.includes("sum")) aggregationType = "sum";

  // Detect chart type
  let chartType = "column"; // default
  if (lower.includes("pie")) chartType = "pie";
  else if (lower.includes("line")) chartType = "line";
  else if (lower.includes("bar")) chartType = "bar";
  else if (lower.includes("scatter")) chartType = "scatter";

  // Detect "by" keyword for grouping
  const byMatch = lower.match(/by\s+(\w+)/);
  let groupByColumn = null;

  if (byMatch) {
    const groupCandidate = byMatch[1];
    groupByColumn = columns.find((c: { name: string; lower: string }) => c.lower.includes(groupCandidate))?.name;
  }

  // If no "by" found, look for common grouping words
  if (!groupByColumn) {
    const groupingKeywords = ["product", "category", "region", "month", "day", "type", "brand"];
    for (const keyword of groupingKeywords) {
      if (lower.includes(keyword)) {
        groupByColumn = columns.find((c: { name: string; lower: string }) => c.lower.includes(keyword))?.name;
        if (groupByColumn) break;
      }
    }
  }

  // Detect value column (what to aggregate)
  let valueColumn = null;
  const valueKeywords = ["sales", "revenue", "price", "cost", "amount", "value", "rating", "score"];
  
  for (const keyword of valueKeywords) {
    if (lower.includes(keyword)) {
      valueColumn = columns.find((c: { name: string; lower: string }) => c.lower.includes(keyword))?.name;
      if (valueColumn) break;
    }
  }

  // Fallbacks
  if (!groupByColumn) {
    // Use first text/categorical column
    groupByColumn = columns.find((c: { name: string; lower: string }) => 
      !["price", "cost", "amount", "value", "rating", "score"].some(k => c.lower.includes(k))
    )?.name || columns[0]?.name;
  }

  if (!valueColumn) {
    // Use first numeric-sounding column
    valueColumn = columns.find((c: { name: string; lower: string }) => 
      ["price", "cost", "amount", "value", "rating", "score", "sales"].some(k => c.lower.includes(k))
    )?.name || columns[columns.length - 1]?.name;
  }

  const title = `${aggregationType.charAt(0).toUpperCase() + aggregationType.slice(1)} ${valueColumn || 'Values'} by ${groupByColumn || 'Category'}`;

  return {
    chartType,
    aggregationType,
    groupByColumn,
    valueColumn,
    title,
    confidence: calculateConfidence(lower, groupByColumn, valueColumn),
  };
}

/**
 * Calculate confidence score for the analysis
 */
function calculateConfidence(prompt: string, groupBy: string | null, value: string | null): number {
  let score = 0.5; // base confidence
  
  if (prompt.includes("by")) score += 0.3;
  if (groupBy) score += 0.1;
  if (value) score += 0.1;
  if (prompt.includes("total") || prompt.includes("sum") || prompt.includes("average")) score += 0.1;
  
  return Math.min(1.0, score);
}

/**
 * Export all aggregated chart tools
 */
export const AGGREGATED_CHART_TOOLS = [
  CreateAggregatedChartTool,
  SmartChartAnalyzerTool,
];