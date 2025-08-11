/**
 * Excel Function Tools - Comprehensive Function Support
 * 
 * Supports all common Excel functions without hardcoding, using intelligent 
 * context awareness and spatial analysis for dynamic range detection.
 */

import {
  createSimpleTool,
  type UniversalTool,
} from "../tool-executor";
import type { UniversalToolContext } from "../universal-context";

/**
 * SMART FORMULA BUILDER - Core Excel Formula Tool
 * 
 * Intelligently builds Excel formulas using context-aware range detection.
 * Supports all common Excel functions: SUM, AVERAGE, COUNT, IF, VLOOKUP, etc.
 */
export const SmartFormulaBuilderTool = createSimpleTool(
  {
    name: "smart_formula_builder",
    description: "Build Excel formulas intelligently with context-aware range detection",
    category: "formula",
    requiredContext: ["tables", "columns", "spatial"],
    invalidatesCache: false,
  },
  async (
    context: UniversalToolContext,
    params: {
      formula_type: string; // SUM, AVERAGE, COUNT, IF, VLOOKUP, etc.
      target_column?: string; // Column to apply formula to
      source_range?: string; // Override auto-detected range
      condition?: string; // For IF, COUNTIF, SUMIF
      lookup_value?: string; // For VLOOKUP, INDEX/MATCH
      lookup_table?: string; // For VLOOKUP
      output_cell?: string; // Where to place the formula
      tableId?: string;
    }
  ) => {
    const {
      formula_type,
      target_column,
      source_range,
      condition,
      lookup_value,
      lookup_table,
      output_cell,
      tableId,
    } = params;

    console.log(`🔧 Building ${formula_type} formula with intelligent context`);

    // Get table context
    const table = context.findTable(tableId);
    if (!table) {
      throw new Error(`Table ${tableId || "primary"} not found`);
    }

    // Build formula based on type
    const formulaResult = await buildIntelligentFormula(
      context,
      formula_type.toLowerCase(),
      {
        table,
        targetColumn: target_column,
        sourceRange: source_range,
        condition,
        lookupValue: lookup_value,
        lookupTable: lookup_table,
        outputCell: output_cell,
      }
    );

    // Place formula in optimal location if no output cell specified
    const finalOutputCell = output_cell || findOptimalFormulaPlacement(
      context,
      formulaResult.suggestedPlacement
    );

    // Apply the formula
    const result = await applyFormula(context, formulaResult.formula, finalOutputCell);

    return {
      formulaType: formula_type.toUpperCase(),
      formula: formulaResult.formula,
      outputCell: finalOutputCell,
      dataRange: formulaResult.dataRange,
      description: formulaResult.description,
      calculation: result,
      message: `Applied ${formula_type.toUpperCase()} formula: ${formulaResult.formula}`,
    };
  }
);

/**
 * Build intelligent formulas based on context
 */
async function buildIntelligentFormula(
  context: UniversalToolContext,
  formulaType: string,
  options: {
    table: any;
    targetColumn?: string;
    sourceRange?: string;
    condition?: string;
    lookupValue?: string;
    lookupTable?: string;
    outputCell?: string;
  }
) {
  const { table, targetColumn, sourceRange, condition, lookupValue, lookupTable } = options;

  switch (formulaType) {
    case "sum":
      return buildSumFormula(context, table, targetColumn, sourceRange);
    
    case "average":
    case "avg":
      return buildAverageFormula(context, table, targetColumn, sourceRange);
    
    case "count":
      return buildCountFormula(context, table, targetColumn, sourceRange);
    
    case "max":
      return buildMaxFormula(context, table, targetColumn, sourceRange);
    
    case "min":
      return buildMinFormula(context, table, targetColumn, sourceRange);
    
    case "if":
      return buildIfFormula(context, table, targetColumn, condition);
    
    case "sumif":
      return buildSumIfFormula(context, table, targetColumn, condition, sourceRange);
    
    case "countif":
      return buildCountIfFormula(context, table, targetColumn, condition, sourceRange);
    
    case "vlookup":
      return buildVLookupFormula(context, table, lookupValue, lookupTable, targetColumn);
    
    case "concatenate":
    case "concat":
      return buildConcatenateFormula(context, table, targetColumn, sourceRange);
    
    case "left":
      return buildLeftFormula(context, table, targetColumn, sourceRange);
    
    case "right":
      return buildRightFormula(context, table, targetColumn, sourceRange);
    
    case "mid":
      return buildMidFormula(context, table, targetColumn, sourceRange);
    
    case "len":
    case "length":
      return buildLenFormula(context, table, targetColumn, sourceRange);
    
    case "upper":
      return buildUpperFormula(context, table, targetColumn, sourceRange);
    
    case "lower":
      return buildLowerFormula(context, table, targetColumn, sourceRange);
    
    case "trim":
      return buildTrimFormula(context, table, targetColumn, sourceRange);
    
    case "round":
      return buildRoundFormula(context, table, targetColumn, sourceRange);
    
    case "abs":
      return buildAbsFormula(context, table, targetColumn, sourceRange);
    
    case "sqrt":
      return buildSqrtFormula(context, table, targetColumn, sourceRange);
    
    case "power":
      return buildPowerFormula(context, table, targetColumn, sourceRange);
    
    default:
      throw new Error(`Unsupported formula type: ${formulaType}. Supported: SUM, AVERAGE, COUNT, MAX, MIN, IF, SUMIF, COUNTIF, VLOOKUP, CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, ROUND, ABS, SQRT, POWER`);
  }
}

/**
 * Mathematical Functions
 */
function buildSumFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  return {
    formula: `=SUM(${range})`,
    dataRange: range,
    description: `Sum all values in ${range}`,
    suggestedPlacement: "below_data",
  };
}

function buildAverageFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  return {
    formula: `=AVERAGE(${range})`,
    dataRange: range,
    description: `Average of values in ${range}`,
    suggestedPlacement: "below_data",
  };
}

function buildCountFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "any");
  return {
    formula: `=COUNT(${range})`,
    dataRange: range,
    description: `Count non-empty values in ${range}`,
    suggestedPlacement: "below_data",
  };
}

function buildMaxFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  return {
    formula: `=MAX(${range})`,
    dataRange: range,
    description: `Maximum value in ${range}`,
    suggestedPlacement: "below_data",
  };
}

function buildMinFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  return {
    formula: `=MIN(${range})`,
    dataRange: range,
    description: `Minimum value in ${range}`,
    suggestedPlacement: "below_data",
  };
}

/**
 * Logical Functions
 */
function buildIfFormula(context: UniversalToolContext, table: any, targetColumn?: string, condition?: string) {
  const range = getIntelligentRange(context, table, targetColumn, "any");
  const defaultCondition = condition || `${range}>0`;
  return {
    formula: `=IF(${defaultCondition},"True","False")`,
    dataRange: range,
    description: `IF condition: ${defaultCondition}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildSumIfFormula(context: UniversalToolContext, table: any, targetColumn?: string, condition?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  const criteriaRange = getIntelligentRange(context, table, targetColumn, "any");
  const defaultCondition = condition || ">0";
  return {
    formula: `=SUMIF(${criteriaRange},"${defaultCondition}",${range})`,
    dataRange: range,
    description: `Sum values in ${range} where ${criteriaRange} meets "${defaultCondition}"`,
    suggestedPlacement: "below_data",
  };
}

function buildCountIfFormula(context: UniversalToolContext, table: any, targetColumn?: string, condition?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "any");
  const defaultCondition = condition || "*";
  return {
    formula: `=COUNTIF(${range},"${defaultCondition}")`,
    dataRange: range,
    description: `Count values in ${range} that meet "${defaultCondition}"`,
    suggestedPlacement: "below_data",
  };
}

/**
 * Lookup Functions
 */
function buildVLookupFormula(context: UniversalToolContext, table: any, lookupValue?: string, lookupTable?: string, targetColumn?: string) {
  const tableRange = lookupTable || table.range;
  const lookupVal = lookupValue || "A2"; // Default to first data cell
  const colIndex = targetColumn ? getColumnIndex(context, table, targetColumn) + 1 : 2;
  return {
    formula: `=VLOOKUP(${lookupVal},${tableRange},${colIndex},FALSE)`,
    dataRange: tableRange,
    description: `Lookup ${lookupVal} in ${tableRange}, return column ${colIndex}`,
    suggestedPlacement: "adjacent_column",
  };
}

/**
 * Text Functions
 */
function buildConcatenateFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const [startCell, endCell] = range.split(":");
  return {
    formula: `=CONCATENATE(${startCell}," ",${endCell || startCell})`,
    dataRange: range,
    description: `Concatenate values in ${range}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildLeftFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string, numChars: number = 5) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0]; // Use first cell
  return {
    formula: `=LEFT(${cellRef},${numChars})`,
    dataRange: range,
    description: `Extract ${numChars} characters from left of ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildRightFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string, numChars: number = 5) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0];
  return {
    formula: `=RIGHT(${cellRef},${numChars})`,
    dataRange: range,
    description: `Extract ${numChars} characters from right of ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildMidFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string, startPos: number = 2, numChars: number = 3) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0];
  return {
    formula: `=MID(${cellRef},${startPos},${numChars})`,
    dataRange: range,
    description: `Extract ${numChars} characters starting at position ${startPos} from ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildLenFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0];
  return {
    formula: `=LEN(${cellRef})`,
    dataRange: range,
    description: `Length of text in ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildUpperFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0];
  return {
    formula: `=UPPER(${cellRef})`,
    dataRange: range,
    description: `Convert ${cellRef} to uppercase`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildLowerFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0];
  return {
    formula: `=LOWER(${cellRef})`,
    dataRange: range,
    description: `Convert ${cellRef} to lowercase`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildTrimFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "text");
  const cellRef = range.split(":")[0];
  return {
    formula: `=TRIM(${cellRef})`,
    dataRange: range,
    description: `Remove extra spaces from ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

/**
 * Math Functions
 */
function buildRoundFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string, decimals: number = 2) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  const cellRef = range.split(":")[0];
  return {
    formula: `=ROUND(${cellRef},${decimals})`,
    dataRange: range,
    description: `Round ${cellRef} to ${decimals} decimal places`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildAbsFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  const cellRef = range.split(":")[0];
  return {
    formula: `=ABS(${cellRef})`,
    dataRange: range,
    description: `Absolute value of ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildSqrtFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  const cellRef = range.split(":")[0];
  return {
    formula: `=SQRT(${cellRef})`,
    dataRange: range,
    description: `Square root of ${cellRef}`,
    suggestedPlacement: "adjacent_column",
  };
}

function buildPowerFormula(context: UniversalToolContext, table: any, targetColumn?: string, sourceRange?: string, power: number = 2) {
  const range = sourceRange || getIntelligentRange(context, table, targetColumn, "numeric");
  const cellRef = range.split(":")[0];
  return {
    formula: `=POWER(${cellRef},${power})`,
    dataRange: range,
    description: `Raise ${cellRef} to power of ${power}`,
    suggestedPlacement: "adjacent_column",
  };
}

/**
 * Intelligent range detection without hardcoding
 */
function getIntelligentRange(
  context: UniversalToolContext,
  table: any,
  targetColumn?: string,
  dataType: "numeric" | "text" | "any" = "any"
): string {
  if (targetColumn) {
    // Use specific column
    const column = context.findColumn(targetColumn, table.id);
    if (column) {
      return context.getColumnRange(column.name, false, table.id); // Exclude header
    }
  }

  // Auto-detect best column based on data type preference
  let bestColumn;
  if (dataType === "numeric") {
    bestColumn = table.columns.find((c: any) => c.isNumeric || c.isCalculable);
  } else if (dataType === "text") {
    bestColumn = table.columns.find((c: any) => !c.isNumeric && !c.isCalculable);
  } else {
    bestColumn = table.columns[0]; // First column as fallback
  }

  if (bestColumn) {
    return context.getColumnRange(bestColumn.name, false, table.id);
  }

  // Ultimate fallback: use table data range (excluding header)
  const tableRange = table.range;
  const match = tableRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (match) {
    const [, startCol, startRow, endCol, endRow] = match;
    const dataStartRow = parseInt(startRow) + 1; // Skip header
    return `${startCol}${dataStartRow}:${endCol}${endRow}`;
  }

  return tableRange;
}

/**
 * Get column index for VLOOKUP
 */
function getColumnIndex(context: UniversalToolContext, table: any, columnName: string): number {
  const column = context.findColumn(columnName, table.id);
  return column ? column.index : 0;
}

/**
 * Find optimal placement for formula results
 */
function findOptimalFormulaPlacement(context: UniversalToolContext, suggestion: string): string {
  switch (suggestion) {
    case "below_data":
      // Find first empty row below the table
      const placement = context.findOptimalPlacement(1, 1);
      return placement.range.split(":")[0]; // Just the top-left cell
    
    case "adjacent_column":
      // Find empty column next to the table
      const adjacentPlacement = context.findOptimalPlacement(1, 1, "right");
      return adjacentPlacement.range.split(":")[0];
    
    default:
      const defaultPlacement = context.findOptimalPlacement(1, 1);
      return defaultPlacement.range.split(":")[0];
  }
}

/**
 * Apply formula to specified cell
 */
async function applyFormula(context: UniversalToolContext, formula: string, outputCell: string) {
  console.log(`📝 Applying formula ${formula} to ${outputCell}`);

  // Parse cell address
  const cellMatch = outputCell.match(/([A-Z]+)(\d+)/);
  if (!cellMatch) {
    throw new Error(`Invalid cell address: ${outputCell}`);
  }

  const [, colStr, rowStr] = cellMatch;
  const row = parseInt(rowStr) - 1; // Convert to 0-based
  const col = colStr.split("").reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;

  // Apply the formula
  const target = context.fWorksheet.getRange(row, col, 1, 1);
  
  if (typeof (target as any).setFormula === "function") {
    (target as any).setFormula(formula);
  } else {
    target.setValue(formula);
  }

  // Execute calculation if available
  try {
    const formulaEngine = context.univerAPI.getFormula();
    if (formulaEngine && typeof formulaEngine.executeCalculation === "function") {
      formulaEngine.executeCalculation();
    }
  } catch (calcError) {
    console.warn("⚠️ Formula calculation failed:", calcError);
  }

  // Get the calculated value
  let calculatedValue;
  try {
    calculatedValue = target.getValue();
  } catch {
    calculatedValue = formula; // Fallback to formula string
  }

  return {
    cell: outputCell,
    formula,
    value: calculatedValue,
  };
}

/**
 * BULK FORMULA APPLIER - Apply multiple formulas efficiently
 */
export const BulkFormulaApplierTool = createSimpleTool(
  {
    name: "bulk_formula_applier",
    description: "Apply multiple Excel formulas efficiently in batch",
    category: "formula",
    requiredContext: ["tables", "columns"],
    invalidatesCache: true,
  },
  async (
    context: UniversalToolContext,
    params: {
      formulas: Array<{
        type: string;
        target_column?: string;
        output_cell?: string;
        condition?: string;
        parameters?: any;
      }>;
      tableId?: string;
    }
  ) => {
    const { formulas, tableId } = params;
    const results = [];

    console.log(`🔢 Applying ${formulas.length} formulas in batch`);

    for (const formulaSpec of formulas) {
      try {
        const result = await buildIntelligentFormula(context, formulaSpec.type.toLowerCase(), {
          table: context.findTable(tableId),
          targetColumn: formulaSpec.target_column,
          condition: formulaSpec.condition,
          ...formulaSpec.parameters,
        });

        const outputCell = formulaSpec.output_cell || findOptimalFormulaPlacement(
          context,
          result.suggestedPlacement
        );

        const applied = await applyFormula(context, result.formula, outputCell);

        results.push({
          type: formulaSpec.type.toUpperCase(),
          ...applied,
          description: result.description,
        });
      } catch (error) {
        console.warn(`⚠️ Failed to apply ${formulaSpec.type} formula:`, error);
        results.push({
          type: formulaSpec.type.toUpperCase(),
          error: error instanceof Error ? error.message : "Unknown error",
        });
      }
    }

    const successful = results.filter(r => !r.error).length;
    const failed = results.length - successful;

    return {
      totalFormulas: formulas.length,
      successful,
      failed,
      results,
      message: `Applied ${successful} formulas successfully${failed > 0 ? `, ${failed} failed` : ''}`,
    };
  }
);

/**
 * Export all Excel function tools
 */
export const EXCEL_FUNCTION_TOOLS = [
  SmartFormulaBuilderTool,
  BulkFormulaApplierTool,
];