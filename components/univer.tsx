"use client";

import { useEffect, useRef } from "react";
import { useTheme } from "next-themes";
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-drawing/lib/index.css";
import "@univerjs/preset-sheets-advanced/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
// Enable formula facade API (VLOOKUP/INDEX/MATCH and 500+ functions)
import "@univerjs/sheets-formula/facade";
// Global tool execution handler
declare global {
  interface Window {
    executeUniverTool: (toolName: string, params?: any) => Promise<any>;
    lastSumCell?: string; // Track the last sum cell created
    __ultraUniverInitialized?: boolean;
    __ultraUniverAPI?: any;
    __ultraUniverLifecycleHookAdded?: boolean;
    __ultraUniverDispose?: () => void;
  }
}

export function Univer() {
  const containerRef = useRef<HTMLDivElement>(null!);
  const { theme } = useTheme();
  useEffect(() => {
    // Dynamic imports to avoid SSR issues
    const initializeUniver = async () => {
      const { UniverSheetsCorePreset } = await import(
        "@univerjs/preset-sheets-core"
      );
      const { UniverSheetsDrawingPreset } = await import(
        "@univerjs/preset-sheets-drawing"
      );
      const { UniverSheetsAdvancedPreset } = await import(
        "@univerjs/preset-sheets-advanced"
      );
      const { UniverSheetsConditionalFormattingPreset } = await import(
        "@univerjs/preset-sheets-conditional-formatting"
      );
      // Use CalculationMode from sheets-formula to align with facade API
      const { CalculationMode } = await import("@univerjs/sheets-formula");
      const sheetsCoreEnUS = await import(
        "@univerjs/preset-sheets-core/locales/en-US"
      ).then((m) => m.default);
      const sheetsDrawingEnUS = await import(
        "@univerjs/preset-sheets-drawing/locales/en-US"
      ).then((m) => m.default);
      const sheetsAdvancedEnUS = await import(
        "@univerjs/preset-sheets-advanced/locales/en-US"
      ).then((m) => m.default);
      const sheetsCFEnUS = await import(
        "@univerjs/preset-sheets-conditional-formatting/locales/en-US"
      ).then((m) => m.default);
      const {
        createUniver,
        LocaleType,
        mergeLocales,
        defaultTheme,
        greenTheme,
      } = await import("@univerjs/presets");
      const { LifecycleStages } = await import("@univerjs/core");

      // Note: Advanced and drawing presets cause dependency conflicts, keeping simple for now

      // Try to import filter preset, fallback if it fails
      let filterPreset = null;
      let filterLocales = null;
      try {
        const filterModule = await import("@univerjs/preset-sheets-filter");
        const filterLocaleModule = await import(
          "@univerjs/preset-sheets-filter/locales/en-US"
        );
        filterPreset = filterModule.UniverSheetsFilterPreset;
        filterLocales = filterLocaleModule.default;

        // Import filter CSS - try multiple possible paths
        // Optional CSS; skip if not present in node_modules to avoid build errors
        // Intentionally skip importing optional CSS to avoid type resolution errors in builds
        console.log("✅ Filter preset loaded successfully");
      } catch (error) {
        console.warn(
          "⚠️ Filter preset not available, continuing without filters:",
          error
        );
      }

      const corePreset = UniverSheetsCorePreset({
        container: containerRef.current,
        // Ensure formulas compute deterministically on load
        formula: {
          initialFormulaComputing: CalculationMode.FORCED,
        },
      });

      // Build presets in a way that allows graceful fallback if advanced/drawing cause injector errors
      let presets: any[] = [corePreset];
      try {
        // Try to include chart-related presets
        presets = [
          corePreset,
          UniverSheetsDrawingPreset(),
          UniverSheetsAdvancedPreset(),
          UniverSheetsConditionalFormattingPreset(),
        ];
      } catch (e) {
        console.warn(
          "⚠️ Failed to prepare chart presets; continuing with core only.",
          e
        );
        presets = [corePreset];
      }

      // Add filter preset if available
      if (filterPreset) {
        presets.push(filterPreset());
      }

      let locales = mergeLocales(sheetsCoreEnUS);
      try {
        locales = filterLocales
          ? mergeLocales(
              sheetsCoreEnUS,
              sheetsDrawingEnUS,
              sheetsAdvancedEnUS,
              sheetsCFEnUS,
              filterLocales
            )
          : mergeLocales(
              sheetsCoreEnUS,
              sheetsDrawingEnUS,
              sheetsAdvancedEnUS,
              sheetsCFEnUS
            );
      } catch (e) {
        console.warn(
          "⚠️ Failed to merge locales for chart presets; using core locales only.",
          e
        );
        locales = filterLocales
          ? mergeLocales(sheetsCoreEnUS, filterLocales)
          : mergeLocales(sheetsCoreEnUS);
      }

      // If already initialized (HMR / re-render), reuse existing API
      if (
        typeof window !== "undefined" &&
        window.__ultraUniverInitialized &&
        window.__ultraUniverAPI
      ) {
        // Remove Univer watermark if API available (HMR-safe)
        try {
          if (typeof (window.__ultraUniverAPI as any).removeWatermark === "function") {
            (window.__ultraUniverAPI as any).removeWatermark();
          }
        } catch {}

        // Re-bind tool executor and exit early
        const univerAPI = window.__ultraUniverAPI;
        (window as any).univerAPI = univerAPI;
        window.executeUniverTool = async (toolName: string, params?: any) => {
          try {
            switch (toolName) {
              case "list_columns":
                return await executeListColumns();
              case "calculate_total":
                return await executeCalculateTotal(params);
              case "create_pivot_table":
                return await executeCreatePivotTable(params);
              case "generate_chart":
                return await executeGenerateChart(params);
              case "format_currency":
                return await executeFormatCurrency(params);
              case "switch_sheet":
                return await executeSwitchSheet(params);
              case "add_filter":
                return await executeAddFilter(params);
              case "conditional_formatting":
                return await executeConditionalFormatting(params);
              default:
                throw new Error(`Unknown tool: ${toolName}`);
            }
          } catch (error) {
            console.error(`Tool execution failed: ${toolName}`, error);
            throw error;
          }
        };
        console.log(
          "🎯 Univer component: Reusing existing univerAPI (guarded re-init)"
        );
        return;
      }

      const { univerAPI, univer } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: locales,
        },
        presets,
        darkMode: theme === "dark",
        theme: theme === "dark" ? greenTheme : defaultTheme,
      });

      // Remove Univer watermark if API available
      try {
        if (typeof (univerAPI as any).removeWatermark === "function") {
          (univerAPI as any).removeWatermark();
        }
      } catch {}

      // Force initial formula computing at lifecycle Starting for future workbooks
      try {
        if (!window.__ultraUniverLifecycleHookAdded) {
          univerAPI.addEvent(
            univerAPI.Event.LifeCycleChanged,
            ({ stage }: { stage: any }) => {
              if (stage === LifecycleStages.Starting) {
                const formula = univerAPI.getFormula();
                formula.setInitialFormulaComputing(CalculationMode.FORCED);
              }
            }
          );
          window.__ultraUniverLifecycleHookAdded = true;
        }
        // Sync dark mode on mount and when theme changes
        try {
          if (typeof univerAPI.toggleDarkMode === "function") {
            univerAPI.toggleDarkMode(theme === "dark");
          }
        } catch {}
      } catch {}

      // Create workbook
      univerAPI.createWorkbook({});

      // EXPOSE univerAPI GLOBALLY for chat component and guard flags
      (window as any).univerAPI = univerAPI;
      window.__ultraUniverAPI = univerAPI;
      window.__ultraUniverInitialized = true;

      // Provide a dispose hook (not used during HMR reuse, but available)
      window.__ultraUniverDispose = () => {
        try {
          univer.dispose?.();
        } catch {}
        window.__ultraUniverInitialized = false;
        window.__ultraUniverAPI = undefined;
        (window as any).univerAPI = undefined;
      };

      // Set up global tool execution handler
      window.executeUniverTool = async (toolName: string, params?: any) => {
        try {
          switch (toolName) {
            case "list_columns":
              return await executeListColumns();
            case "calculate_total":
              return await executeCalculateTotal(params);
            case "create_pivot_table":
              return await executeCreatePivotTable(params);
            case "generate_chart":
              return await executeGenerateChart(params);
            case "format_currency":
              return await executeFormatCurrency(params);
            case "switch_sheet":
              return await executeSwitchSheet(params);
            case "add_filter":
              return await executeAddFilter(params);
            case "conditional_formatting":
              return await executeConditionalFormatting(params);
            default:
              throw new Error(`Unknown tool: ${toolName}`);
          }
        } catch (error) {
          console.error(`Tool execution failed: ${toolName}`, error);
          throw error;
        }
      };

      // Tool execution functions
      const executeConditionalFormatting = async (params: any) => {
        if (!univerAPI) throw new Error("Univer API not available");
        const fWorkbook: any = univerAPI.getActiveWorkbook();
        if (!fWorkbook) throw new Error("No active workbook available");
        const fWorksheet: any = fWorkbook.getActiveSheet();

        const {
          range = "A1:A100",
          ruleType,
          // Numeric comparisons
          min,
          max,
          equals,
          // Text comparisons
          contains,
          startsWith,
          endsWith,
          // Formula
          formula,
          // Style
          background,
          fontColor,
          bold,
          italic,
        } = params || {};

        const fRange = fWorksheet.getRange(range);
        const builder = fWorksheet.newConditionalFormattingRule();

        // Choose condition
        switch (ruleType) {
          case "number_between":
            builder.whenNumberBetween(min, max);
            break;
          case "number_gt":
            builder.whenNumberGreaterThan(min);
            break;
          case "number_gte":
            builder.whenNumberGreaterThanOrEqualTo(min);
            break;
          case "number_lt":
            builder.whenNumberLessThan(max);
            break;
          case "number_lte":
            builder.whenNumberLessThanOrEqualTo(max);
            break;
          case "number_eq":
            builder.whenNumberEqualTo(equals);
            break;
          case "number_neq":
            builder.whenNumberNotEqualTo(equals);
            break;
          case "text_contains":
            builder.whenTextContains(contains);
            break;
          case "text_not_contains":
            builder.whenTextDoesNotContain(contains);
            break;
          case "text_starts_with":
            builder.whenTextStartsWith(startsWith);
            break;
          case "text_ends_with":
            builder.whenTextEndsWith(endsWith);
            break;
          case "not_empty":
            builder.whenCellNotEmpty();
            break;
          case "empty":
            builder.whenCellEmpty();
            break;
          case "formula":
            if (formula) builder.whenFormulaSatisfied(formula);
            break;
          case "color_scale":
            builder.setColorScale("green-yellow-red");
            break;
          case "data_bar":
            builder.setDataBar();
            break;
          case "unique":
            builder.setUniqueValues();
            break;
          case "duplicate":
            builder.setDuplicateValues();
            break;
          default:
            // Default useful rule: negatives red text
            builder.whenNumberLessThan(0);
        }

        // Apply formats
        if (background) builder.setBackground(background);
        if (fontColor) builder.setFontColor(fontColor);
        if (bold !== undefined) builder.setBold(!!bold);
        if (italic !== undefined) builder.setItalic(!!italic);

        const rule = builder.setRanges([fRange.getRange()]).build();
        fWorksheet.addConditionalFormattingRule(rule);

        return {
          success: true,
          message: `Added conditional formatting on ${range} (${
            ruleType || "number_lt 0"
          })`,
          range,
          ruleType: ruleType || "number_lt",
        };
      };
      const executeListColumns = async () => {
        if (!univerAPI) {
          throw new Error("Univer API not available");
        }

        const workbook = univerAPI.getActiveWorkbook();
        if (!workbook) {
          throw new Error("No active workbook available");
        }

        const worksheet = workbook.getActiveSheet();
        if (!worksheet) {
          throw new Error("No active worksheet available");
        }

        try {
          // Get worksheet snapshot using proper Univer API
          const sheetSnapshot = worksheet.getSheet().getSnapshot();
          console.log("📊 list_columns: Got sheet snapshot:", {
            hasSnapshot: !!sheetSnapshot,
            name: sheetSnapshot?.name,
            rowCount: sheetSnapshot?.rowCount,
            columnCount: sheetSnapshot?.columnCount,
            hasCellData: !!sheetSnapshot?.cellData,
          });

          if (!sheetSnapshot || !sheetSnapshot.cellData) {
            return {
              error: "No data found",
              message:
                "The spreadsheet appears to be empty or the data hasn't loaded yet.",
              columns: [],
              rowCount: 0,
              sheetName: sheetSnapshot?.name || "Sheet1",
            };
          }

          // Find column headers - they might not be in row 0
          const cellData = sheetSnapshot.cellData;
          const columns: string[] = [];
          let headerRow = -1;

          // Debug: Log first few rows to see structure
          console.log("🔍 list_columns: Analyzing sheet structure:");
          for (
            let row = 0;
            row < Math.min(3, sheetSnapshot.rowCount || 3);
            row++
          ) {
            const rowData = cellData[row];
            if (rowData) {
              const rowValues = Object.keys(rowData)
                .map((col) => (rowData as Record<string, any>)[col]?.v)
                .filter((v) => v !== undefined);
              console.log(`  Row ${row}:`, rowValues);
            } else {
              console.log(`  Row ${row}: empty`);
            }
          }

          // Find the row with the most consecutive text values (likely headers)
          for (
            let row = 0;
            row < Math.min(5, sheetSnapshot.rowCount || 5);
            row++
          ) {
            const rowData = cellData[row] || {};
            const tempColumns: string[] = [];

            for (let col = 0; col < (sheetSnapshot.columnCount || 26); col++) {
              const cell = rowData[col];
              if (
                cell &&
                cell.v !== undefined &&
                cell.v !== null &&
                cell.v !== "" &&
                typeof cell.v === "string" &&
                !cell.f // Not a formula
              ) {
                tempColumns.push(String(cell.v));
              } else if (tempColumns.length > 0) {
                break; // Stop when we hit an empty cell after finding headers
              }
            }

            // If this row has more columns than our current best, use it
            if (
              tempColumns.length > columns.length &&
              tempColumns.length >= 2
            ) {
              columns.splice(0, columns.length, ...tempColumns);
              headerRow = row;
              console.log(
                `📋 Found better header row ${row} with ${tempColumns.length} columns:`,
                tempColumns
              );
            }
          }

          // Count rows with data (starting after header row)
          let rowCount = 0;
          const dataStartRow = headerRow + 1;

          for (
            let row = dataStartRow;
            row < (sheetSnapshot.rowCount || 1000);
            row++
          ) {
            const rowData = cellData[row];
            if (rowData) {
              let hasData = false;
              // Check if row has data in any of the header columns
              for (let col = 0; col < columns.length; col++) {
                const cell = rowData[col];
                if (
                  cell &&
                  cell.v !== undefined &&
                  cell.v !== null &&
                  cell.v !== ""
                ) {
                  hasData = true;
                  break;
                }
              }
              if (hasData) {
                rowCount++; // Count actual data rows, not row index
              }
            }
          }

          console.log(
            `📊 list_columns: Found ${columns.length} columns and ${rowCount} rows`
          );

          return {
            columns,
            rowCount,
            sheetName: sheetSnapshot.name || "Sheet1",
            message: `Found ${columns.length} columns and ${rowCount} rows of data.`,
          };
        } catch (error) {
          console.error(
            "❌ list_columns: Error accessing worksheet data:",
            error
          );
          throw new Error(
            `Failed to access worksheet data: ${
              error instanceof Error ? error.message : "Unknown error"
            }`
          );
        }
      };

      // Helpers for range/column parsing
      const parseA1Range = (a1: string) => {
        const [start, end] = a1.split(":");
        const colToIndex = (s: string) => s.toUpperCase().charCodeAt(0) - 65; // A=0
        const startCol = colToIndex(start.replace(/\d+/g, ""));
        const startRow = parseInt(start.replace(/\D+/g, ""), 10) - 1; // 0-based
        const endCol = colToIndex(end.replace(/\d+/g, ""));
        const endRow = parseInt(end.replace(/\D+/g, ""), 10) - 1;
        return { startRow, startCol, endRow, endCol };
      };

      const colLetterToOffset = (letter: string) =>
        letter.toUpperCase().charCodeAt(0) - 65; // A=0

      const executeCalculateTotal = async (params: any) => {
        console.log("🔍 calculate_total: Starting execution...", params);

        // Check readiness using the local univerAPI instance
        if (!univerAPI) {
          console.log("❌ calculate_total: Univer API not available");
          throw new Error("Univer API not available");
        }

        const workbook = univerAPI.getActiveWorkbook();
        if (!workbook) {
          console.log("❌ calculate_total: No active workbook");
          throw new Error("No active workbook available");
        }

        const worksheet = workbook.getActiveSheet();
        if (!worksheet) {
          console.log("❌ calculate_total: No active worksheet");
          throw new Error("No active worksheet available");
        }

        console.log(
          "✅ calculate_total: Univer is ready, accessing worksheet data..."
        );

        const { column, data_range, tableId } = params;

        try {
          // Get worksheet snapshot using proper Univer API
          const sheetSnapshot = worksheet.getSheet().getSnapshot();

          if (!sheetSnapshot || !sheetSnapshot.cellData) {
            return {
              error: "No data found",
              message: "The spreadsheet appears to be empty.",
            };
          }

          const cellData = sheetSnapshot.cellData;

          // If a data_range or tableId is provided, restrict detection to that range
          let scopedStartRow = 0;
          let scopedEndRow = sheetSnapshot.rowCount || 1000;
          let scopedStartCol = 0;
          let scopedEndCol = (sheetSnapshot.columnCount || 26) - 1;

          const scopedRange =
            data_range || (tableId && String(tableId).split(":")[1]);
          if (scopedRange && /^[A-Z]+\d+:[A-Z]+\d+$/i.test(scopedRange)) {
            const { startRow, endRow, startCol, endCol } =
              parseA1Range(scopedRange);
            scopedStartRow = startRow;
            scopedEndRow = endRow;
            scopedStartCol = startCol;
            scopedEndCol = endCol;
          }

          // Find headers using the same logic as list_columns
          const headers: string[] = [];
          let headerRow = -1;

          // Find the row with the most consecutive text values (likely headers)
          for (
            let row = scopedStartRow;
            row <= Math.min(scopedStartRow + 4, scopedEndRow);
            row++
          ) {
            const rowData = cellData[row] || {};
            const tempColumns: string[] = [];

            for (let col = scopedStartCol; col <= scopedEndCol; col++) {
              const cell = rowData[col];
              if (
                cell &&
                cell.v !== undefined &&
                cell.v !== null &&
                cell.v !== "" &&
                typeof cell.v === "string" &&
                !cell.f
              ) {
                tempColumns.push(String(cell.v));
              } else if (tempColumns.length > 0) {
                break;
              }
            }

            if (
              tempColumns.length > headers.length &&
              tempColumns.length >= 2
            ) {
              headers.splice(0, headers.length, ...tempColumns);
              headerRow = row; // absolute row index
            }
          }

          console.log("📊 calculate_total: Available headers:", headers);

          // Find column index
          let columnIndex = -1;
          if (/^[A-Z]+$/.test(column)) {
            // Interpret column letters relative to scoped range if provided
            const offset = colLetterToOffset(column);
            columnIndex = scopedStartCol + offset;
          } else {
            // Column name provided (e.g., 'Search_Volume')
            const relIdx = headers.findIndex((h) =>
              h.toLowerCase().includes(column.toLowerCase())
            );
            if (relIdx >= 0) columnIndex = scopedStartCol + relIdx;
          }

          if (columnIndex === -1) {
            return {
              error: "Column not found",
              message: `Column '${column}' not found. Available columns: ${headers.join(
                ", "
              )}`,
            };
          }

          console.log(
            `📊 calculate_total: Found column at index ${columnIndex}`
          );

          // Calculate total from the data
          let total = 0;
          let count = 0;
          let dataRowsProcessed = 0;

          // Process all rows (starting after header row)
          const dataStartRow = headerRow + 1;
          for (let row = dataStartRow; row <= scopedEndRow; row++) {
            const rowData = cellData[row];
            if (rowData) {
              const cell = rowData[columnIndex];
              if (
                cell &&
                cell.v !== undefined &&
                cell.v !== null &&
                cell.v !== ""
              ) {
                const value = parseFloat(String(cell.v)) || 0;
                if (!isNaN(value)) {
                  total += value;
                  count++;
                }
                dataRowsProcessed++;
              }
            }
          }

          console.log(
            `📊 calculate_total: Processed ${dataRowsProcessed} rows, found ${count} numeric values, total: ${total}`
          );

          // ACTUALLY PLACE THE SUM IN THE SPREADSHEET
          // First, check if there's already a sum in this column
          let lastDataRow = dataStartRow;
          let existingSumRow = -1;

          for (let row = dataStartRow; row <= scopedEndRow; row++) {
            const rowData = cellData[row];
            if (rowData && rowData[columnIndex]) {
              const cell = rowData[columnIndex];
              if (
                cell &&
                cell.v !== undefined &&
                cell.v !== null &&
                cell.v !== ""
              ) {
                // Check if this cell contains a SUM formula
                if (cell.f && String(cell.f).includes("SUM(")) {
                  existingSumRow = row;
                  console.log(
                    `📊 Found existing SUM formula in row ${row + 1}`
                  );
                } else {
                  // This is data, update lastDataRow
                  lastDataRow = row;
                }
              }
            }
          }

          let sumRow: number;
          let sumCell: string;
          const columnLetter = String.fromCharCode(65 + columnIndex);

          if (existingSumRow >= 0) {
            // Update existing sum
            sumRow = existingSumRow;
            sumCell = `${columnLetter}${sumRow + 1}`;
            console.log(`📍 Updating existing SUM formula in cell ${sumCell}`);
          } else {
            // Create new sum
            sumRow = lastDataRow + 2; // Place sum 2 rows below last data
            sumCell = `${columnLetter}${sumRow + 1}`; // +1 because Univer uses 1-based indexing for display
            console.log(`📍 Creating new SUM formula in cell ${sumCell}`);
          }

          // Create SUM formula range
          const dataRange = `${columnLetter}${
            dataStartRow + 1
          }:${columnLetter}${lastDataRow + 1}`;
          const sumFormula = `=SUM(${dataRange})`;

          // Place the formula in the spreadsheet using Univer API
          const sumRange = worksheet.getRange(sumRow, columnIndex, 1, 1);
          sumRange.setValue(sumFormula);
          // Execute formula calculation to avoid stale or flickering values
          try {
            const formula = univerAPI.getFormula();
            formula.executeCalculation();
          } catch {}

          console.log(
            `✅ calculate_total: Successfully placed sum in ${sumCell}`
          );

          // Store the sum cell globally for formatting operations
          window.lastSumCell = sumCell;

          return {
            total,
            count,
            average: count > 0 ? total / count : 0,
            column: headers[columnIndex] || column,
            columnIndex,
            dataRowsProcessed,
            sumCell,
            formula: sumFormula,
            dataRange,
            message: `Calculated total of ${total} from ${count} numeric values in column '${
              headers[columnIndex] || column
            }' and placed SUM formula in cell ${sumCell}`,
          };
        } catch (error) {
          console.error("❌ calculate_total: Error processing data:", error);
          throw new Error(
            `Failed to calculate total: ${
              error instanceof Error ? error.message : "Unknown error"
            }`
          );
        }
      };

      const executeCreatePivotTable = async (params: any) => {
        console.log("🔍 executeCreatePivotTable: Starting...", params);

        if (!univerAPI) {
          throw new Error("Univer API not available");
        }

        const {
          groupBy,
          valueColumn,
          aggFunc,
          destination,
          sheetName,
          data_range,
          tableId,
        } = params;

        try {
          const workbook = univerAPI.getActiveWorkbook();
          if (!workbook) {
            throw new Error("No active workbook available");
          }
          let worksheet: any = workbook.getActiveSheet();
          // Create/switch sheet if requested
          try {
            if (
              sheetName &&
              typeof (workbook as any).getSheetByName === "function"
            ) {
              const existing = (workbook as any).getSheetByName(sheetName);
              if (existing) {
                worksheet = existing;
              } else if (typeof (workbook as any).insertSheet === "function") {
                worksheet = (workbook as any).insertSheet(sheetName);
              }
            }
          } catch (e) {
            console.warn("⚠️ Sheet create/switch not supported:", e);
          }

          // Get worksheet snapshot
          const sheetSnapshot = worksheet.getSheet().getSnapshot();
          if (!sheetSnapshot || !sheetSnapshot.cellData) {
            throw new Error("No data found in spreadsheet");
          }

          const cellData = sheetSnapshot.cellData;

          // Scope to range if provided
          const parseA1Range = (a1: string) => {
            const [start, end] = a1.split(":");
            const colToIndex = (s: string) =>
              s.toUpperCase().charCodeAt(0) - 65; // A=0
            const startCol = colToIndex(start.replace(/\d+/g, ""));
            const startRow = parseInt(start.replace(/\D+/g, ""), 10) - 1; // 0-based
            const endCol = colToIndex(end.replace(/\d+/g, ""));
            const endRow = parseInt(end.replace(/\D+/g, ""), 10) - 1;
            return { startRow, startCol, endRow, endCol };
          };
          let scopedStartRow = 0;
          let scopedEndRow = sheetSnapshot.rowCount || 1000;
          let scopedStartCol = 0;
          let scopedEndCol = (sheetSnapshot.columnCount || 26) - 1;
          const scopedRange =
            data_range || (tableId && String(tableId).split(":")[1]);
          if (scopedRange && /^[A-Z]+\d+:[A-Z]+\d+$/i.test(scopedRange)) {
            const { startRow, endRow, startCol, endCol } =
              parseA1Range(scopedRange);
            scopedStartRow = startRow;
            scopedEndRow = endRow;
            scopedStartCol = startCol;
            scopedEndCol = endCol;
          }

          // Find headers using same logic as other tools
          const headers: string[] = [];
          let headerRow = -1;

          for (
            let row = scopedStartRow;
            row <= Math.min(scopedStartRow + 4, scopedEndRow);
            row++
          ) {
            const rowData = cellData[row] || {};
            const tempColumns: string[] = [];

            for (let col = scopedStartCol; col <= scopedEndCol; col++) {
              const cell = rowData[col];
              if (cell && typeof cell.v === "string" && !cell.f) {
                tempColumns.push(String(cell.v));
              } else if (tempColumns.length > 0) {
                break;
              }
            }

            if (
              tempColumns.length > headers.length &&
              tempColumns.length >= 2
            ) {
              headers.splice(0, headers.length, ...tempColumns);
              headerRow = row;
            }
          }

          console.log(
            "📊 executeCreatePivotTable: Available headers:",
            headers
          );

          // Find column indices
          let groupByIndex = -1;
          let valueColumnIndex = -1;

          if (/^[A-Z]+$/.test(groupBy)) {
            groupByIndex = groupBy.charCodeAt(0) - 65;
          } else {
            groupByIndex = headers.findIndex((h) =>
              h.toLowerCase().includes(groupBy.toLowerCase())
            );
          }

          if (/^[A-Z]+$/.test(valueColumn)) {
            valueColumnIndex = valueColumn.charCodeAt(0) - 65;
          } else {
            valueColumnIndex = headers.findIndex((h) =>
              h.toLowerCase().includes(valueColumn.toLowerCase())
            );
          }

          if (groupByIndex === -1 || valueColumnIndex === -1) {
            throw new Error(
              `Column not found. Available: ${headers.join(", ")}`
            );
          }

          // Create pivot from data
          const pivotMap = new Map<string, number[]>();
          const dataStartRow = headerRow + 1;
          for (let row = dataStartRow; row <= scopedEndRow; row++) {
            const rowData = cellData[row];
            if (rowData) {
              const groupValue = rowData[groupByIndex]?.v;
              const numericValue = parseFloat(
                String(rowData[valueColumnIndex]?.v)
              );

              if (groupValue && !isNaN(numericValue)) {
                const groupKey = String(groupValue);
                if (!pivotMap.has(groupKey)) {
                  pivotMap.set(groupKey, []);
                }
                pivotMap.get(groupKey)!.push(numericValue);
              }
            }
          }

          // Calculate aggregations
          const pivotData: Array<{ group: string; value: number }> = [];
          for (const [group, values] of pivotMap.entries()) {
            let aggregatedValue: number;
            switch (aggFunc) {
              case "sum":
                aggregatedValue = values.reduce((sum, val) => sum + val, 0);
                break;
              case "average":
                aggregatedValue =
                  values.reduce((sum, val) => sum + val, 0) / values.length;
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
            pivotData.push({ group, value: aggregatedValue });
          }

          pivotData.sort((a, b) => a.group.localeCompare(b.group));

          // Write pivot table to destination
          let destCell = destination || "H2";
          if (destCell.includes("!")) destCell = destCell.split("!")[1];
          const [destCol, destRow] = [
            destCell.charCodeAt(0) - 65,
            parseInt(destCell.slice(1)) - 1,
          ];

          // Write headers
          const headerRange1 = worksheet.getRange(destRow, destCol, 1, 1);
          headerRange1.setValue(headers[groupByIndex]);

          const headerRange2 = worksheet.getRange(destRow, destCol + 1, 1, 1);
          headerRange2.setValue(
            `${aggFunc.toUpperCase()} of ${headers[valueColumnIndex]}`
          );

          // Write data
          for (let i = 0; i < pivotData.length; i++) {
            const dataRow = destRow + 1 + i;
            const groupRange = worksheet.getRange(dataRow, destCol, 1, 1);
            groupRange.setValue(pivotData[i].group);

            const valueRange = worksheet.getRange(dataRow, destCol + 1, 1, 1);
            valueRange.setValue(pivotData[i].value);
          }

          console.log(
            `✅ executeCreatePivotTable: Created pivot table at ${destCell}`
          );

          return {
            success: true,
            pivotData,
            groupBy: headers[groupByIndex],
            valueColumn: headers[valueColumnIndex],
            aggFunc,
            destination: destCell,
            message: `Created pivot table grouping by '${headers[groupByIndex]}' with ${aggFunc} of '${headers[valueColumnIndex]}'. Found ${pivotData.length} groups.`,
          };
        } catch (error) {
          console.error("❌ executeCreatePivotTable: Error:", error);
          throw error;
        }
      };

      const executeGenerateChart = async (params: any) => {
        console.log("🔍 executeGenerateChart: Starting...", params);

        if (!univerAPI) {
          throw new Error("Univer API not available");
        }

        const {
          data_range,
          tableId,
          chart_type,
          title,
          position,
          width = 400,
          height = 300,
        } = params;

        try {
          const workbook = univerAPI.getActiveWorkbook();
          if (!workbook) {
            throw new Error("No active workbook available");
          }
          const worksheet = workbook.getActiveSheet();

          // Get worksheet snapshot
          const sheetSnapshot = worksheet.getSheet().getSnapshot();
          if (!sheetSnapshot || !sheetSnapshot.cellData) {
            throw new Error("No data found in spreadsheet");
          }

          // Auto-detect data range if not provided; honor tableId if present
          let finalDataRange =
            data_range || (tableId && String(tableId).split(":")[1]);
          if (!finalDataRange) {
            // Find headers first
            const cellData = sheetSnapshot.cellData;
            const headers: string[] = [];
            let headerRow = -1;

            // Find the row with the most consecutive text values (likely headers)
            for (
              let row = 0;
              row < Math.min(5, sheetSnapshot.rowCount || 5);
              row++
            ) {
              const rowData = cellData[row] || {};
              const tempColumns: string[] = [];

              for (
                let col = 0;
                col < (sheetSnapshot.columnCount || 26);
                col++
              ) {
                const cell = rowData[col];
                if (cell && typeof cell.v === "string" && !cell.f) {
                  tempColumns.push(String(cell.v));
                } else if (tempColumns.length > 0) {
                  break;
                }
              }

              if (
                tempColumns.length > headers.length &&
                tempColumns.length >= 2
              ) {
                headers.splice(0, headers.length, ...tempColumns);
                headerRow = row;
              }
            }

            // Count data rows
            let dataRowCount = 0;
            const dataStartRow = headerRow + 1;
            for (
              let row = dataStartRow;
              row < (sheetSnapshot.rowCount || 1000);
              row++
            ) {
              const rowData = cellData[row];
              if (rowData) {
                let hasData = false;
                for (let col = 0; col < headers.length; col++) {
                  const cell = rowData[col];
                  if (
                    cell &&
                    cell.v !== undefined &&
                    cell.v !== null &&
                    cell.v !== ""
                  ) {
                    hasData = true;
                    break;
                  }
                }
                if (hasData) dataRowCount++;
              }
            }

            // Auto-detect range: from A1 to last column with data, including all data rows
            const lastCol = String.fromCharCode(65 + headers.length - 1);
            finalDataRange = `A${headerRow + 1}:${lastCol}${
              dataStartRow + dataRowCount
            }`;
            console.log(
              `📊 executeGenerateChart: Auto-detected data range: ${finalDataRange}`
            );
          }

          // Determine chart position
          const chartPosition = position || "H2";

          // Map chart types to Univer enum values
          const univerChartTypeMap: { [key: string]: string } = {
            column: "Column",
            line: "Line",
            pie: "Pie",
            bar: "Bar",
            scatter: "Scatter",
          };
          const univerChartType = univerChartTypeMap[chart_type] || "Column";

          console.log(
            `📊 executeGenerateChart: Creating ${univerChartType} chart with range ${finalDataRange}`
          );

          // Insert native Univer chart via facade API
          try {
            const fWorkbook: any = univerAPI.getActiveWorkbook();
            if (!fWorkbook) throw new Error("No active workbook available");
            const fWorksheet: any = fWorkbook.getActiveSheet();
            const enumType = (univerAPI as any).Enum?.ChartType || {};
            const enumMap: Record<string, any> = {
              column: enumType.Column,
              line: enumType.Line,
              pie: enumType.Pie,
              bar: enumType.Bar,
              scatter: enumType.Scatter,
            };
            const chartTypeEnum = enumMap[chart_type] ?? enumType.Column;

            // Convert position like "H2" to row/col anchors (0-based)
            const posCol = (
              chartPosition.match(/[A-Z]+/i)?.[0] || "H"
            ).toUpperCase();
            const posRowNum =
              parseInt(chartPosition.replace(/\D+/g, ""), 10) || 2;
            const anchorRow = posRowNum - 1;
            const anchorCol = posCol.charCodeAt(0) - 65;

            const builder = fWorksheet
              .newChart()
              .setChartType(chartTypeEnum)
              .addRange(finalDataRange)
              .setPosition(anchorRow, anchorCol, 0, 0)
              .setWidth(width)
              .setHeight(height);

            if (title) {
              builder.setOptions("title.text", title);
            }

            const chartInfo = builder.build();
            await fWorksheet.insertChart(chartInfo);
          } catch (chartError) {
            console.warn(
              "⚠️ Native chart insertion failed, falling back.",
              chartError
            );
            await createImprovedChart(
              worksheet,
              finalDataRange,
              chart_type,
              title,
              chartPosition
            );
          }

          console.log(
            `✅ executeGenerateChart: Chart created at ${chartPosition}`
          );

          console.log(
            `✅ executeGenerateChart: Chart placeholder created at ${chartPosition}`
          );

          return {
            success: true,
            chartType: univerChartType,
            dataRange: finalDataRange,
            position: chartPosition,
            title,
            width,
            height,
            message: `Created ${chart_type} chart "${title}" using data range ${finalDataRange}. Chart placeholder placed at ${chartPosition}.`,
          };
        } catch (error) {
          console.error("❌ executeGenerateChart: Error:", error);
          throw error;
        }
      };

      const executeFormatCurrency = async (params: any) => {
        console.log("💰 executeFormatCurrency: Starting...", params);

        if (!univerAPI) {
          throw new Error("Univer API not available");
        }

        let { range, currency, decimals = 2 } = params;

        // If no range provided, try to use the last sum cell
        if (!range && window.lastSumCell) {
          range = window.lastSumCell;
          console.log(`💰 executeFormatCurrency: Using last sum cell ${range}`);
        }

        if (!range) {
          throw new Error("No range provided and no recent sum cell available");
        }

        try {
          const fWorkbook = univerAPI.getActiveWorkbook();
          if (!fWorkbook) {
            throw new Error("No active workbook available");
          }
          const fWorksheet = fWorkbook.getActiveSheet();
          const fRange = fWorksheet.getRange(range);

          // Currency format patterns
          const currencyFormats: { [key: string]: string } = {
            USD: `$#,##0.${"0".repeat(decimals)}`,
            EUR: `€#,##0.${"0".repeat(decimals)}`,
            GBP: `£#,##0.${"0".repeat(decimals)}`,
            JPY: `¥#,##0`,
            CAD: `C$#,##0.${"0".repeat(decimals)}`,
            AUD: `A$#,##0.${"0".repeat(decimals)}`,
          };

          const formatPattern =
            currencyFormats[currency.toUpperCase()] ||
            `${currency}#,##0.${"0".repeat(decimals)}`;

          // Apply currency formatting using Univer's proper API
          fRange.setNumberFormat(formatPattern);

          console.log(
            `✅ executeFormatCurrency: Applied ${currency} formatting to ${range}`
          );

          return {
            success: true,
            range,
            currency,
            decimals,
            formatPattern,
            message: `Applied ${currency} currency formatting to ${range}`,
          };
        } catch (error) {
          console.error("❌ executeFormatCurrency: Error:", error);
          throw error;
        }
      };

      const executeSwitchSheet = async (params: any) => {
        console.log("🔍 executeSwitchSheet: Starting...", params);

        // TODO: Implement proper sheet switching with correct Univer API methods
        // Current implementation uses methods that don't exist on FWorksheet
        return {
          success: false,
          message:
            "Sheet switching functionality requires proper Univer API implementation",
          error: "Not implemented with correct API methods",
        };
      };

      const executeAddFilter = async (params: any) => {
        console.log("🔍 executeAddFilter: Starting...", params);

        if (!univerAPI) {
          throw new Error("Univer API not available");
        }

        const { range, column, filterValues, action = "add" } = params;

        try {
          const workbook = univerAPI.getActiveWorkbook();
          if (!workbook) {
            throw new Error("No active workbook available");
          }

          const worksheet = workbook.getActiveSheet();
          if (!worksheet) {
            throw new Error("No active worksheet available");
          }

          // Handle different actions
          if (action === "clear") {
            try {
              // Clear all filters on the worksheet
              const existingFilter = (worksheet as any).getFilter();
              if (existingFilter) {
                (existingFilter as any).remove();
                return {
                  success: true,
                  message: "All filters cleared successfully",
                };
              } else {
                return {
                  success: true,
                  message: "No filters to clear",
                };
              }
            } catch (error) {
              return {
                success: false,
                message: `Failed to clear filters: ${
                  error instanceof Error ? error.message : "Unknown error"
                }`,
              };
            }
          }

          // Auto-detect data range if not provided
          let filterRange = range;
          if (!filterRange) {
            const sheetSnapshot = worksheet.getSheet().getSnapshot();
            const cellData = sheetSnapshot.cellData || {};

            // Find headers and data extent
            let maxRow = 0;
            let maxCol = 0;
            for (const rowIndex in cellData) {
              const row = parseInt(rowIndex);
              maxRow = Math.max(maxRow, row);
              for (const colIndex in cellData[rowIndex]) {
                const col = parseInt(colIndex);
                maxCol = Math.max(maxCol, col);
              }
            }

            if (maxRow > 0 && maxCol > 0) {
              filterRange = `A1:${String.fromCharCode(65 + maxCol)}${
                maxRow + 1
              }`;
            } else {
              filterRange = "A1:G51"; // Default range
            }
          }

          console.log(`🔍 executeAddFilter: Using range ${filterRange}`);

          // Parse the range to get the actual range object
          const [startCell, endCell] = filterRange.split(":");
          const startCol = startCell.charCodeAt(0) - 65;
          const startRow = parseInt(startCell.slice(1)) - 1;
          const endCol = endCell.charCodeAt(0) - 65;
          const endRow = parseInt(endCell.slice(1)) - 1;

          // Get the range object
          const fRange = worksheet.getRange(
            startRow,
            startCol,
            endRow - startRow + 1,
            endCol - startCol + 1
          );

          if (action === "add") {
            try {
              // Try to use Univer's facade API for filters
              console.log(
                "🔍 executeAddFilter: Attempting to use Univer facade API"
              );

              // Get the facade API
              const facade = (univerAPI as any).facade;
              if (facade && facade.FWorkbook) {
                const fWorkbook = facade.FWorkbook.getActiveWorkbook(univerAPI);
                if (fWorkbook) {
                  const fWorksheet = fWorkbook.getActiveSheet();
                  if (fWorksheet) {
                    console.log("🔍 executeAddFilter: Using facade FWorksheet");

                    // Try to use the correct filter API from the facade
                    const fRange2 = fWorksheet.getRange(filterRange);
                    console.log(
                      "🔍 executeAddFilter: Got facade range:",
                      fRange2
                    );

                    // Try different filter method approaches
                    if (typeof (fRange2 as any).createFilter === "function") {
                      const filter = (fRange2 as any).createFilter();
                      console.log(
                        "🔍 executeAddFilter: Created filter via facade range"
                      );

                      if (column && filterValues && filterValues.length > 0) {
                        // Find column index
                        let columnIndex = -1;
                        if (/^[A-Z]+$/.test(column.toUpperCase())) {
                          columnIndex = column.toUpperCase().charCodeAt(0) - 65;
                        } else {
                          // Find by column name
                          const sheetSnapshot = worksheet
                            .getSheet()
                            .getSnapshot();
                          const cellData = sheetSnapshot.cellData || {};
                          const headerRow = cellData[startRow] || {};

                          for (let col = startCol; col <= endCol; col++) {
                            const cell = headerRow[col];
                            if (
                              cell &&
                              cell.v &&
                              String(cell.v)
                                .toLowerCase()
                                .includes(column.toLowerCase())
                            ) {
                              columnIndex = col;
                              break;
                            }
                          }
                        }

                        if (
                          columnIndex >= 0 &&
                          filter &&
                          typeof (filter as any).setColumnFilterCriteria ===
                            "function"
                        ) {
                          // Apply filter criteria
                          (filter as any).setColumnFilterCriteria(columnIndex, {
                            colId: columnIndex,
                            filters: { filters: filterValues },
                          });

                          return {
                            success: true,
                            message: `Filter applied to column ${column}. Showing only: ${filterValues.join(
                              ", "
                            )}`,
                            range: filterRange,
                            column: column,
                            columnIndex: columnIndex,
                            filterValues: filterValues,
                          };
                        }
                      }

                      return {
                        success: true,
                        message: `Auto filter created for range ${filterRange}. Click dropdown arrows in headers to filter data.`,
                        range: filterRange,
                      };
                    }
                  }
                }
              }

              // Fallback: Basic filter simulation by hiding rows
              console.log("🔍 executeAddFilter: Using row hiding fallback");
              if (column && filterValues && filterValues.length > 0) {
                // Find column index
                let columnIndex = -1;
                if (/^[A-Z]+$/.test(column.toUpperCase())) {
                  columnIndex = column.toUpperCase().charCodeAt(0) - 65;
                } else {
                  // Find by column name
                  const sheetSnapshot = worksheet.getSheet().getSnapshot();
                  const cellData = sheetSnapshot.cellData || {};
                  const headerRow = cellData[startRow] || {};

                  for (let col = startCol; col <= endCol; col++) {
                    const cell = headerRow[col];
                    if (
                      cell &&
                      cell.v &&
                      String(cell.v)
                        .toLowerCase()
                        .includes(column.toLowerCase())
                    ) {
                      columnIndex = col;
                      break;
                    }
                  }
                }

                if (columnIndex >= 0) {
                  // Simple filter simulation: create a message showing what would be filtered
                  const sheetSnapshot = worksheet.getSheet().getSnapshot();
                  const cellData = sheetSnapshot.cellData || {};

                  let matchingRows = 0;
                  for (let row = startRow + 1; row <= endRow; row++) {
                    const rowData = cellData[row];
                    if (rowData && rowData[columnIndex]) {
                      const cellValue = String(
                        rowData[columnIndex].v || ""
                      ).toLowerCase();
                      const matches = filterValues.some(
                        (filterVal: string) =>
                          cellValue.includes(filterVal.toLowerCase()) ||
                          filterVal.toLowerCase().includes(cellValue)
                      );
                      if (matches) {
                        matchingRows++;
                      }
                    }
                  }

                  return {
                    success: true,
                    message: `Filter simulation: Found ${matchingRows} rows matching '${filterValues.join(
                      ", "
                    )}' in column ${column}. Note: Visual filtering requires filter UI components.`,
                    range: filterRange,
                    column: column,
                    columnIndex: columnIndex,
                    filterValues: filterValues,
                    matchingRows: matchingRows,
                  };
                }
              }

              throw new Error(
                "Unable to apply filter - filter API not available"
              );
            } catch (filterError) {
              console.error("Filter creation failed:", filterError);
              return {
                success: false,
                message: `Filter operation failed: ${
                  filterError instanceof Error
                    ? filterError.message
                    : "Unknown error"
                }. The filter functionality may require additional Univer components.`,
                range: filterRange,
              };
            }
          } else if (action === "remove" && column) {
            try {
              const fFilter = (fRange as any).getFilter();
              if (fFilter) {
                // Find column index
                let columnIndex = -1;
                if (/^[A-Z]+$/.test(column.toUpperCase())) {
                  columnIndex = column.toUpperCase().charCodeAt(0) - 65;
                } else {
                  // Find by column name
                  const sheetSnapshot = worksheet.getSheet().getSnapshot();
                  const cellData = sheetSnapshot.cellData || {};
                  const headerRow = cellData[startRow] || {};

                  for (let col = startCol; col <= endCol; col++) {
                    const cell = headerRow[col];
                    if (
                      cell &&
                      cell.v &&
                      String(cell.v)
                        .toLowerCase()
                        .includes(column.toLowerCase())
                    ) {
                      columnIndex = col;
                      break;
                    }
                  }
                }

                if (columnIndex >= 0) {
                  (fFilter as any).removeColumnFilterCriteria(columnIndex);
                  return {
                    success: true,
                    message: `Filter removed from column ${column}`,
                    column: column,
                    columnIndex: columnIndex,
                  };
                } else {
                  throw new Error(`Column '${column}' not found`);
                }
              } else {
                return {
                  success: true,
                  message: "No filters to remove",
                };
              }
            } catch (error) {
              return {
                success: false,
                message: `Failed to remove filter: ${
                  error instanceof Error ? error.message : "Unknown error"
                }`,
              };
            }
          }

          return {
            success: false,
            message: `Unknown filter action: ${action}`,
          };
        } catch (error) {
          console.error("❌ executeAddFilter: Error:", error);
          throw error;
        }
      };

      // Function to create improved chart display with working data
      const createImprovedChart = async (
        worksheet: any,
        dataRange: string,
        chartType: string,
        title: string,
        position: string
      ) => {
        try {
          console.log(
            `📊 createImprovedChart: Creating ${chartType} chart with data range ${dataRange}`
          );

          // Parse position to get row and column coordinates
          const [posCol, posRow] = [
            position.charCodeAt(0) - 65,
            parseInt(position.slice(1)) - 1,
          ];

          // Extract data from the range
          const chartData = await extractChartData(worksheet, dataRange);

          if (!chartData || chartData.length === 0) {
            console.log("❌ createImprovedChart: No data found in range");
            const errorCell = worksheet.getRange(posRow, posCol, 1, 1);
            errorCell.setValue(`❌ No data found in ${dataRange}`);
            return;
          }

          // Generate chart URL with the improved data detection
          const chartUrl = await generateImprovedChartUrl(
            chartData,
            chartType,
            title
          );

          // Create a compact chart display
          const titleCell = worksheet.getRange(posRow, posCol, 1, 1);
          titleCell.setValue(`📊 ${title}`);

          const urlCell = worksheet.getRange(posRow + 1, posCol, 1, 1);
          urlCell.setValue(chartUrl);

          const infoCell = worksheet.getRange(posRow + 2, posCol, 1, 1);
          infoCell.setValue(
            `${chartType.toUpperCase()}: ${chartData.length - 1} data points`
          );

          console.log(`✅ createImprovedChart: Chart created: ${chartUrl}`);
        } catch (error) {
          console.error("❌ createImprovedChart: Error creating chart:", error);

          const [posCol, posRow] = [
            position.charCodeAt(0) - 65,
            parseInt(position.slice(1)) - 1,
          ];
          const errorCell = worksheet.getRange(posRow, posCol, 1, 1);
          errorCell.setValue(`❌ Chart error: ${String(error).slice(0, 50)}`);
        }
      };

      // Function to extract chart data from range
      const extractChartData = async (worksheet: any, range: string) => {
        try {
          // Parse range like "A1:G51"
          const [startCell, endCell] = range.split(":");

          // Parse start cell (e.g., "A1")
          const startCol = startCell.charCodeAt(0) - 65;
          const startRow = parseInt(startCell.slice(1)) - 1;

          // Parse end cell (e.g., "G51")
          const endCol = endCell.charCodeAt(0) - 65;
          const endRow = parseInt(endCell.slice(1)) - 1;

          // Get sheet snapshot to access data
          const sheetSnapshot = worksheet.getSheet().getSnapshot();
          const cellData = sheetSnapshot.cellData;

          const extractedData = [];

          // Extract data row by row
          for (
            let row = startRow;
            row <= Math.min(endRow, startRow + 100);
            row++
          ) {
            // Limit to 100 rows for charts
            const rowData = cellData[row];
            if (rowData) {
              const rowValues = [];
              for (let col = startCol; col <= endCol; col++) {
                const cell = rowData[col];
                rowValues.push(cell ? cell.v : "");
              }
              // Only add rows that have some data
              if (rowValues.some((val) => val !== "" && val != null)) {
                extractedData.push(rowValues);
              }
            }
          }

          console.log(
            `📊 extractChartData: Extracted ${extractedData.length} rows of data`
          );
          return extractedData;
        } catch (error) {
          console.error("❌ extractChartData: Error extracting data:", error);
          return [];
        }
      };

      // Function to generate improved chart URL with better data detection
      const generateImprovedChartUrl = async (
        data: any[],
        chartType: string,
        title: string
      ) => {
        try {
          if (!data || data.length < 2) return null;

          const headers = data[0]; // First row is headers
          const dataRows = data.slice(1); // Skip headers row

          console.log(
            `📊 generateImprovedChartUrl: Processing ${dataRows.length} data rows with headers:`,
            headers
          );

          // Validate we have meaningful data
          if (dataRows.length === 0) {
            console.error("❌ generateImprovedChartUrl: No data rows found");
            return `https://quickchart.io/chart?c=${encodeURIComponent(
              JSON.stringify({
                type: "bar",
                data: {
                  labels: ["No Data"],
                  datasets: [{ label: "Error", data: [0] }],
                },
              })
            )}`;
          }

          // Dynamic column detection strategy
          let labelColumnIndex = -1;
          let valueColumnIndex = -1;

          // Strategy 1: Look for common text/category column names for labels
          const textColumnNames = [
            "product",
            "category",
            "region",
            "salesperson",
            "name",
            "item",
            "type",
          ];
          for (let i = 0; i < headers.length; i++) {
            const header = String(headers[i]).toLowerCase();
            if (textColumnNames.some((name) => header.includes(name))) {
              labelColumnIndex = i;
              break;
            }
          }

          // Strategy 2: Look for numeric columns for values (prefer rightmost numeric columns)
          const numericColumnNames = [
            "total",
            "sale",
            "amount",
            "value",
            "price",
            "revenue",
            "cost",
          ];
          for (let i = headers.length - 1; i >= 0; i--) {
            const header = String(headers[i]).toLowerCase();
            if (numericColumnNames.some((name) => header.includes(name))) {
              valueColumnIndex = i;
              break;
            }
          }

          // Fallback: If no semantic matches, use first column for labels, last for values
          if (labelColumnIndex === -1) {
            labelColumnIndex = 0;
          }
          if (valueColumnIndex === -1) {
            valueColumnIndex = Math.max(1, headers.length - 1);
          }

          console.log(
            `📊 Using column ${labelColumnIndex} (${headers[labelColumnIndex]}) for labels, column ${valueColumnIndex} (${headers[valueColumnIndex]}) for values`
          );

          // Extract unique labels and aggregate values
          const labelValueMap = new Map<string, number>();

          for (const row of dataRows) {
            const label = String(row[labelColumnIndex] || "Unknown");
            const value = parseFloat(String(row[valueColumnIndex])) || 0;

            if (labelValueMap.has(label)) {
              labelValueMap.set(label, labelValueMap.get(label)! + value);
            } else {
              labelValueMap.set(label, value);
            }
          }

          const labels = Array.from(labelValueMap.keys());
          const salesData = Array.from(labelValueMap.values());

          console.log(`📊 Chart data prepared: ${labels.length} categories`, {
            labels,
            values: salesData,
          });

          const chartTypeMap: { [key: string]: string } = {
            column: "bar",
            bar: "horizontalBar",
            line: "line",
            pie: "pie",
            scatter: "scatter",
          };

          const config = {
            type: chartTypeMap[chartType] || "bar",
            data: {
              labels: labels,
              datasets: [
                {
                  label: headers[valueColumnIndex] || "Values",
                  data: salesData,
                  backgroundColor: [
                    "rgba(255, 99, 132, 0.8)",
                    "rgba(54, 162, 235, 0.8)",
                    "rgba(255, 205, 86, 0.8)",
                    "rgba(75, 192, 192, 0.8)",
                    "rgba(153, 102, 255, 0.8)",
                    "rgba(255, 159, 64, 0.8)",
                  ],
                },
              ],
            },
            options: {
              responsive: true,
              plugins: {
                title: {
                  display: true,
                  text: title,
                },
              },
            },
          };

          // Generate QuickChart URL
          const chartConfigString = JSON.stringify(config);
          const encodedConfig = encodeURIComponent(chartConfigString);

          return `https://quickchart.io/chart?w=600&h=400&c=${encodedConfig}`;
        } catch (error) {
          console.error("❌ generateImprovedChartUrl: Error:", error);
          return `https://quickchart.io/chart?c=${encodeURIComponent(
            JSON.stringify({
              type: "bar",
              data: { labels: ["Error"], datasets: [{ data: [0] }] },
            })
          )}`;
        }
      };

      console.log("🎯 Univer component: Initialization complete");
    };

    initializeUniver().catch(console.error);
  }, []);

  return <div ref={containerRef} className="h-full" />;
}
