// services/pivotTableService.ts
import { UniverService } from "./univerService";

interface PivotTableOptions {
  groupBy: string;
  valueColumn: string;
  aggFunc: "sum" | "average" | "count" | "max" | "min";
  filter?: Record<string, any>;
  destination?: string;
}

interface PivotTableResult {
  data: any[];
  summary: string;
  destination: string;
}

export class PivotTableService {
  private static instance: PivotTableService;
  private univerService: UniverService;

  private constructor() {
    this.univerService = UniverService.getInstance();
  }

  static getInstance(): PivotTableService {
    if (!PivotTableService.instance) {
      PivotTableService.instance = new PivotTableService();
    }
    return PivotTableService.instance;
  }

  async createPivotTable(
    options: PivotTableOptions
  ): Promise<PivotTableResult> {
    try {
      const { groupBy, valueColumn, aggFunc, filter, destination } = options;

      // Get sheet context to understand the data structure
      const context = await this.univerService.getSheetContext();

      if (!context.tables || context.tables.length === 0) {
        throw new Error("No table data found in the sheet");
      }

      const table = context.tables[0];
      const headers = table.headers;
      const data = table.dataPreview;

      // Find column indices
      const groupByIndex = headers.findIndex((h) =>
        h.toLowerCase().includes(groupBy.toLowerCase())
      );
      const valueColumnIndex = headers.findIndex((h) =>
        h.toLowerCase().includes(valueColumn.toLowerCase())
      );

      if (groupByIndex === -1) {
        throw new Error(
          `Column '${groupBy}' not found. Available columns: ${headers.join(
            ", "
          )}`
        );
      }

      if (valueColumnIndex === -1) {
        throw new Error(
          `Column '${valueColumn}' not found. Available columns: ${headers.join(
            ", "
          )}`
        );
      }

      // Create pivot table data
      const pivotData = this.createPivotData(
        data,
        groupByIndex,
        valueColumnIndex,
        aggFunc,
        filter
      );

      // Determine destination
      const destRange = destination || this.findNextAvailableRange(context);

      // Write pivot table to sheet
      await this.writePivotTableToSheet(
        pivotData,
        destRange,
        headers[groupByIndex],
        headers[valueColumnIndex],
        aggFunc
      );

      return {
        data: pivotData,
        summary: `Created pivot table grouped by ${groupBy} with ${aggFunc} of ${valueColumn}`,
        destination: destRange,
      };
    } catch (error) {
      throw new Error(
        `Failed to create pivot table: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  private createPivotData(
    data: string[][],
    groupByIndex: number,
    valueColumnIndex: number,
    aggFunc: string,
    filter?: Record<string, any>
  ): any[] {
    const groups = new Map<string, number[]>();

    // Group data
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const groupKey = row[groupByIndex] || "Unknown";
      const value = parseFloat(row[valueColumnIndex]) || 0;

      // Apply filter if specified
      if (filter) {
        let shouldInclude = true;
        for (const [key, filterValue] of Object.entries(filter)) {
          const columnIndex = data[0]?.findIndex((h) =>
            h.toLowerCase().includes(key.toLowerCase())
          );
          if (columnIndex !== -1 && row[columnIndex] !== filterValue) {
            shouldInclude = false;
            break;
          }
        }
        if (!shouldInclude) continue;
      }

      if (!groups.has(groupKey)) {
        groups.set(groupKey, []);
      }
      groups.get(groupKey)!.push(value);
    }

    // Apply aggregation function
    const pivotData: any[] = [];
    for (const [group, values] of groups) {
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

      pivotData.push({
        group,
        value: aggregatedValue,
        count: values.length,
      });
    }

    return pivotData;
  }

  private findNextAvailableRange(context: any): string {
    // Find next available range starting from column K
    return "K1";
  }

  private async writePivotTableToSheet(
    pivotData: any[],
    destRange: string,
    groupByHeader: string,
    valueColumnHeader: string,
    aggFunc: string
  ): Promise<void> {
    // Write headers
    const headerRow = parseInt(destRange.match(/\d+/)?.[0] || "1");
    const headerCol = destRange.match(/[A-Z]+/)?.[0] || "K";

    await this.univerService.setCellValue(
      `${headerCol}${headerRow}`,
      groupByHeader
    );
    await this.univerService.setCellValue(
      `${this.getNextColumn(headerCol)}${headerRow}`,
      `${aggFunc}(${valueColumnHeader})`
    );

    // Write data
    for (let i = 0; i < pivotData.length; i++) {
      const row = headerRow + 1 + i;
      await this.univerService.setCellValue(
        `${headerCol}${row}`,
        pivotData[i].group
      );
      await this.univerService.setCellValue(
        `${this.getNextColumn(headerCol)}${row}`,
        pivotData[i].value
      );
    }
  }

  private getNextColumn(col: string): string {
    let result = "";
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - 64);
    }
    num++;
    while (num > 0) {
      num--;
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26);
    }
    return result;
  }
}
