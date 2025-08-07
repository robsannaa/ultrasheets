# 04. Pivot Table Service (HIGH)

## What?

Implement the `create_pivot_table` tool that creates pivot tables with grouping, aggregation, and filtering capabilities, as specified in Section 3.2 of the integration spec.

## Where?

- **New file**: `services/pivotTableService.ts`
- **Dependencies**: UniverService, ComprehensiveSpreadsheetService
- **Integration**: `app/api/chat/route.ts` (add to tools)

## How?

### 4.1 Create PivotTableService

```typescript
// services/pivotTableService.ts
import { UniverService } from "./univerService";
import { ComprehensiveSpreadsheetService } from "./comprehensiveSpreadsheetService";

interface PivotTableConfig {
  groupBy: string;
  valueColumn: string;
  aggFunc: "sum" | "average" | "count" | "max" | "min";
  filter?: Record<string, any>;
  destination?: string;
}

interface PivotResult {
  data: any[];
  summary: {
    totalGroups: number;
    totalValue: number;
    groupByColumn: string;
    valueColumn: string;
    aggregationFunction: string;
  };
  destination: string;
}

export class PivotTableService {
  private static instance: PivotTableService;
  private univerService: UniverService;
  private contextService: ComprehensiveSpreadsheetService;

  constructor() {
    this.univerService = UniverService.getInstance();
    this.contextService = new ComprehensiveSpreadsheetService();
  }

  static getInstance(): PivotTableService {
    if (!PivotTableService.instance) {
      PivotTableService.instance = new PivotTableService();
    }
    return PivotTableService.instance;
  }

  async createPivotTable(config: PivotTableConfig): Promise<PivotResult> {
    // 1. Get sheet context and data
    const context = await this.univerService.getSheetContext();
    const data = await this.extractDataFromSheet(context);

    // 2. Apply filters if specified
    const filteredData = config.filter
      ? this.applyFilters(data, config.filter)
      : data;

    // 3. Group and aggregate data
    const groupedData = this.groupAndAggregate(
      filteredData,
      config.groupBy,
      config.valueColumn,
      config.aggFunc
    );

    // 4. Write pivot table to sheet
    const destination =
      config.destination || (await this.findPivotDestination());
    await this.writePivotTableToSheet(groupedData, destination, config);

    // 5. Return result
    return {
      data: groupedData,
      summary: {
        totalGroups: groupedData.length,
        totalValue: this.calculateTotalValue(groupedData, config.aggFunc),
        groupByColumn: config.groupBy,
        valueColumn: config.valueColumn,
        aggregationFunction: config.aggFunc,
      },
      destination,
    };
  }

  private async extractDataFromSheet(context: any): Promise<any[]> {
    // Extract data from Univer sheet context
    // Convert to array of objects format
    const table = context.tables[0];
    if (!table) throw new Error("No table data found");

    return table.dataPreview.map((row: string[], index: number) => {
      const obj: any = {};
      table.headers.forEach((header: string, colIndex: number) => {
        obj[header] = row[colIndex] || "";
      });
      return obj;
    });
  }

  private applyFilters(data: any[], filters: Record<string, any>): any[] {
    return data.filter((row) => {
      return Object.entries(filters).every(([key, value]) => {
        return row[key] === value;
      });
    });
  }

  private groupAndAggregate(
    data: any[],
    groupBy: string,
    valueColumn: string,
    aggFunc: string
  ): any[] {
    const groups = new Map<string, any[]>();

    // Group data
    data.forEach((row) => {
      const groupKey = row[groupBy];
      if (!groups.has(groupKey)) {
        groups.set(groupKey, []);
      }
      groups.get(groupKey)!.push(row);
    });

    // Aggregate each group
    return Array.from(groups.entries()).map(([groupKey, groupData]) => {
      const aggregatedValue = this.aggregate(groupData, valueColumn, aggFunc);
      return {
        [groupBy]: groupKey,
        [`${valueColumn}_${aggFunc}`]: aggregatedValue,
        count: groupData.length,
      };
    });
  }

  private aggregate(data: any[], column: string, func: string): number {
    const values = data.map((row) => parseFloat(row[column]) || 0);

    switch (func) {
      case "sum":
        return values.reduce((a, b) => a + b, 0);
      case "average":
        return values.reduce((a, b) => a + b, 0) / values.length;
      case "count":
        return values.length;
      case "max":
        return Math.max(...values);
      case "min":
        return Math.min(...values);
      default:
        return 0;
    }
  }

  private async writePivotTableToSheet(
    data: any[],
    destination: string,
    config: PivotTableConfig
  ): Promise<void> {
    // Write headers
    const headers = Object.keys(data[0] || {});
    const headerRow = parseInt(destination.match(/\d+/)?.[0] || "1");
    const startCol = destination.match(/[A-Z]+/)?.[0] || "A";

    // Write headers
    for (let i = 0; i < headers.length; i++) {
      const cell = `${String.fromCharCode(
        startCol.charCodeAt(0) + i
      )}${headerRow}`;
      await this.univerService.setCellValue(cell, headers[i]);
    }

    // Write data
    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      for (let colIndex = 0; colIndex < headers.length; colIndex++) {
        const cell = `${String.fromCharCode(
          startCol.charCodeAt(0) + colIndex
        )}${headerRow + rowIndex + 1}`;
        await this.univerService.setCellValue(cell, row[headers[colIndex]]);
      }
    }
  }

  private async findPivotDestination(): Promise<string> {
    // Find next available area for pivot table
    // Start from K1, move right/down as needed
    return "K1";
  }

  private calculateTotalValue(data: any[], aggFunc: string): number {
    const values = data.map((row) => {
      const key = Object.keys(row).find(
        (k) => k.includes("_sum") || k.includes("_average")
      );
      return key ? parseFloat(row[key]) || 0 : 0;
    });

    return values.reduce((a, b) => a + b, 0);
  }
}
```

### 4.2 Update create_pivot_table Tool

```typescript
// In app/api/chat/route.ts
create_pivot_table: tool({
  description: "Create a pivot table with grouping and aggregation",
  parameters: z.object({
    groupBy: z.string().describe("Column to group by (e.g., 'Region')"),
    valueColumn: z.string().describe("Column to aggregate (e.g., 'Sales')"),
    aggFunc: z
      .enum(["sum", "average", "count", "max", "min"])
      .describe("Aggregation function"),
    filter: z.record(z.any()).optional().describe("Optional filter criteria"),
    destination: z
      .string()
      .optional()
      .describe("Destination range (e.g., 'K1')"),
  }),
  execute: async ({ groupBy, valueColumn, aggFunc, filter, destination }) => {
    const pivotService = PivotTableService.getInstance();
    const result = await pivotService.createPivotTable({
      groupBy,
      valueColumn,
      aggFunc,
      filter,
      destination,
    });

    return {
      message: `Created pivot table grouped by ${groupBy}`,
      data: result.data,
      summary: result.summary,
      location: result.destination,
    };
  },
});
```

## How to Test?

1. **Unit Tests**: `__tests__/services/pivotTableService.test.ts`

   ```typescript
   describe("PivotTableService", () => {
     it("should create pivot table with sum aggregation", async () => {
       const service = PivotTableService.getInstance();
       const result = await service.createPivotTable({
         groupBy: "Region",
         valueColumn: "Sales",
         aggFunc: "sum",
       });
       expect(result.data).toHaveLength(3); // Assuming 3 regions
       expect(result.summary.totalGroups).toBe(3);
     });
   });
   ```

2. **Integration Tests**:

   - Test with real sheet data
   - Verify pivot table appears in sheet
   - Test different aggregation functions

3. **Manual Testing**:
   ```bash
   # Test pivot table creation
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"create pivot table by region"}]}'
   ```

## Important Dependencies to Not Break

- **UniverService data extraction** - Don't break existing sheet reading
- **Cell writing operations** - Preserve existing cell update patterns
- **Sheet context** - Don't break ComprehensiveSpreadsheetService
- **Data types** - Handle mixed data types properly

## Dependencies That Will Work Thanks to This

- **Chart generation** - Pivot tables provide data for charts
- **Financial analysis** - Complex aggregations support analysis
- **Data visualization** - Structured data for charts
- **User experience** - Professional pivot table output

## Implementation Strategy (Least Disruptive)

1. **Start with data extraction** - Test reading sheet data first
2. **Implement grouping logic** - Core pivot functionality
3. **Add aggregation functions** - One function at a time
4. **Test sheet writing** - Verify pivot table appears correctly
5. **Add filtering** - Enhance with filter capabilities

## Priority Order

1. Create PivotTableService with basic structure
2. Implement data extraction from sheet
3. Add grouping and aggregation logic
4. Implement sheet writing functionality
5. Add filtering capabilities
6. Integrate with create_pivot_table tool
7. Add comprehensive error handling
