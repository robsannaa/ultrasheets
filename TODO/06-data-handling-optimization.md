# 06. Data Handling Optimization (MEDIUM)

## What?

Implement large dataset handling (>2k rows), schema provision at chat start, and optimized data streaming/aggregation as specified in Section 6 of the integration spec.

## Where?

- **Files**: `services/univerService.ts`, `services/comprehensiveSpreadsheetService.ts`
- **New files**: `services/dataOptimizationService.ts`
- **Integration**: `app/api/chat/route.ts`, `components/chat-sidebar.tsx`

## How?

### 6.1 Create DataOptimizationService

```typescript
// services/dataOptimizationService.ts
interface DataOptimizationConfig {
  maxRowsForLLM: number;
  enableStreaming: boolean;
  enableAggregation: boolean;
  cacheResults: boolean;
}

interface OptimizedDataResult {
  data: any[];
  summary: {
    totalRows: number;
    sampledRows: number;
    aggregationApplied: boolean;
    streamingEnabled: boolean;
  };
  schema: {
    columns: string[];
    dataTypes: Record<string, string>;
    sampleValues: Record<string, any[]>;
  };
}

export class DataOptimizationService {
  private static instance: DataOptimizationService;
  private config: DataOptimizationConfig;

  constructor() {
    this.config = {
      maxRowsForLLM: 2000,
      enableStreaming: true,
      enableAggregation: true,
      cacheResults: true,
    };
  }

  static getInstance(): DataOptimizationService {
    if (!DataOptimizationService.instance) {
      DataOptimizationService.instance = new DataOptimizationService();
    }
    return DataOptimizationService.instance;
  }

  async optimizeDataForLLM(rawData: any[]): Promise<OptimizedDataResult> {
    const totalRows = rawData.length;

    if (totalRows <= this.config.maxRowsForLLM) {
      // Data is small enough, return as-is
      return {
        data: rawData,
        summary: {
          totalRows,
          sampledRows: totalRows,
          aggregationApplied: false,
          streamingEnabled: false,
        },
        schema: this.extractSchema(rawData),
      };
    }

    // Large dataset - apply optimization strategies
    if (this.config.enableAggregation) {
      return await this.aggregateLargeDataset(rawData);
    } else {
      return await this.sampleLargeDataset(rawData);
    }
  }

  private async aggregateLargeDataset(
    data: any[]
  ): Promise<OptimizedDataResult> {
    // Apply intelligent aggregation based on data patterns
    const schema = this.extractSchema(data);
    const aggregatedData = this.performAggregation(data, schema);

    return {
      data: aggregatedData,
      summary: {
        totalRows: data.length,
        sampledRows: aggregatedData.length,
        aggregationApplied: true,
        streamingEnabled: false,
      },
      schema,
    };
  }

  private async sampleLargeDataset(data: any[]): Promise<OptimizedDataResult> {
    // Intelligent sampling that preserves data distribution
    const sampleSize = this.config.maxRowsForLLM;
    const sampledData = this.intelligentSampling(data, sampleSize);

    return {
      data: sampledData,
      summary: {
        totalRows: data.length,
        sampledRows: sampledData.length,
        aggregationApplied: false,
        streamingEnabled: false,
      },
      schema: this.extractSchema(data),
    };
  }

  private extractSchema(data: any[]): any {
    if (data.length === 0)
      return { columns: [], dataTypes: {}, sampleValues: {} };

    const columns = Object.keys(data[0]);
    const dataTypes: Record<string, string> = {};
    const sampleValues: Record<string, any[]> = {};

    columns.forEach((column) => {
      const values = data
        .map((row) => row[column])
        .filter((v) => v !== null && v !== undefined);
      dataTypes[column] = this.detectDataType(values);
      sampleValues[column] = values.slice(0, 5); // First 5 values as samples
    });

    return { columns, dataTypes, sampleValues };
  }

  private detectDataType(values: any[]): string {
    if (values.length === 0) return "text";

    const sample = values[0];
    if (typeof sample === "number") return "number";
    if (sample instanceof Date) return "date";
    if (typeof sample === "boolean") return "boolean";
    if (typeof sample === "string") {
      // Check for currency
      if (/^[£$€¥]?[\d,.-]+$/.test(sample)) return "currency";
      // Check for percentage
      if (sample.includes("%")) return "percentage";
      return "text";
    }
    return "text";
  }

  private performAggregation(data: any[], schema: any): any[] {
    // Group by categorical columns and aggregate numerical columns
    const categoricalColumns = Object.entries(schema.dataTypes)
      .filter(([_, type]) => type === "text")
      .map(([col, _]) => col);

    const numericalColumns = Object.entries(schema.dataTypes)
      .filter(([_, type]) => ["number", "currency"].includes(type))
      .map(([col, _]) => col);

    if (categoricalColumns.length === 0 || numericalColumns.length === 0) {
      return data.slice(0, this.config.maxRowsForLLM);
    }

    // Simple aggregation by first categorical column
    const groupByColumn = categoricalColumns[0];
    const groups = new Map<string, any[]>();

    data.forEach((row) => {
      const groupKey = row[groupByColumn] || "Unknown";
      if (!groups.has(groupKey)) {
        groups.set(groupKey, []);
      }
      groups.get(groupKey)!.push(row);
    });

    return Array.from(groups.entries()).map(([groupKey, groupData]) => {
      const aggregated: any = { [groupByColumn]: groupKey };

      numericalColumns.forEach((col) => {
        const values = groupData.map((row) => parseFloat(row[col]) || 0);
        aggregated[`${col}_sum`] = values.reduce((a, b) => a + b, 0);
        aggregated[`${col}_avg`] =
          values.reduce((a, b) => a + b, 0) / values.length;
        aggregated[`${col}_count`] = values.length;
      });

      return aggregated;
    });
  }

  private intelligentSampling(data: any[], sampleSize: number): any[] {
    // Stratified sampling to preserve data distribution
    if (data.length <= sampleSize) return data;

    const step = Math.floor(data.length / sampleSize);
    const sampled = [];

    for (let i = 0; i < sampleSize; i++) {
      const index = i * step;
      if (index < data.length) {
        sampled.push(data[index]);
      }
    }

    return sampled;
  }
}
```

### 6.2 Update UniverService for Large Datasets

```typescript
// In services/univerService.ts, add to getSheetContext method
async getSheetContext(): Promise<SheetContext> {
  // ... existing code ...

  // Add optimization for large datasets
  const optimizationService = DataOptimizationService.getInstance();
  const optimizedData = await optimizationService.optimizeDataForLLM(rawData);

  return {
    ...context,
    optimizationSummary: optimizedData.summary,
    optimizedSchema: optimizedData.schema
  };
}
```

### 6.3 Add Schema Provision at Chat Start

```typescript
// In components/chat-sidebar.tsx, add to initial messages
const initialMessages = [
  {
    id: "1",
    role: "assistant" as const,
    content:
      "Hello! I'm your advanced spreadsheet assistant. How can I help you analyze your spreadsheet data today?",
  },
  // Add schema information if available
  ...(sheetContext
    ? [
        {
          id: "schema",
          role: "system" as const,
          content: `Available data: ${formatContextForAI(sheetContext)}`,
        },
      ]
    : []),
];
```

## How to Test?

1. **Large Dataset Testing**:

   ```typescript
   // Create test dataset with 5000+ rows
   const largeDataset = Array.from({ length: 5000 }, (_, i) => ({
     id: i,
     region: ["North", "South", "East", "West"][i % 4],
     sales: Math.random() * 10000,
     date: new Date(2024, i % 12, i % 28),
   }));

   // Test optimization
   const optimizationService = DataOptimizationService.getInstance();
   const result = await optimizationService.optimizeDataForLLM(largeDataset);
   expect(result.summary.sampledRows).toBeLessThanOrEqual(2000);
   ```

2. **Schema Detection Testing**:

   - Test with mixed data types
   - Verify currency detection
   - Test date detection

3. **Performance Testing**:
   - Test with 10k+ rows
   - Measure memory usage
   - Test aggregation performance

## Important Dependencies to Not Break

- **Existing sheet context** - Don't break current data extraction
- **Tool execution** - Ensure tools work with optimized data
- **Memory usage** - Don't cause memory leaks with large datasets
- **Response times** - Keep chat responsive

## Dependencies That Will Work Thanks to This

- **Large dataset analysis** - Handle big spreadsheets efficiently
- **Performance optimization** - Faster tool execution
- **Memory management** - Better resource utilization
- **User experience** - Responsive with large files

## Implementation Strategy (Least Disruptive)

1. **Start with sampling** - Implement basic sampling first
2. **Add schema detection** - Enhance data type recognition
3. **Implement aggregation** - Add intelligent aggregation
4. **Test with large datasets** - Verify performance
5. **Integrate with existing services** - Connect to UniverService

## Priority Order

1. Create DataOptimizationService with basic sampling
2. Implement schema detection and extraction
3. Add intelligent aggregation for large datasets
4. Integrate with UniverService.getSheetContext()
5. Add schema provision to chat initialization
6. Test with large datasets
7. Optimize performance and memory usage
