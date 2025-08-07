# 01. Core Tools Implementation (CRITICAL)

## What?

Implement the four core spreadsheet tools as specified in the integration spec:

- `list_columns` - Returns column names & row count
- `create_pivot_table` - Creates pivot tables with grouping and aggregation
- `calculate_total` - Calculates totals for specific columns
- `generate_chart` - Creates charts via QuickChart and inserts to Dashboard sheet

## Where?

- **File**: `app/api/chat/route.ts`
- **Dependencies**: `services/univerService.ts`, `services/comprehensiveSpreadsheetService.ts`
- **New files**: `services/chartService.ts`, `services/pivotTableService.ts`

## How?

### 1.1 list_columns Tool

```typescript
list_columns: tool({
  description: "Get column names and row count from the current sheet",
  parameters: z.object({}),
  execute: async () => {
    const context = await UniverService.getInstance().getSheetContext();
    return {
      columns: context.tables[0]?.headers || [],
      rowCount: context.tables[0]?.recordCount || 0,
      sheetName: context.worksheetName,
    };
  },
});
```

### 1.2 create_pivot_table Tool

```typescript
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
    // Implementation in pivotTableService.ts
  },
});
```

### 1.3 calculate_total Tool

```typescript
calculate_total: tool({
  description: "Calculate total for a specific column",
  parameters: z.object({
    column: z.string().describe("Column name or letter (e.g., 'Sales' or 'B')"),
    filter: z.record(z.any()).optional().describe("Optional filter criteria"),
  }),
  execute: async ({ column, filter }) => {
    // Implementation using UniverService
  },
});
```

### 1.4 generate_chart Tool

```typescript
generate_chart: tool({
  description: "Generate chart from data and insert to Dashboard sheet",
  parameters: z.object({
    data: z.any().describe("Data source (pivot result or range)"),
    x: z.string().describe("X-axis column"),
    y: z.string().describe("Y-axis column"),
    chart_type: z
      .enum(["bar", "line", "pie", "scatter"])
      .describe("Chart type"),
    title: z.string().describe("Chart title"),
  }),
  execute: async ({ data, x, y, chart_type, title }) => {
    // Implementation in chartService.ts
  },
});
```

## How to Test?

1. **Unit Tests**: Create `__tests__/api/chat.test.ts`

   - Test each tool with mock UniverService
   - Verify parameter validation
   - Test error handling

2. **Integration Tests**:

   - Load sample spreadsheet data
   - Test end-to-end tool execution
   - Verify spreadsheet updates

3. **Manual Testing**:

   ```bash
   # Test list_columns
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"list columns"}]}'

   # Test pivot table
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"create pivot table by region"}]}'
   ```

## Important Dependencies to Not Break

- **UniverService.getInstance()** - Singleton pattern must be preserved
- **Existing financial_intelligence tool** - Keep working alongside new tools
- **Chat message flow** - Don't break existing message handling
- **Error handling patterns** - Maintain consistent error responses

## Dependencies That Will Work Thanks to This

- **Agent loop implementation** - These tools enable multi-step execution
- **Chart generation** - Pivot tables provide data for charts
- **Financial analysis** - Basic tools support complex calculations
- **User experience** - Users can perform basic spreadsheet operations

## Implementation Strategy (Least Disruptive)

1. **Add tools incrementally** - One tool at a time
2. **Use existing service patterns** - Follow UniverService patterns
3. **Maintain backward compatibility** - Keep existing tools working
4. **Add comprehensive error handling** - Don't break existing flows
5. **Test thoroughly** - Each tool before moving to next

## Priority Order

1. `list_columns` (foundation for other tools)
2. `calculate_total` (simple, builds confidence)
3. `create_pivot_table` (complex, enables analysis)
4. `generate_chart` (depends on pivot tables)
