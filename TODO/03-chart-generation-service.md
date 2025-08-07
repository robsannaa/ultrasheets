# 03. Chart Generation Service (HIGH)

## What?

Implement the `generate_chart` tool using QuickChart to create PNG charts and insert them into a "Dashboard" sheet, as specified in Section 3.4 of the integration spec.

## Where?

- **New file**: `services/chartService.ts`
- **Dependencies**: QuickChart API, UniverService
- **Integration**: `app/api/chat/route.ts` (add to tools)

## How?

### 3.1 Install QuickChart Dependency

```bash
pnpm add quickchart-js
```

### 3.2 Create ChartService

```typescript
// services/chartService.ts
import QuickChart from "quickchart-js";

export class ChartService {
  private static instance: ChartService;

  static getInstance(): ChartService {
    if (!ChartService.instance) {
      ChartService.instance = new ChartService();
    }
    return ChartService.instance;
  }

  async generateChart(
    data: any[],
    xColumn: string,
    yColumn: string,
    chartType: string,
    title: string
  ): Promise<{
    chartUrl: string;
    dashboardCell: string;
  }> {
    // 1. Transform data for QuickChart
    const chartData = this.transformDataForChart(
      data,
      xColumn,
      yColumn,
      chartType
    );

    // 2. Create QuickChart instance
    const chart = new QuickChart();
    chart.setConfig({
      type: chartType,
      data: chartData,
      options: {
        title: { display: true, text: title },
        responsive: true,
        plugins: {
          legend: { display: true },
        },
      },
    });

    // 3. Generate PNG URL
    const chartUrl = await chart.getUrl();

    // 4. Insert into Dashboard sheet
    const dashboardCell = await this.insertChartToDashboard(chartUrl, title);

    return { chartUrl, dashboardCell };
  }

  private transformDataForChart(
    data: any[],
    xColumn: string,
    yColumn: string,
    chartType: string
  ) {
    // Transform pivot table or range data to QuickChart format
    if (chartType === "bar" || chartType === "line") {
      return {
        labels: data.map((row) => row[xColumn]),
        datasets: [
          {
            label: yColumn,
            data: data.map((row) => row[yColumn]),
          },
        ],
      };
    }
    // Handle other chart types...
  }

  private async insertChartToDashboard(
    chartUrl: string,
    title: string
  ): Promise<string> {
    const univerService = UniverService.getInstance();

    // 1. Ensure Dashboard sheet exists
    await this.ensureDashboardSheet();

    // 2. Find next available cell for chart
    const nextCell = await this.findNextChartCell();

    // 3. Insert chart image (as URL reference)
    await univerService.setCellValue(nextCell, `=IMAGE("${chartUrl}")`);

    // 4. Add title above chart
    const titleCell = this.getCellAbove(nextCell);
    await univerService.setCellValue(titleCell, title);

    return nextCell;
  }

  private async ensureDashboardSheet(): Promise<void> {
    // Check if Dashboard sheet exists, create if not
    // Implementation depends on Univer API for sheet management
  }

  private async findNextChartCell(): Promise<string> {
    // Find next available cell in Dashboard sheet
    // Start from A1, move right/down as needed
    return "A1"; // Placeholder
  }
}
```

### 3.3 Update generate_chart Tool

```typescript
// In app/api/chat/route.ts
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
    const chartService = ChartService.getInstance();
    const result = await chartService.generateChart(
      data,
      x,
      y,
      chart_type,
      title
    );

    return {
      message: `Generated ${chart_type} chart: "${title}"`,
      chartUrl: result.chartUrl,
      location: `Dashboard sheet, cell ${result.dashboardCell}`,
      chartType: chart_type,
    };
  },
});
```

## How to Test?

1. **Unit Tests**: `__tests__/services/chartService.test.ts`

   ```typescript
   describe("ChartService", () => {
     it("should generate bar chart from pivot data", async () => {
       const service = ChartService.getInstance();
       const result = await service.generateChart(
         [
           { Region: "North", Sales: 1000 },
           { Region: "South", Sales: 1500 },
         ],
         "Region",
         "Sales",
         "bar",
         "Sales by Region"
       );
       expect(result.chartUrl).toContain("quickchart.io");
       expect(result.dashboardCell).toBeDefined();
     });
   });
   ```

2. **Integration Tests**:

   - Test with real pivot table data
   - Verify chart appears in Dashboard sheet
   - Test different chart types

3. **Manual Testing**:
   ```bash
   # Test chart generation
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"create a bar chart of sales by region"}]}'
   ```

## Important Dependencies to Not Break

- **UniverService singleton** - Don't break existing spreadsheet operations
- **Existing tool execution** - Keep other tools working
- **Image handling** - Ensure Univer supports IMAGE() function
- **Sheet management** - Don't break existing sheets

## Dependencies That Will Work Thanks to This

- **Pivot table analysis** - Charts visualize pivot results
- **Financial reporting** - Professional chart output
- **User experience** - Visual data representation
- **Dashboard functionality** - Centralized chart management

## Implementation Strategy (Least Disruptive)

1. **Start with QuickChart integration** - Test chart generation first
2. **Add Dashboard sheet creation** - Ensure sheet management works
3. **Test with simple data** - Use mock data before real pivot results
4. **Add error handling** - Handle chart generation failures
5. **Optimize performance** - Cache chart URLs when possible

## Priority Order

1. Install and test QuickChart
2. Create ChartService with basic functionality
3. Implement Dashboard sheet management
4. Add chart insertion logic
5. Integrate with generate_chart tool
6. Add comprehensive error handling
