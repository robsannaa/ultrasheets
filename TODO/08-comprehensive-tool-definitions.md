# 08. Comprehensive Tool Definitions (CRITICAL)

## What?

Implement ALL spreadsheet functions identified from Univer AI documentation, properly categorized by Facade API vs Plugin mode, covering every aspect of spreadsheet operations.

## Where?

- **Files**: `app/api/chat/route.ts`, `services/univerService.ts`
- **New files**: `services/facadeOperationsService.ts`, `services/pluginOperationsService.ts`
- **Dependencies**: All Univer plugins and Facade API

## How?

### 8.1 Facade API vs Plugin Mode Decision Matrix

| Functionality             | API Type    | Rationale                        | Implementation                                              |
| ------------------------- | ----------- | -------------------------------- | ----------------------------------------------------------- |
| **Basic Cell Operations** | Facade API  | Simple, direct access            | `univerAPI.getActiveWorkbook().getActiveSheet().getRange()` |
| **Formatting**            | Plugin Mode | Complex styling requires plugins | `@univerjs/sheets-conditional-formatting`                   |
| **Charts**                | Plugin Mode | Advanced visualization features  | `@univerjs/sheets-charts`                                   |
| **Tables**                | Plugin Mode | Complex table operations         | `@univerjs/sheets-table`                                    |
| **Formulas**              | Plugin Mode | Formula engine integration       | `@univerjs/sheets-formula`                                  |
| **Data Validation**       | Plugin Mode | Validation rules and UI          | `@univerjs/sheets-data-validation`                          |
| **Comments**              | Plugin Mode | Thread comments system           | `@univerjs/sheets-thread-comment`                           |
| **Hyperlinks**            | Plugin Mode | Link management                  | `@univerjs/sheets-hyper-link`                               |
| **Drawing**               | Plugin Mode | Graphics and shapes              | `@univerjs/sheets-drawing`                                  |
| **Filtering/Sorting**     | Plugin Mode | Advanced data operations         | `@univerjs/sheets-filter`, `@univerjs/sheets-sort`          |

### 8.2 Complete Tool Definitions

#### 8.2.1 Basic Cell Operations (Facade API)

```typescript
// services/facadeOperationsService.ts
export class FacadeOperationsService {
  private univerAPI: any;

  constructor(univerAPI: any) {
    this.univerAPI = univerAPI;
  }

  // Cell Value Operations
  async setCellValue(cell: string, value: any): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const range = worksheet.getRange(cell);
    range.setValue(value);
    return `Set cell ${cell} to ${value}`;
  }

  async getCellValue(cell: string): Promise<any> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const range = worksheet.getRange(cell);
    return range.getValue();
  }

  // Range Operations
  async setRangeValues(range: string, values: any[][]): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    rangeObj.setValues(values);
    return `Set range ${range} with ${values.length} rows`;
  }

  async getRangeValues(range: string): Promise<any[][]> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    return rangeObj.getValues();
  }

  // Row/Column Operations
  async insertRows(startRow: number, count: number): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    worksheet.insertRows(startRow, count);
    return `Inserted ${count} rows starting at row ${startRow}`;
  }

  async deleteRows(startRow: number, count: number): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    worksheet.deleteRows(startRow, count);
    return `Deleted ${count} rows starting at row ${startRow}`;
  }

  async insertColumns(startCol: number, count: number): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    worksheet.insertColumns(startCol, count);
    return `Inserted ${count} columns starting at column ${startCol}`;
  }

  async deleteColumns(startCol: number, count: number): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    worksheet.deleteColumns(startCol, count);
    return `Deleted ${count} columns starting at column ${startCol}`;
  }

  // Clear Operations
  async clearRange(range: string): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    rangeObj.clear();
    return `Cleared range ${range}`;
  }

  async clearRangeContents(range: string): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    rangeObj.clearContents();
    return `Cleared contents of range ${range}`;
  }

  // Merge Operations
  async mergeCells(range: string): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    rangeObj.merge();
    return `Merged cells in range ${range}`;
  }

  async unmergeCells(range: string): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    rangeObj.unmerge();
    return `Unmerged cells in range ${range}`;
  }

  // Freeze Panes
  async freezePanes(row: number, column: number): Promise<string> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    worksheet.freezePanes(row, column);
    return `Froze panes at row ${row}, column ${column}`;
  }

  // Get Range Coordinates
  async getRangeCoordinates(range: string): Promise<{
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  }> {
    const workbook = this.univerAPI.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();
    const rangeObj = worksheet.getRange(range);
    return {
      startRow: rangeObj.getRow(),
      endRow: rangeObj.getLastRow(),
      startCol: rangeObj.getColumn(),
      endCol: rangeObj.getLastColumn(),
    };
  }
}
```

#### 8.2.2 Plugin-Based Operations Service

```typescript
// services/pluginOperationsService.ts
import { ICommandService } from "@univerjs/core";

export class PluginOperationsService {
  private commandService: ICommandService;

  constructor(commandService: ICommandService) {
    this.commandService = commandService;
  }

  // Conditional Formatting
  async applyConditionalFormatting(
    range: string,
    rules: any[]
  ): Promise<string> {
    // Use @univerjs/sheets-conditional-formatting plugin
    // Let Univer handle range parsing and rule processing
    for (const rule of rules) {
      await this.commandService.executeCommand(
        "sheet.command.add-conditional-formatting",
        {
          unitId: "current",
          subUnitId: "current",
          rule: {
            ...rule,
            ranges: [range], // Let Univer parse the range string
          },
        }
      );
    }
    return `Applied conditional formatting to ${range}`;
  }

  // Data Validation
  async addDataValidation(range: string, validationRule: any): Promise<string> {
    // Use @univerjs/sheets-data-validation plugin
    // Let Univer handle range parsing and validation rule processing
    await this.commandService.executeCommand(
      "sheet.command.add-data-validation",
      {
        unitId: "current",
        subUnitId: "current",
        rule: {
          ...validationRule,
          ranges: [range], // Let Univer parse the range string
        },
      }
    );
    return `Added data validation to ${range}`;
  }

  // Comments
  async addComment(cell: string, comment: string): Promise<string> {
    // Use @univerjs/sheets-thread-comment plugin
    // Let Univer handle cell parsing
    await this.commandService.executeCommand(
      "thread-comment.command.add-comment",
      {
        unitId: "current",
        subUnitId: "current",
        comment: {
          content: comment,
          cell: cell, // Let Univer parse the cell reference
        },
      }
    );
    return `Added comment to ${cell}`;
  }

  // Hyperlinks
  async addHyperlink(
    cell: string,
    url: string,
    text?: string
  ): Promise<string> {
    // Use @univerjs/sheets-hyper-link plugin
    // Let Univer handle cell parsing and link processing
    await this.commandService.executeCommand("sheet.command.add-hyperlink", {
      unitId: "current",
      subUnitId: "current",
      link: {
        payload: url,
        id: Date.now().toString(),
        cell: cell, // Let Univer parse the cell reference
      },
    });
    return `Added hyperlink to ${cell}`;
  }

  // Tables
  async createTable(
    range: string,
    hasHeaders: boolean = true
  ): Promise<string> {
    // Use @univerjs/sheets-table plugin
    // Let Univer handle range parsing and table creation
    await this.commandService.executeCommand("sheet.command.create-table", {
      unitId: "current",
      subUnitId: "current",
      range: range, // Let Univer parse the range string
      hasHeaders,
    });
    return `Created table in range ${range}`;
  }

  // Filtering
  async applyFilter(range: string, filterCriteria: any): Promise<string> {
    // Use @univerjs/sheets-filter plugin
    // Let Univer handle range parsing and filter processing
    await this.commandService.executeCommand("sheet.command.apply-filter", {
      unitId: "current",
      subUnitId: "current",
      range: range, // Let Univer parse the range string
      criteria: filterCriteria,
    });
    return `Applied filter to ${range}`;
  }

  // Sorting
  async sortRange(
    range: string,
    sortBy: string,
    ascending: boolean = true
  ): Promise<string> {
    // Use @univerjs/sheets-sort plugin
    // Let Univer handle range parsing and sort processing
    await this.commandService.executeCommand("sheet.command.sort-range", {
      unitId: "current",
      subUnitId: "current",
      range: range, // Let Univer parse the range string
      sortBy,
      ascending,
    });
    return `Sorted range ${range} by ${sortBy}`;
  }

  // Drawing/Graphics
  async addDrawing(range: string, drawingData: any): Promise<string> {
    // Use @univerjs/sheets-drawing plugin
    // Let Univer handle range parsing and drawing processing
    await this.commandService.executeCommand("sheet.command.add-drawing", {
      unitId: "current",
      subUnitId: "current",
      range: range, // Let Univer parse the range string
      drawing: drawingData,
    });
    return `Added drawing to ${range}`;
  }

  // Rich Text
  async setRichText(cell: string, richTextData: any): Promise<string> {
    // Use @univerjs/sheets-formula plugin for rich text
    // Let Univer handle cell parsing and rich text processing
    await this.commandService.executeCommand("sheet.command.set-rich-text", {
      unitId: "current",
      subUnitId: "current",
      cell: cell, // Let Univer parse the cell reference
      richText: richTextData,
    });
    return `Set rich text in ${cell}`;
  }
}
```

#### 8.2.3 Complete Tool Definitions for Chat Route

```typescript
// In app/api/chat/route.ts - Complete tool definitions
const tools = {
  // === BASIC CELL OPERATIONS (Facade API) ===
  set_cell_value: tool({
    description: "Set a cell value",
    parameters: z.object({
      cell: z.string().describe("Cell reference (e.g., 'A1')"),
      value: z.any().describe("Value to set"),
    }),
    execute: async ({ cell, value }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.setCellValue(cell, value);
    },
  }),

  get_cell_value: tool({
    description: "Get a cell value",
    parameters: z.object({
      cell: z.string().describe("Cell reference (e.g., 'A1')"),
    }),
    execute: async ({ cell }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.getCellValue(cell);
    },
  }),

  set_range_values: tool({
    description: "Set values for a range of cells",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:B5')"),
      values: z.array(z.array(z.any())).describe("2D array of values"),
    }),
    execute: async ({ range, values }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.setRangeValues(range, values);
    },
  }),

  // === ROW/COLUMN OPERATIONS (Facade API) ===
  insert_rows: tool({
    description: "Insert rows",
    parameters: z.object({
      startRow: z.number().describe("Starting row number (1-based)"),
      count: z.number().describe("Number of rows to insert"),
    }),
    execute: async ({ startRow, count }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.insertRows(startRow, count);
    },
  }),

  delete_rows: tool({
    description: "Delete rows",
    parameters: z.object({
      startRow: z.number().describe("Starting row number (1-based)"),
      count: z.number().describe("Number of rows to delete"),
    }),
    execute: async ({ startRow, count }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.deleteRows(startRow, count);
    },
  }),

  insert_columns: tool({
    description: "Insert columns",
    parameters: z.object({
      startCol: z.number().describe("Starting column number (1-based)"),
      count: z.number().describe("Number of columns to insert"),
    }),
    execute: async ({ startCol, count }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.insertColumns(startCol, count);
    },
  }),

  delete_columns: tool({
    description: "Delete columns",
    parameters: z.object({
      startCol: z.number().describe("Starting column number (1-based)"),
      count: z.number().describe("Number of columns to delete"),
    }),
    execute: async ({ startCol, count }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.deleteColumns(startCol, count);
    },
  }),

  // === CLEAR OPERATIONS (Facade API) ===
  clear_range: tool({
    description: "Clear a range (values and formatting)",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:B5')"),
    }),
    execute: async ({ range }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.clearRange(range);
    },
  }),

  clear_range_contents: tool({
    description: "Clear only contents of a range (keep formatting)",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:B5')"),
    }),
    execute: async ({ range }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.clearRangeContents(range);
    },
  }),

  // === MERGE OPERATIONS (Facade API) ===
  merge_cells: tool({
    description: "Merge cells in a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:B2')"),
    }),
    execute: async ({ range }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.mergeCells(range);
    },
  }),

  unmerge_cells: tool({
    description: "Unmerge cells in a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:B2')"),
    }),
    execute: async ({ range }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.unmergeCells(range);
    },
  }),

  // === FREEZE PANES (Facade API) ===
  freeze_panes: tool({
    description: "Freeze panes at specified row and column",
    parameters: z.object({
      row: z.number().describe("Row number to freeze at"),
      column: z.number().describe("Column number to freeze at"),
    }),
    execute: async ({ row, column }) => {
      const facadeService = new FacadeOperationsService(univerAPI);
      return await facadeService.freezePanes(row, column);
    },
  }),

  // === CONDITIONAL FORMATTING (Plugin Mode) ===
  apply_conditional_formatting: tool({
    description: "Apply conditional formatting to a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:B10')"),
      rules: z
        .array(
          z.object({
            type: z.enum([
              "greater_than",
              "less_than",
              "between",
              "contains_text",
              "duplicate_values",
            ]),
            value: z.any().describe("Value for the condition"),
            format: z.object({
              backgroundColor: z.string().optional(),
              fontColor: z.string().optional(),
              bold: z.boolean().optional(),
              italic: z.boolean().optional(),
            }),
          })
        )
        .describe("Array of formatting rules"),
    }),
    execute: async ({ range, rules }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.applyConditionalFormatting(range, rules);
    },
  }),

  // === DATA VALIDATION (Plugin Mode) ===
  add_data_validation: tool({
    description: "Add data validation to a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:A10')"),
      type: z.enum(["list", "number", "date", "text_length", "custom"]),
      criteria: z.any().describe("Validation criteria"),
      message: z.string().optional().describe("Error message"),
    }),
    execute: async ({ range, type, criteria, message }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.addDataValidation(range, {
        type,
        criteria,
        message,
      });
    },
  }),

  // === COMMENTS (Plugin Mode) ===
  add_comment: tool({
    description: "Add a comment to a cell",
    parameters: z.object({
      cell: z.string().describe("Cell reference (e.g., 'A1')"),
      comment: z.string().describe("Comment text"),
    }),
    execute: async ({ cell, comment }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.addComment(cell, comment);
    },
  }),

  // === HYPERLINKS (Plugin Mode) ===
  add_hyperlink: tool({
    description: "Add a hyperlink to a cell",
    parameters: z.object({
      cell: z.string().describe("Cell reference (e.g., 'A1')"),
      url: z.string().describe("URL for the hyperlink"),
      text: z.string().optional().describe("Display text (optional)"),
    }),
    execute: async ({ cell, url, text }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.addHyperlink(cell, url, text);
    },
  }),

  // === TABLES (Plugin Mode) ===
  create_table: tool({
    description: "Create a table from a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:D10')"),
      hasHeaders: z
        .boolean()
        .optional()
        .describe("Whether the range has headers"),
    }),
    execute: async ({ range, hasHeaders = true }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.createTable(range, hasHeaders);
    },
  }),

  // === FILTERING (Plugin Mode) ===
  apply_filter: tool({
    description: "Apply a filter to a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:D10')"),
      column: z.number().describe("Column to filter (1-based)"),
      criteria: z.any().describe("Filter criteria"),
    }),
    execute: async ({ range, column, criteria }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.applyFilter(range, { column, criteria });
    },
  }),

  // === SORTING (Plugin Mode) ===
  sort_range: tool({
    description: "Sort a range by a column",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:D10')"),
      sortBy: z.string().describe("Column to sort by (e.g., 'A' or 'Name')"),
      ascending: z
        .boolean()
        .optional()
        .describe("Sort order (true for ascending)"),
    }),
    execute: async ({ range, sortBy, ascending = true }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.sortRange(range, sortBy, ascending);
    },
  }),

  // === DRAWING (Plugin Mode) ===
  add_drawing: tool({
    description: "Add a drawing to a range",
    parameters: z.object({
      range: z.string().describe("Range reference (e.g., 'A1:D5')"),
      type: z.enum(["rectangle", "circle", "line", "arrow", "text"]),
      properties: z
        .object({
          width: z.number().optional(),
          height: z.number().optional(),
          color: z.string().optional(),
          text: z.string().optional(),
        })
        .describe("Drawing properties"),
    }),
    execute: async ({ range, type, properties }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.addDrawing(range, { type, ...properties });
    },
  }),

  // === RICH TEXT (Plugin Mode) ===
  set_rich_text: tool({
    description: "Set rich text formatting in a cell",
    parameters: z.object({
      cell: z.string().describe("Cell reference (e.g., 'A1')"),
      text: z.string().describe("Text content"),
      formatting: z
        .object({
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.boolean().optional(),
          color: z.string().optional(),
          fontSize: z.number().optional(),
        })
        .describe("Text formatting"),
    }),
    execute: async ({ cell, text, formatting }) => {
      const pluginService = new PluginOperationsService(commandService);
      return await pluginService.setRichText(cell, { text, formatting });
    },
  }),

  // === EXISTING TOOLS (Enhanced) ===
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
  }),

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
  }),

  calculate_total: tool({
    description: "Calculate total for a specific column",
    parameters: z.object({
      column: z
        .string()
        .describe("Column name or letter (e.g., 'Sales' or 'B')"),
      filter: z.record(z.any()).optional().describe("Optional filter criteria"),
    }),
    execute: async ({ column, filter }) => {
      const univerService = UniverService.getInstance();

      // Let Univer handle column parsing and formula generation
      await univerService.setFormula("H1", `SUM(${column}:${column})`);
      return `Calculated total for column ${column}`;
    },
  }),

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
  }),
};
```

## How to Test?

### 8.3 Testing Strategy

1. **Facade API Tests**:

   ```typescript
   describe("FacadeOperationsService", () => {
     it("should set cell value", async () => {
       const service = new FacadeOperationsService(mockUniverAPI);
       const result = await service.setCellValue("A1", "Test Value");
       expect(result).toContain("Set cell A1 to Test Value");
     });
   });
   ```

2. **Plugin Mode Tests**:

   ```typescript
   describe("PluginOperationsService", () => {
     it("should apply conditional formatting", async () => {
       const service = new PluginOperationsService(mockCommandService);
       const result = await service.applyConditionalFormatting("A1:B10", [
         {
           type: "greater_than",
           value: 100,
           format: { backgroundColor: "red" },
         },
       ]);
       expect(result).toContain("Applied conditional formatting");
     });
   });
   ```

3. **Integration Tests**:

   ```bash
   # Test basic operations
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"set cell A1 to 100"}]}'

   # Test complex operations
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"create a table from A1:D10 with headers"}]}'
   ```

## Important Dependencies to Not Break

- **Existing UniverService singleton** - Must continue working
- **Current Facade API usage** - Don't break existing operations
- **Plugin registration** - Ensure all plugins are properly registered
- **Command service** - Maintain command execution patterns

## Dependencies That Will Work Thanks to This

- **Complete spreadsheet functionality** - All operations available
- **Professional-grade features** - Enterprise-level capabilities
- **User experience** - Full Excel-like functionality
- **Future enhancements** - Foundation for advanced features

## Implementation Strategy (Least Disruptive)

1. **Start with Facade API** - Implement basic operations first
2. **Add Plugin Mode gradually** - One plugin at a time
3. **Test thoroughly** - Each operation before moving to next
4. **Maintain compatibility** - Keep existing functionality working
5. **Document usage** - Clear examples for each tool

## 🚨 CRITICAL: Avoid Hardcoding and Manual Mapping

### ❌ **BAD PRACTICES TO AVOID**

1. **Hardcoded Color Mappings** - Never create static color maps like:

   ```typescript
   // ❌ BAD - Hardcoded color mapping
   const colorMap: { [key: string]: string } = {
     blue: "#0000FF",
     red: "#FF0000",
     green: "#00FF00",
     // ... more hardcoded values
   };
   ```

2. **Hardcoded Natural Language Parsing** - Never manually parse natural language patterns:

   ```typescript
   // ❌ BAD - Hardcoded natural language parsing
   if (
     lowerTarget === "header" ||
     lowerTarget === "headers" ||
     lowerTarget === "first row"
   ) {
     return "A1:Z1"; // First row
   }
   if (lowerTarget === "first column" || lowerTarget === "column a") {
     return "A:A"; // First column
   }
   if (lowerTarget.match(/^column [a-z]$/)) {
     const col = lowerTarget.split(" ")[1].toUpperCase();
     return `${col}:${col}`;
   }
   ```

3. **Manual Data Transformations** - Don't manually transform data formats
4. **Static Configuration Objects** - Avoid hardcoded configuration mappings
5. **Manual Range Parsing** - Don't manually parse cell references
6. **Hardcoded Validation Rules** - Don't create static validation mappings

### ✅ **CORRECT APPROACHES**

1. **Use Univer's Built-in APIs** - Let Univer handle all transformations:

   ```typescript
   // ✅ GOOD - Use Univer's native color handling
   await this.commandService.executeCommand("sheet.command.set-cell-format", {
     unitId: "current",
     subUnitId: "current",
     cell: { row: 0, column: 0 },
     format: { backgroundColor: "blue" }, // Let Univer handle color mapping
   });
   ```

2. **Leverage Univer's Range Parsing** - Use built-in range utilities:

   ```typescript
   // ✅ GOOD - Use Univer's range parsing
   const range = this.univerAPI.parseRange("A1:B10");
   ```

3. **Use Plugin Commands** - Let plugins handle their own data:

   ```typescript
   // ✅ GOOD - Let conditional formatting plugin handle rules
   await this.commandService.executeCommand(
     "sheet.command.add-conditional-formatting",
     {
       unitId: "current",
       subUnitId: "current",
       rule: rule, // Pass the rule as-is, let plugin handle it
     }
   );
   ```

4. **Dynamic Configuration** - Use environment variables or dynamic loading:

   ```typescript
   // ✅ GOOD - Dynamic configuration
   const config = await this.loadConfiguration();
   ```

5. **Let LLM Agent Handle Natural Language** - Don't parse natural language in services:
   ```typescript
   // ✅ GOOD - Let the LLM agent determine the range
   // The agent should convert "header row" to "A1:Z1" before calling the tool
   await this.commandService.executeCommand("sheet.command.set-range-format", {
     unitId: "current",
     subUnitId: "current",
     range: "A1:Z1", // Agent provides exact range, not natural language
     format: { bold: true },
   });
   ```

### **Implementation Rules**

1. **Never hardcode mappings** - Always use Univer's built-in capabilities
2. **Let plugins handle their data** - Don't pre-process data for plugins
3. **Use command service** - Execute commands, don't manipulate data directly
4. **Dynamic everything** - Configuration should be loaded, not hardcoded
5. **Agent-driven** - Let the LLM agent determine the appropriate actions
6. **No natural language parsing in services** - The LLM agent should convert natural language to exact references before calling tools
7. **Tools expect exact inputs** - All tool parameters should be precise (e.g., "A1:Z1", not "header row")

## Required Plugin Dependencies

```bash
# Core plugins
pnpm add @univerjs/sheets-conditional-formatting
pnpm add @univerjs/sheets-data-validation
pnpm add @univerjs/sheets-thread-comment
pnpm add @univerjs/sheets-hyper-link
pnpm add @univerjs/sheets-drawing
pnpm add @univerjs/sheets-filter
pnpm add @univerjs/sheets-sort
pnpm add @univerjs/sheets-table
pnpm add @univerjs/sheets-formula

# UI plugins
pnpm add @univerjs/sheets-conditional-formatting-ui
pnpm add @univerjs/sheets-data-validation-ui
pnpm add @univerjs/sheets-thread-comment-ui
pnpm add @univerjs/sheets-hyper-link-ui
pnpm add @univerjs/sheets-drawing-ui
pnpm add @univerjs/sheets-filter-ui
pnpm add @univerjs/sheets-sort-ui
pnpm add @univerjs/sheets-table-ui
pnpm add @univerjs/sheets-formula-ui
```

## Priority Order

1. Implement FacadeOperationsService with basic cell operations
2. Add row/column operations (insert/delete)
3. Add clear and merge operations
4. Add freeze panes functionality
5. Implement PluginOperationsService with conditional formatting
6. Add data validation capabilities
7. Add comments and hyperlinks
8. Add table operations
9. Add filtering and sorting
10. Add drawing and rich text capabilities
11. Integrate all tools into chat route
12. Add comprehensive error handling and testing
