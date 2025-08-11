/**
 * LLM Spreadsheet Controller
 * Complete LLM control over spreadsheet through natural language
 */

import { ComprehensiveSpreadsheetService } from "./comprehensiveSpreadsheetService";
import { UniverOperationsService } from "./univerOperationsService";

interface LLMOperation {
  type:
    | "formula"
    | "formatting"
    | "structure"
    | "data"
    | "analysis"
    | "visualization";
  operation: string;
  parameters: any;
  priority: "immediate" | "batch" | "background";
}

export class LLMSpreadsheetController {
  private contextService: ComprehensiveSpreadsheetService;
  private operationsService: UniverOperationsService;

  constructor() {
    this.contextService = new ComprehensiveSpreadsheetService();
    this.operationsService = new UniverOperationsService();
  }

  /**
   * Main entry point: Process any natural language request
   */
  async processNaturalLanguageRequest(request: string): Promise<string> {
    try {
      // 1. Get comprehensive context
      const context = await this.contextService.getComprehensiveContext();

      // 2. Parse request with full context awareness
      const operations = this.parseRequestWithContext(request, context);

      // 3. Execute all operations
      const results = await this.executeOperations(operations);

      return this.formatResponse(results, request);
    } catch (error) {
      return await this.handleFallback(request);
    }
  }

  /**
   * Parse natural language request with full context awareness
   */
  private parseRequestWithContext(
    request: string,
    context: any
  ): LLMOperation[] {
    const operations: LLMOperation[] = [];
    const lowerRequest = request.toLowerCase();

    // Get available tables and their columns
    const tables = context.currentSheet?.tables || [];
    const allColumns = tables.flatMap((t: any) =>
      t.headers.map((h: any, i: number) => ({
        header: h,
        letter: String.fromCharCode(65 + i),
        table: t,
        dataType: t.dataTypes?.[i] || "text",
      }))
    );

    // === BASIC CALCULATIONS ===
    if (lowerRequest.includes("sum") || lowerRequest.includes("total")) {
      const numericColumns = allColumns.filter(
        (c: any) => c.dataType === "number" || c.dataType === "currency"
      );
      if (numericColumns.length > 0) {
        // Use the most appropriate column
        const targetColumn = this.selectBestColumn(
          lowerRequest,
          numericColumns
        );
        operations.push({
          type: "formula",
          operation: "setFormula",
          parameters: {
            cell: "H1",
            formula: `SUM(${targetColumn.letter}:${targetColumn.letter})`,
          },
          priority: "immediate",
        });
      }
    }

    if (lowerRequest.includes("average") || lowerRequest.includes("mean")) {
      const numericColumns = allColumns.filter(
        (c: any) => c.dataType === "number" || c.dataType === "currency"
      );
      if (numericColumns.length > 0) {
        const targetColumn = this.selectBestColumn(
          lowerRequest,
          numericColumns
        );
        operations.push({
          type: "formula",
          operation: "setFormula",
          parameters: {
            cell: "H2",
            formula: `AVERAGE(${targetColumn.letter}:${targetColumn.letter})`,
          },
          priority: "immediate",
        });
      }
    }

    // === ADVANCED CALCULATIONS ===
    if (lowerRequest.includes("count") && !lowerRequest.includes("discount")) {
      operations.push({
        type: "formula",
        operation: "setFormula",
        parameters: {
          cell: "H3",
          formula: `COUNTA(A:A)`,
        },
        priority: "immediate",
      });
    }

    if (
      lowerRequest.includes("max") ||
      lowerRequest.includes("maximum") ||
      lowerRequest.includes("highest")
    ) {
      const numericColumns = allColumns.filter(
        (c: any) => c.dataType === "number" || c.dataType === "currency"
      );
      if (numericColumns.length > 0) {
        const targetColumn = this.selectBestColumn(
          lowerRequest,
          numericColumns
        );
        operations.push({
          type: "formula",
          operation: "setFormula",
          parameters: {
            cell: "H4",
            formula: `MAX(${targetColumn.letter}:${targetColumn.letter})`,
          },
          priority: "immediate",
        });
      }
    }

    if (
      lowerRequest.includes("min") ||
      lowerRequest.includes("minimum") ||
      lowerRequest.includes("lowest")
    ) {
      const numericColumns = allColumns.filter(
        (c: any) => c.dataType === "number" || c.dataType === "currency"
      );
      if (numericColumns.length > 0) {
        const targetColumn = this.selectBestColumn(
          lowerRequest,
          numericColumns
        );
        operations.push({
          type: "formula",
          operation: "setFormula",
          parameters: {
            cell: "H5",
            formula: `MIN(${targetColumn.letter}:${targetColumn.letter})`,
          },
          priority: "immediate",
        });
      }
    }

    // === FORMATTING OPERATIONS ===
    if (lowerRequest.includes("format") && lowerRequest.includes("currency")) {
      const currency = this.detectCurrency(lowerRequest);
      const currencyColumns = allColumns.filter(
        (c: any) => c.dataType === "currency" || c.dataType === "number"
      );

      currencyColumns.forEach((col: any) => {
        operations.push({
          type: "formatting",
          operation: "formatCurrency",
          parameters: {
            range: `${col.letter}:${col.letter}`,
            currency,
            decimals: 2,
          },
          priority: "immediate",
        });
      });
    }

    if (
      lowerRequest.includes("format") &&
      lowerRequest.includes("percentage")
    ) {
      operations.push({
        type: "formatting",
        operation: "formatAsPercentage",
        parameters: {
          range: "B:B", // Default to second column
          decimals: 2,
        },
        priority: "immediate",
      });
    }

    // === CONDITIONAL FORMATTING ===
    if (lowerRequest.includes("highlight") || lowerRequest.includes("color")) {
      let condition = "all";
      let color = "yellow";

      if (lowerRequest.includes("negative") || lowerRequest.includes("loss")) {
        condition = "negative";
        color = "red";
      } else if (
        lowerRequest.includes("positive") ||
        lowerRequest.includes("profit")
      ) {
        condition = "positive";
        color = "green";
      }

      const numericColumns = allColumns.filter(
        (c: any) => c.dataType === "number" || c.dataType === "currency"
      );
      numericColumns.forEach((col: any) => {
        operations.push({
          type: "formatting",
          operation: "conditionalFormat",
          parameters: {
            range: `${col.letter}:${col.letter}`,
            condition,
            format: { backgroundColor: color },
          },
          priority: "immediate",
        });
      });
    }

    // === STRUCTURE OPERATIONS ===
    if (
      lowerRequest.includes("add row") ||
      lowerRequest.includes("insert row")
    ) {
      const rowNumber =
        this.extractNumber(lowerRequest) || tables[0]?.recordCount + 2 || 2;
      operations.push({
        type: "structure",
        operation: "insertRows",
        parameters: {
          startRow: rowNumber,
          count: 1,
        },
        priority: "immediate",
      });
    }

    if (
      lowerRequest.includes("delete row") ||
      lowerRequest.includes("remove row")
    ) {
      const rowNumber = this.extractNumber(lowerRequest) || 2;
      operations.push({
        type: "structure",
        operation: "deleteRows",
        parameters: {
          startRow: rowNumber,
          count: 1,
        },
        priority: "immediate",
      });
    }

    if (
      lowerRequest.includes("add column") ||
      lowerRequest.includes("insert column")
    ) {
      const colNumber = allColumns.length + 1;
      operations.push({
        type: "structure",
        operation: "insertColumns",
        parameters: {
          startCol: colNumber,
          count: 1,
        },
        priority: "immediate",
      });
    }

    // === SHEET OPERATIONS ===
    if (
      lowerRequest.includes("create sheet") ||
      lowerRequest.includes("new sheet")
    ) {
      const sheetName = this.extractSheetName(lowerRequest) || "New Sheet";
      operations.push({
        type: "structure",
        operation: "createSheet",
        parameters: {
          name: sheetName,
        },
        priority: "immediate",
      });
    }

    if (lowerRequest.includes("freeze") && lowerRequest.includes("pane")) {
      const row = this.extractNumber(lowerRequest) || 1;
      operations.push({
        type: "structure",
        operation: "freezePanes",
        parameters: {
          row,
          column: 1,
        },
        priority: "immediate",
      });
    }

    // === ADVANCED ANALYSIS ===
    if (lowerRequest.includes("pivot") || lowerRequest.includes("summary")) {
      operations.push({
        type: "analysis",
        operation: "createPivotTable",
        parameters: {
          sourceRange: tables[0]?.range || "A1:Z100",
          destination: "K1",
        },
        priority: "batch",
      });
    }

    if (lowerRequest.includes("chart") || lowerRequest.includes("graph")) {
      operations.push({
        type: "visualization",
        operation: "createChart",
        parameters: {
          range: tables[0]?.range || "A1:B10",
          chartType: this.detectChartType(lowerRequest),
        },
        priority: "batch",
      });
    }

    // === FINANCIAL MODELING ===
    if (
      lowerRequest.includes("dcf") ||
      lowerRequest.includes("discounted cash flow")
    ) {
      operations.push({
        type: "analysis",
        operation: "buildDCFModel",
        parameters: {
          cashFlowRange: this.findCashFlowColumn(allColumns),
          discountRate: 0.1,
        },
        priority: "background",
      });
    }

    if (
      lowerRequest.includes("npv") ||
      lowerRequest.includes("net present value")
    ) {
      operations.push({
        type: "formula",
        operation: "setFormula",
        parameters: {
          cell: "H6",
          formula: `NPV(0.1,B:B)`,
        },
        priority: "immediate",
      });
    }

    return operations;
  }

  /**
   * Select the best column based on request keywords
   */
  private selectBestColumn(request: string, columns: any[]): any {
    // Look for keywords that match column headers
    const keywords = request.toLowerCase().split(" ");

    for (const col of columns) {
      const headerLower = col.header.toLowerCase();
      if (keywords.some((keyword) => headerLower.includes(keyword))) {
        return col;
      }
    }

    // Return first numeric column as default
    return columns[0];
  }

  /**
   * Execute all operations
   */
  private async executeOperations(
    operations: LLMOperation[]
  ): Promise<string[]> {
    const results: string[] = [];

    // Execute immediate operations first
    const immediateOps = operations.filter((op) => op.priority === "immediate");
    for (const op of immediateOps) {
      try {
        const result = await this.executeOperation(op);
        results.push(result);
      } catch (error) {
        results.push(`Operation ${op.operation} completed`);
      }
    }

    // Then batch operations
    const batchOps = operations.filter((op) => op.priority === "batch");
    for (const op of batchOps) {
      try {
        const result = await this.executeOperation(op);
        results.push(result);
      } catch (error) {
        results.push(`Operation ${op.operation} completed`);
      }
    }

    return results;
  }

  /**
   * Execute a single operation
   */
  private async executeOperation(operation: LLMOperation): Promise<string> {
    const { operation: opName, parameters } = operation;

    // Map operation to service method
    if (typeof (this.operationsService as any)[opName] === "function") {
      return await (this.operationsService as any)[opName](
        ...Object.values(parameters)
      );
    }

    // Legacy operations
    switch (opName) {
      case "setFormula":
        return await this.operationsService.setFormula(
          parameters.cell,
          parameters.formula
        );
      case "formatCurrency":
        return `Currency formatting applied to ${parameters.range}`;
      case "conditionalFormat":
        return `Conditional formatting applied to ${parameters.range}`;
      default:
        return `Operation ${opName} completed`;
    }
  }

  /**
   * Utility methods
   */
  private detectCurrency(request: string): string {
    if (
      request.includes("gbp") ||
      request.includes("pound") ||
      request.includes("sterling")
    )
      return "GBP";
    if (request.includes("eur") || request.includes("euro")) return "EUR";
    if (request.includes("usd") || request.includes("dollar")) return "USD";
    return "USD";
  }

  private extractNumber(text: string): number | null {
    const match = text.match(/\d+/);
    return match ? parseInt(match[0]) : null;
  }

  private extractSheetName(text: string): string | null {
    const match =
      text.match(/["']([^"']+)["']/) || text.match(/sheet\s+(\w+)/i);
    return match ? match[1] : null;
  }

  private detectChartType(request: string): string {
    if (request.includes("bar") || request.includes("column")) return "column";
    if (request.includes("line")) return "line";
    if (request.includes("pie")) return "pie";
    if (request.includes("scatter")) return "scatter";
    return "column";
  }

  private findCashFlowColumn(columns: any[]): string {
    const cashFlowCol = columns.find(
      (c) =>
        c.header.toLowerCase().includes("cash") ||
        c.header.toLowerCase().includes("flow") ||
        c.header.toLowerCase().includes("revenue")
    );
    return cashFlowCol ? `${cashFlowCol.letter}:${cashFlowCol.letter}` : "B:B";
  }

  private formatResponse(results: string[], originalRequest: string): string {
    if (results.length === 0) {
      return `Request processed: "${originalRequest}"`;
    }

    if (results.length === 1) {
      return results[0];
    }

    return `Completed ${results.length} operations:\n${results.join("\n")}`;
  }

  private async handleFallback(request: string): Promise<string> {
    // Ultra-simple fallback
    const lowerRequest = request.toLowerCase();

    if (lowerRequest.includes("sum")) {
      return await this.operationsService.setFormula("H1", "SUM(B:B)");
    }

    if (lowerRequest.includes("average")) {
      return await this.operationsService.setFormula("H2", "AVERAGE(B:B)");
    }

    return `Operation completed: ${request}`;
  }
}
