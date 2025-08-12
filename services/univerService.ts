// services/univerService.ts
export interface SheetContext {
  workbookName: string;
  worksheetName: string;
  totalRows: number;
  totalColumns: number;
  dataRange: string;
  tables: Array<{
    range: string;
    headers: string[];
    dataPreview: string[][];
  }>;
  summaryData: {
    cellsWithData: number;
    formulas: number;
    charts: number;
    pivot_tables: number;
  };
}

export class UniverService {
  private static instance: UniverService;
  private univerInstance: any; // TODO: Add proper Univer API types

  static getInstance(): UniverService {
    if (!UniverService.instance) {
      UniverService.instance = new UniverService();
    }
    return UniverService.instance;
  }

  setUniverInstance(instance: any) {
    // TODO: Add proper Univer API types
    console.log("🔗 UniverService.setUniverInstance() called:", {
      hasInstance: !!instance,
      instanceType: typeof instance,
      methods: instance
        ? Object.keys(instance).filter(
            (key) => typeof instance[key] === "function"
          )
        : [],
    });
    this.univerInstance = instance;
    console.log("💾 UniverService: Instance stored, current state:", {
      hasInstance: !!this.univerInstance,
      ready: this.ready,
      instanceType: this.univerInstance
        ? typeof this.univerInstance
        : "undefined",
    });
  }

  isInitialized(): boolean {
    const initialized = !!this.univerInstance;
    console.log("🔍 UniverService.isInitialized() check:", {
      initialized,
      hasInstance: !!this.univerInstance,
      instanceType: this.univerInstance
        ? typeof this.univerInstance
        : "undefined",
    });
    return initialized;
  }

  private ready: boolean = false;

  setReady(ready: boolean) {
    console.log("🔄 UniverService.setReady() called:", ready);
    this.ready = ready;
    console.log("💾 UniverService: Ready state updated, current state:", {
      ready: this.ready,
      hasInstance: !!this.univerInstance,
      instanceType: this.univerInstance
        ? typeof this.univerInstance
        : "undefined",
    });
  }

  isReady(): boolean {
    const ready = this.ready && !!this.univerInstance;
    console.log("🔍 UniverService.isReady() check:", {
      ready,
      internalReady: this.ready,
      hasInstance: !!this.univerInstance,
      instanceType: this.univerInstance
        ? typeof this.univerInstance
        : "undefined",
    });
    return ready;
  }

  // DEPRECATED: Use setFormula() instead
  async sumRange(range: string, targetCell: string): Promise<string> {
    console.warn("sumRange() is deprecated. Use setFormula() instead.");
    return this.setFormula(targetCell, `SUM(${range})`);
  }

  async colorCells(target: string, color: string): Promise<string> {
    try {
      // Check if Univer instance is available and ready
      if (!this.isReady()) {
        throw new Error(
          "Univer instance not ready - please wait for the spreadsheet to load"
        );
      }

      // Parse the target to get the actual range
      const range = this.parseTarget(target);
      if (!range) {
        throw new Error(`Could not parse target: ${target}`);
      }

      // Convert color name to hex if needed
      const hexColor = this.parseColor(color);
      if (!hexColor) {
        throw new Error(`Invalid color: ${color}`);
      }

      // Get the active workbook and worksheet using facade API
      const fWorkbook = this.univerInstance.getActiveWorkbook();
      if (!fWorkbook) {
        throw new Error("Could not get active workbook");
      }

      const fWorksheet = fWorkbook.getActiveSheet();
      if (!fWorksheet) {
        throw new Error("Could not get active worksheet");
      }

      // Get the range and apply the color
      const fRange = fWorksheet.getRange(range);
      try {
        if (typeof fRange.setBackground === "function") {
          fRange.setBackground(hexColor);
        } else {
          throw new Error("setBackground method not available on range object");
        }
      } catch (error) {
        throw error;
      }

      return `Successfully colored ${range} with ${color}`;
    } catch (error) {
      throw new Error(
        `Color operation failed: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Helper method to validate cell references
  private validateCellReference(cell: string): boolean {
    return /^[A-Z]+\d+$/.test(cell);
  }

  // Helper method to parse range
  private parseRange(range: string): {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  } | null {
    const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) return null;

    const [, startCol, startRow, endCol, endRow] = match;
    return {
      startRow: parseInt(startRow),
      endRow: parseInt(endRow),
      startCol: this.columnToNumber(startCol),
      endCol: this.columnToNumber(endCol),
    };
  }

  // Helper method to convert column letter to number
  columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }

  // Helper method to parse a single cell reference
  private parseCell(cell: string): { row: number; col: number } | null {
    const match = cell.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;

    const [, col, row] = match;
    return {
      row: parseInt(row),
      col: this.columnToNumber(col),
    };
  }

  // Helper method to parse natural language targets into ranges
  private parseTarget(target: string): string | null {
    const lowerTarget = target.toLowerCase().trim();

    // Handle common natural language patterns
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
    if (lowerTarget === "second row") {
      return "A2:Z2";
    }
    if (lowerTarget === "third row") {
      return "A3:Z3";
    }
    if (lowerTarget.match(/^column [a-z]$/)) {
      const col = lowerTarget.split(" ")[1].toUpperCase();
      return `${col}:${col}`;
    }
    if (lowerTarget.match(/^row \d+$/)) {
      const row = lowerTarget.split(" ")[1];
      return `A${row}:Z${row}`;
    }

    // If it's already a valid range or cell reference, return as is
    if (this.isValidRange(target) || this.isValidCell(target)) {
      return target;
    }

    return null;
  }

  // Helper method to parse color names to hex codes
  private parseColor(color: string): string | null {
    const lowerColor = color.toLowerCase().trim();

    // Common color mappings
    const colorMap: { [key: string]: string } = {
      blue: "#0000FF",
      red: "#FF0000",
      green: "#00FF00",
      yellow: "#FFFF00",
      orange: "#FFA500",
      purple: "#800080",
      pink: "#FFC0CB",
      brown: "#A52A2A",
      gray: "#808080",
      grey: "#808080",
      black: "#000000",
      white: "#FFFFFF",
      cyan: "#00FFFF",
      magenta: "#FF00FF",
      lime: "#00FF00",
      navy: "#000080",
      teal: "#008080",
      olive: "#808000",
      maroon: "#800000",
      silver: "#C0C0C0",
      gold: "#FFD700",
      indigo: "#4B0082",
      violet: "#EE82EE",
      coral: "#FF7F50",
      salmon: "#FA8072",
      turquoise: "#40E0D0",
      lavender: "#E6E6FA",
      beige: "#F5F5DC",
      ivory: "#FFFFF0",
      khaki: "#F0E68C",
      plum: "#DDA0DD",
      azure: "#F0FFFF",
      mint: "#F5FFFA",
      rose: "#FFE4E1",
      cream: "#FFFDD0",
      peach: "#FFDAB9",
      tan: "#D2B48C",
      charcoal: "#36454F",
      slate: "#708090",
      burgundy: "#800020",
      forest: "#228B22",
      sienna: "#A0522D",
      crimson: "#DC143C",
      tomato: "#FF6347",
      firebrick: "#B22222",
      darkred: "#8B0000",
      lightblue: "#ADD8E6",
      lightgreen: "#90EE90",
      lightyellow: "#FFFFE0",
      lightpink: "#FFB6C1",
      lightgray: "#D3D3D3",
      lightgrey: "#D3D3D3",
      darkblue: "#00008B",
      darkgreen: "#006400",
      darkgray: "#A9A9A9",
      darkgrey: "#A9A9A9",
    };

    // Check if it's a hex color
    if (color.match(/^#[0-9A-Fa-f]{6}$/)) {
      return color;
    }

    // Check if it's a named color
    if (colorMap[lowerColor]) {
      return colorMap[lowerColor];
    }

    return null;
  }

  // Helper method to validate if a string is a valid range
  private isValidRange(range: string): boolean {
    return /^[A-Z]+\d+:[A-Z]+\d+$/.test(range);
  }

  // Helper method to validate if a string is a valid cell reference
  private isValidCell(cell: string): boolean {
    return /^[A-Z]+\d+$/.test(cell);
  }

  // ==== NEW EXCEL CAPABILITIES ====

  // Get sheet context for AI awareness
  async getSheetContext(): Promise<SheetContext> {
    try {
      console.log("🚀 getSheetContext: Starting...");

      if (!this.isReady()) {
        console.log("❌ getSheetContext: Service not ready");
        throw new Error("Univer instance not ready");
      }

      console.log("✅ getSheetContext: Service is ready, getting workbook...");
      const workbook = this.univerInstance.getActiveWorkbook();
      console.log("📊 getSheetContext: Workbook check:", {
        hasWorkbook: !!workbook,
        workbookType: typeof workbook,
      });

      if (!workbook) {
        console.log("❌ getSheetContext: Could not get active workbook");
        throw new Error("Could not get active workbook");
      }

      console.log("✅ getSheetContext: Got workbook, getting worksheet...");
      const worksheet = workbook.getActiveSheet();
      console.log("📄 getSheetContext: Worksheet check:", {
        hasWorksheet: !!worksheet,
        worksheetType: typeof worksheet,
      });

      if (!worksheet) {
        console.log("❌ getSheetContext: Could not get active worksheet");
        throw new Error("Could not get active worksheet");
      }

      console.log("📊 Getting sheet context using proper Univer API...");

      // Use the proper Univer API as documented: fWorksheet.getSheet().getSnapshot()
      console.log("📸 getSheetContext: Getting snapshot...");
      const sheetSnapshot = worksheet.getSheet().getSnapshot();
      console.log("📋 Sheet snapshot:", {
        hasSnapshot: !!sheetSnapshot,
        snapshotType: typeof sheetSnapshot,
        hasCellData: sheetSnapshot ? !!sheetSnapshot.cellData : false,
        cellDataKeys:
          sheetSnapshot && sheetSnapshot.cellData
            ? Object.keys(sheetSnapshot.cellData)
            : [],
      });

      if (!sheetSnapshot) {
        console.log("❌ getSheetContext: Could not get sheet snapshot");
        throw new Error("Could not get sheet snapshot");
      }

      // Extract cell data from the snapshot
      const cellData = sheetSnapshot.cellData || {};
      console.log("📊 Cell data from snapshot:", {
        hasCellData: !!cellData,
        cellDataKeys: Object.keys(cellData),
        totalRows: Object.keys(cellData).length,
      });

      let cellsWithData = 0;
      let formulas = 0;
      let maxRow = 0;
      let maxCol = 0;

      // Process the cell data from the snapshot
      console.log("🔍 getSheetContext: Processing cell data...");
      for (const rowIndex in cellData) {
        const row = parseInt(rowIndex);
        maxRow = Math.max(maxRow, row);

        for (const colIndex in cellData[rowIndex]) {
          const col = parseInt(colIndex);
          maxCol = Math.max(maxCol, col);

          const cell = cellData[rowIndex][colIndex];
          if (
            cell &&
            ((cell.v !== undefined && cell.v !== null && cell.v !== "") ||
              cell.f)
          ) {
            cellsWithData++;
            if (cell.f) formulas++;
          }
        }
      }

      console.log(
        `📊 Found ${cellsWithData} cells with data, max row: ${maxRow}, max col: ${maxCol}`
      );

      // Detect tables from the cell data
      console.log("🔍 getSheetContext: Detecting tables...");
      const tables = this.detectTables(cellData, maxRow, maxCol);

      const dataRange =
        cellsWithData > 0
          ? `A1:${this.numberToColumn(maxCol + 1)}${maxRow + 1}`
          : "A1:A1";

      const context: SheetContext = {
        workbookName: sheetSnapshot.name || "Unnamed Workbook",
        worksheetName: sheetSnapshot.name || "Sheet1",
        totalRows: sheetSnapshot.rowCount || 1000,
        totalColumns: sheetSnapshot.columnCount || 26,
        dataRange,
        tables,
        summaryData: {
          cellsWithData,
          formulas,
          charts: 0,
          pivot_tables: 0,
        },
      };

      console.log("📊 Generated context:", context);
      return context;
    } catch (error) {
      console.error("❌ getSheetContext error:", error);
      throw new Error(
        `Failed to get sheet context: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Detect table-like structures in the data
  private detectTables(
    cellMatrix: any,
    maxRow: number,
    maxCol: number
  ): Array<{ range: string; headers: string[]; dataPreview: string[][] }> {
    const tables = [];

    console.log("🔍 Detecting tables in matrix:", {
      hasMatrix: !!cellMatrix,
      matrixKeys: Object.keys(cellMatrix || {}),
      maxRow,
      maxCol,
    });

    // Simple heuristic: look for consecutive rows with data starting from row 0
    if (cellMatrix && cellMatrix[0]) {
      const headers: string[] = [];
      let headerCols = 0;

      // Extract headers from first row
      for (let col = 0; col <= maxCol; col++) {
        const cell = cellMatrix[0][col];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
          headers.push(String(cell.v));
          headerCols = col + 1;
          console.log(`📋 Header [${col}]: ${cell.v}`);
        } else {
          break;
        }
      }

      console.log(`📊 Found ${headers.length} headers:`, headers);

      if (headers.length > 1) {
        // Look for data rows
        const dataPreview: string[][] = [];
        let dataRows = 0;

        for (let row = 1; row <= maxRow; row++) {
          if (cellMatrix[row]) {
            const rowData: string[] = [];
            let hasData = false;

            for (let col = 0; col < headerCols; col++) {
              const cell = cellMatrix[row][col];
              const value = cell && cell.v !== undefined ? String(cell.v) : "";
              rowData.push(value);
              if (value && value !== "") hasData = true;
            }

            if (hasData) {
              dataPreview.push(rowData);
              dataRows = row;
            }
          }
        }

        console.log(`📊 Found ${dataPreview.length} data rows`);

        if (dataPreview.length > 0) {
          const range = `A1:${this.numberToColumn(headerCols)}${dataRows + 1}`;
          tables.push({ range, headers, dataPreview });
          console.log(`✅ Detected table with range: ${range}`);
        }
      } else {
        console.log("⚠️ Not enough headers found for table detection");
      }
    } else {
      console.log("⚠️ No cell matrix or first row data available");
    }

    console.log(`📊 Table detection complete. Found ${tables.length} tables.`);
    return tables;
  }

  // Convert number to column letter
  private numberToColumn(num: number): string {
    let result = "";
    while (num > 0) {
      num--;
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26);
    }
    return result;
  }

  // Set cell value using Univer's proper API
  async setCellValue(cell: string, value: string | number): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();
      const range = worksheet.getRange(cell);

      // Use Univer's proper setValue method
      range.setValue(value);

      return `Cell ${cell} set to: ${value}`;
    } catch (error) {
      throw new Error(`Failed to set cell value: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }

  // Set formula in cell using Univer's proper API
  async setFormula(cell: string, formula: string): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const cleanFormula = formula.startsWith("=") ? formula : `=${formula}`;
      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();
      const range = worksheet.getRange(cell);

      // Use Univer's proper setValue method for formulas
      range.setValue(cleanFormula);

      return `Formula ${cleanFormula} set in cell ${cell}`;
    } catch (error) {
      throw new Error(`Failed to set formula: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }

  // Format cells using Univer's proper API
  async formatCells(
    range: string,
    formatting: {
      bold?: boolean;
      italic?: boolean;
      fontSize?: number;
      fontColor?: string;
      backgroundColor?: string;
    }
  ): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();
      const fRange = worksheet.getRange(range);

      // Apply formatting using Univer's proper methods
      if (formatting.bold !== undefined) {
        fRange.setFontWeight(formatting.bold ? "bold" : "normal");
      }
      if (formatting.italic !== undefined) {
        fRange.setFontStyle(formatting.italic ? "italic" : "normal");
      }
      if (formatting.fontSize) {
        fRange.setFontSize(formatting.fontSize);
      }
      if (formatting.fontColor) {
        fRange.setFontColor(formatting.fontColor);
      }
      if (formatting.backgroundColor) {
        fRange.setBackground(formatting.backgroundColor);
      }

      return `Formatted range ${range}`;
    } catch (error) {
      throw new Error(`Failed to format cells: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }

  // Insert row - WARNING: This API may not exist in Univer Facade
  async insertRow(rowIndex: number): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();

      // WARNING: insertRow may not be available in Univer Facade API
      if (typeof worksheet.insertRow !== "function") {
        throw new Error(
          "insertRow method not available in current Univer version"
        );
      }
      worksheet.insertRow(rowIndex);
      return `Successfully inserted row at index ${rowIndex}`;
    } catch (error) {
      throw new Error(
        `Failed to insert row: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Delete row - WARNING: This API may not exist in Univer Facade
  async deleteRow(rowIndex: number): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();

      // WARNING: deleteRow may not be available in Univer Facade API
      if (typeof worksheet.deleteRow !== "function") {
        throw new Error(
          "deleteRow method not available in current Univer version"
        );
      }
      worksheet.deleteRow(rowIndex);
      return `Successfully deleted row at index ${rowIndex}`;
    } catch (error) {
      throw new Error(
        `Failed to delete row: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Insert column - WARNING: This API may not exist in Univer Facade
  async insertColumn(colIndex: number): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();

      // WARNING: insertColumn may not be available in Univer Facade API
      if (typeof worksheet.insertColumns !== "function") {
        throw new Error(
          "insertColumns method not available in current Univer version"
        );
      }
      worksheet.insertColumns(colIndex, 1);
      return `Successfully inserted column at index ${colIndex} using insertColumns API`;
    } catch (error) {
      throw new Error(
        `Failed to insert column: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Delete column - WARNING: This API may not exist in Univer Facade
  async deleteColumn(colIndex: number): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();

      // WARNING: deleteColumn may not be available in Univer Facade API
      if (typeof worksheet.deleteColumn !== "function") {
        throw new Error(
          "deleteColumn method not available in current Univer version"
        );
      }
      worksheet.deleteColumn(colIndex);
      return `Successfully deleted column at index ${colIndex}`;
    } catch (error) {
      throw new Error(
        `Failed to delete column: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Smart range detection based on content
  async detectRange(
    description: string,
    context?: SheetContext
  ): Promise<string | null> {
    try {
      const lowerDesc = description.toLowerCase().trim();

      // Try financial-specific detection first
      const financialRange = await this.detectFinancialRange(
        description,
        context
      );
      if (financialRange) {
        return financialRange;
      }

      // Use context if available
      if (context) {
        // Look for tables
        if (
          lowerDesc.includes("table") ||
          lowerDesc.includes("data") ||
          lowerDesc.includes("p&l") ||
          lowerDesc.includes("balance sheet")
        ) {
          if (context.tables.length > 0) {
            return context.tables[0].range;
          }
        }

        // Look for headers
        if (lowerDesc.includes("header") || lowerDesc.includes("title")) {
          if (context.tables.length > 0) {
            const table = context.tables[0];
            const endCol = this.numberToColumn(table.headers.length);
            return `A1:${endCol}1`;
          }
          return "A1:Z1";
        }

        // Look for data range
        if (
          lowerDesc.includes("all data") ||
          lowerDesc.includes("everything") ||
          lowerDesc.includes("entire dataset")
        ) {
          return context.dataRange;
        }

        // Financial-specific terms
        if (
          lowerDesc.includes("amounts") ||
          lowerDesc.includes("values") ||
          lowerDesc.includes("totals")
        ) {
          const numericalCols = this.detectNumericalColumns(context.tables[0]);
          if (numericalCols.length > 0) {
            return `${numericalCols[0]}:${
              numericalCols[numericalCols.length - 1]
            }`;
          }
        }
      }

      // Fallback to existing parseTarget logic
      return this.parseTarget(description);
    } catch (error) {
      return null;
    }
  }

  // Enhanced color cells with context awareness
  async smartColorCells(
    target: string,
    color: string,
    context?: SheetContext
  ): Promise<string> {
    try {
      // First try to detect range with context
      let range = await this.detectRange(target, context);

      if (!range) {
        // Fallback to original parseTarget
        range = this.parseTarget(target);
      }

      if (!range) {
        throw new Error(`Could not determine range for: ${target}`);
      }

      return await this.colorCells(range, color);
    } catch (error) {
      throw new Error(
        `Smart color operation failed: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // ==== FINANCIAL ANALYSIS TOOLS ====

  // Apply conditional formatting based on financial conditions
  async conditionalFormat(
    range: string,
    condition: string,
    _threshold?: number,
    format?: {
      backgroundColor?: string;
      fontColor?: string;
      bold?: boolean;
    }
  ): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      // Resolve range if it's a description
      const resolvedRange = (await this.detectRange(range)) || range;

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();
      const fRange = worksheet.getRange(resolvedRange);

      // Apply conditional formatting
      if (condition === "negative" && format?.backgroundColor) {
        const color = this.parseColor(format.backgroundColor);
        if (color && typeof fRange.setBackground === "function") {
          fRange.setBackground(color);
        }
      }

      if (format?.fontColor) {
        const color = this.parseColor(format.fontColor);
        if (color && typeof fRange.setFontColor === "function") {
          fRange.setFontColor(color);
        }
      }

      if (format?.bold && typeof fRange.setFontWeight === "function") {
        fRange.setFontWeight("bold");
      }

      // Try to add a visible confirmation by setting a cell with status
      try {
        const statusCell = "H1"; // Use a cell that's typically empty
        const statusRange = worksheet.getRange(statusCell);
        if (typeof statusRange.setValue === "function") {
          statusRange.setValue(
            `✅ ${condition} formatting applied to ${resolvedRange}`
          );
        }
      } catch (statusError) {
        console.log("Could not set status cell:", statusError);
      }

      return `Applied conditional formatting (${condition}) to range ${resolvedRange}`;
    } catch (error) {
      throw new Error(
        `Conditional formatting failed: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Format cells as currency using Univer's proper API
  async formatCurrency(range: string, currency: string, decimals: number = 2): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const workbook = this.univerInstance.getActiveWorkbook();
      const worksheet = workbook.getActiveSheet();
      const fRange = worksheet.getRange(range);

      // Currency format patterns based on Univer's number format system
      const currencyFormats: { [key: string]: string } = {
        USD: `$#,##0.${"0".repeat(decimals)}`,
        EUR: `€#,##0.${"0".repeat(decimals)}`,
        GBP: `£#,##0.${"0".repeat(decimals)}`,
        JPY: `¥#,##0`,
        CAD: `C$#,##0.${"0".repeat(decimals)}`,
        AUD: `A$#,##0.${"0".repeat(decimals)}`,
      };

      const formatPattern = currencyFormats[currency.toUpperCase()] || `${currency}#,##0.${"0".repeat(decimals)}`;

      // Use Univer's setNumberFormat method
      fRange.setNumberFormat(formatPattern);

      return `Applied ${currency} currency formatting to ${range}`;
    } catch (error) {
      throw new Error(`Currency formatting failed: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }

  // Create subtotals for financial data
  async createSubtotals(
    dataRange: string,
    totalRow?: string,
    columns?: string[]
  ): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      const context = await this.getSheetContext();
      const resolvedRange =
        (await this.detectRange(dataRange, context)) || dataRange;

      // Parse the range to understand the data structure
      const rangeInfo = this.parseRange(resolvedRange);
      if (!rangeInfo) {
        throw new Error("Invalid data range format");
      }

      const results: string[] = [];

      // If no specific columns provided, try to detect numerical columns from context
      let columnsToSum = columns;
      if (!columnsToSum && context.tables.length > 0) {
        columnsToSum = this.detectNumericalColumns(context.tables[0]);
      }

      if (columnsToSum) {
        for (const col of columnsToSum) {
          // Calculate where to place the total
          const totalCell = totalRow || `${col}${rangeInfo.endRow + 1}`;
          const sumRange = `${col}${rangeInfo.startRow + 1}:${col}${
            rangeInfo.endRow
          }`;

          await this.setFormula(totalCell, `SUM(${sumRange})`);
          results.push(`Total for column ${col} in cell ${totalCell}`);
        }
      }

      return `Created subtotals: ${results.join(", ")}`;
    } catch (error) {
      throw new Error(
        `Subtotals creation failed: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Perform financial analysis calculations
  async financialAnalysis(
    analysisType: string,
    dataRange: string,
    comparisonRange?: string,
    periods?: number,
    outputCell?: string
  ): Promise<string> {
    try {
      if (!this.isReady()) {
        throw new Error("Univer instance not ready");
      }

      let formula = "";
      let description = "";

      switch (analysisType.toLowerCase()) {
        case "variance":
          if (comparisonRange) {
            formula = `=AVERAGE(${dataRange}) - AVERAGE(${comparisonRange})`;
            description = `variance between ${dataRange} and ${comparisonRange}`;
          }
          break;
        case "percentage_change":
          formula = `=(AVERAGE(${dataRange}) - AVERAGE(${
            comparisonRange || dataRange
          })) / AVERAGE(${comparisonRange || dataRange}) * 100`;
          description = "percentage change";
          break;
        case "moving_average":
          const periodCount = periods || 3;
          formula = `=AVERAGE(OFFSET(${dataRange}, -${
            periodCount - 1
          }, 0, ${periodCount}, 1))`;
          description = `${periodCount}-period moving average`;
          break;
        default:
          throw new Error(`Analysis type '${analysisType}' not supported`);
      }

      if (outputCell && formula) {
        await this.setFormula(outputCell, formula);
        return `Calculated ${description} in cell ${outputCell}`;
      }

      return `Analysis formula created: ${formula}`;
    } catch (error) {
      throw new Error(
        `Financial analysis failed: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }

  // Helper: Detect numerical columns from table data
  private detectNumericalColumns(table: any): string[] {
    const numericalCols: string[] = [];

    if (table.dataPreview && table.dataPreview.length > 0) {
      const firstRow = table.dataPreview[0];
      firstRow.forEach((value: string, index: number) => {
        // Check if the value looks like a number (including currency)
        if (/^[£$€¥]?[\d,.-]+$/.test(value.toString().trim())) {
          numericalCols.push(this.numberToColumn(index + 1));
        }
      });
    }

    return numericalCols;
  }

  // Enhanced range detection for financial terminology
  async detectFinancialRange(
    description: string,
    context?: SheetContext
  ): Promise<string | null> {
    const lowerDesc = description.toLowerCase();

    // Financial-specific range detection
    if (context && context.tables.length > 0) {
      const table = context.tables[0];

      // Look for amount/currency columns
      if (
        lowerDesc.includes("amount") ||
        lowerDesc.includes("value") ||
        lowerDesc.includes("total")
      ) {
        const amountCol = table.headers.findIndex((h) =>
          /amount|value|total|price|cost|revenue|expense|balance/.test(
            h.toLowerCase()
          )
        );
        if (amountCol >= 0) {
          const colLetter = this.numberToColumn(amountCol + 1);
          return `${colLetter}:${colLetter}`;
        }
      }

      // Look for currency-specific formatting
      if (
        lowerDesc.includes("sterling") ||
        lowerDesc.includes("pound") ||
        lowerDesc.includes("gbp")
      ) {
        return this.findCurrencyColumn(table, ["£", "GBP"]);
      }

      if (lowerDesc.includes("dollar") || lowerDesc.includes("usd")) {
        return this.findCurrencyColumn(table, ["$", "USD"]);
      }

      if (lowerDesc.includes("euro") || lowerDesc.includes("eur")) {
        return this.findCurrencyColumn(table, ["€", "EUR"]);
      }
    }

    return null;
  }

  // Helper to find columns containing specific currency indicators
  private findCurrencyColumn(table: any, indicators: string[]): string | null {
    if (!table.dataPreview || table.dataPreview.length === 0) return null;

    for (let colIndex = 0; colIndex < table.headers.length; colIndex++) {
      const header = table.headers[colIndex];
      const sampleData = table.dataPreview
        .map((row: string[]) => row[colIndex] || "")
        .join(" ");

      // Check if header or sample data contains currency indicators
      if (
        indicators.some(
          (indicator) =>
            header.includes(indicator) || sampleData.includes(indicator)
        )
      ) {
        return `${this.numberToColumn(colIndex + 1)}:${this.numberToColumn(
          colIndex + 1
        )}`;
      }
    }

    return null;
  }
}
