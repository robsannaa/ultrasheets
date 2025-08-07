/**
 * Comprehensive Spreadsheet Service
 * Provides complete LLM control over all Univer spreadsheet operations
 */

import { UniverService } from "./univerService";

interface ComprehensiveSheetContext {
  // Multi-sheet awareness
  workbooks: Array<{
    id: string;
    name: string;
    activeSheetId: string;
    sheets: Array<{
      id: string;
      name: string;
      isActive: boolean;
      isHidden: boolean;
      tabColor?: string;
      protection?: any;
    }>;
  }>;
  
  // Current sheet detailed analysis
  currentSheet: {
    id: string;
    name: string;
    dimensions: {
      totalRows: number;
      totalColumns: number;
      usedRange: string;
      frozenRows: number;
      frozenColumns: number;
    };
    
    // Multiple tables detection
    tables: Array<{
      id: string;
      range: string;
      headers: string[];
      dataTypes: string[];
      recordCount: number;
      hasFormulas: boolean;
      sampleData: any[][];
    }>;
    
    // Comprehensive formula analysis
    formulas: Array<{
      cell: string;
      formula: string;
      result: any;
      dependencies: string[];
      isDynamic: boolean;
    }>;
    
    // Formatting and styling
    formatting: {
      conditionalFormats: Array<{
        range: string;
        rule: string;
        style: any;
      }>;
      namedRanges: Array<{
        name: string;
        range: string;
        scope: 'workbook' | 'sheet';
      }>;
      mergedCells: Array<{
        range: string;
      }>;
      customNumberFormats: Array<{
        range: string;
        format: string;
      }>;
    };
    
    // Data validation
    validationRules: Array<{
      range: string;
      type: string;
      criteria: any;
    }>;
    
    // Charts and objects
    charts: Array<{
      id: string;
      type: string;
      dataRange: string;
      position: any;
    }>;
    
    // Comments and notes
    comments: Array<{
      cell: string;
      author: string;
      text: string;
    }>;
  };
}

export class ComprehensiveSpreadsheetService {
  private univerService: UniverService;

  constructor() {
    this.univerService = UniverService.getInstance();
  }

  /**
   * Get comprehensive sheet context including all structures
   */
  async getComprehensiveContext(): Promise<ComprehensiveSheetContext> {
    try {
      // Get basic context first
      const basicContext = await this.univerService.getSheetContext();
      
      // Get workbook structure
      const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
      if (!workbook) {
        throw new Error('No active workbook');
      }

      // Analyze all sheets in workbook
      const workbookData = workbook.save();
      const sheets = Object.values(workbookData.sheets || {}).map((sheet: any) => ({
        id: sheet.id,
        name: sheet.name,
        isActive: sheet.id === workbookData.sheetOrder?.[0],
        isHidden: sheet.hidden === 1,
        tabColor: sheet.tabColor,
        protection: sheet.protection
      }));

      // Get current sheet detailed analysis
      const currentSheet = await this.analyzeCurrentSheetComprehensively(workbook, basicContext);

      return {
        workbooks: [{
          id: workbookData.id || 'default',
          name: workbookData.name || 'Unnamed Workbook',
          activeSheetId: workbookData.sheetOrder?.[0] || '',
          sheets
        }],
        currentSheet
      };

    } catch (error) {
      console.error('Error getting comprehensive context:', error);
      // Return minimal context to prevent total failure
      return this.getMinimalContext();
    }
  }

  /**
   * Analyze current sheet comprehensively
   */
  private async analyzeCurrentSheetComprehensively(workbook: any, basicContext: any): Promise<any> {
    const worksheet = workbook.getActiveSheet();
    const worksheetData = workbook.save().sheets?.[workbook.save().sheetOrder?.[0]];
    
    return {
      id: worksheetData?.id || 'sheet1',
      name: worksheetData?.name || 'Sheet1',
      dimensions: {
        totalRows: worksheetData?.rowCount || 1000,
        totalColumns: worksheetData?.columnCount || 26,
        usedRange: basicContext.dataRange || 'A1:A1',
        frozenRows: worksheetData?.freeze?.ySplit || 0,
        frozenColumns: worksheetData?.freeze?.xSplit || 0,
      },
      
      // Enhanced table detection
      tables: await this.detectAllTables(worksheetData),
      
      // Comprehensive formula analysis
      formulas: await this.analyzeAllFormulas(worksheetData),
      
      // Formatting analysis
      formatting: await this.analyzeFormatting(worksheetData),
      
      // Data validation
      validationRules: await this.analyzeDataValidation(worksheetData),
      
      // Charts and objects (placeholder - would need chart plugin)
      charts: [],
      
      // Comments (placeholder - would need comment plugin)
      comments: []
    };
  }

  /**
   * Detect ALL tables in the sheet, not just the first one
   */
  private async detectAllTables(worksheetData: any): Promise<any[]> {
    const tables = [];
    const cellMatrix = worksheetData?.cellData || {};
    
    // Look for table patterns throughout the sheet
    const potentialTableStarts = [];
    
    // Find rows that look like headers (text in consecutive columns)
    for (const rowIndex in cellMatrix) {
      const row = parseInt(rowIndex);
      const rowData = cellMatrix[rowIndex];
      let consecutiveTextCells = 0;
      let headerCandidates = [];
      
      for (let col = 0; col < 20; col++) { // Check first 20 columns
        const cell = rowData?.[col];
        if (cell && cell.v && typeof cell.v === 'string' && !cell.f) {
          consecutiveTextCells++;
          headerCandidates.push(cell.v);
        } else if (consecutiveTextCells > 0) {
          break; // End of header row
        }
      }
      
      if (consecutiveTextCells >= 2) {
        potentialTableStarts.push({
          row,
          headers: headerCandidates,
          columnCount: consecutiveTextCells
        });
      }
    }
    
    // Analyze each potential table
    for (const tableStart of potentialTableStarts) {
      const tableData = this.analyzeTableData(cellMatrix, tableStart);
      if (tableData.recordCount > 0) {
        tables.push({
          id: `table_${tableStart.row}`,
          range: tableData.range,
          headers: tableStart.headers,
          dataTypes: tableData.dataTypes,
          recordCount: tableData.recordCount,
          hasFormulas: tableData.hasFormulas,
          sampleData: tableData.sampleData
        });
      }
    }
    
    return tables;
  }

  /**
   * Analyze table data structure
   */
  private analyzeTableData(cellMatrix: any, tableStart: any): any {
    const startRow = tableStart.row + 1; // Data starts after header
    const columnCount = tableStart.columnCount;
    const sampleData = [];
    const dataTypes = [];
    let recordCount = 0;
    let hasFormulas = false;
    
    // Analyze up to 100 rows of data
    for (let row = startRow; row < startRow + 100; row++) {
      const rowData = cellMatrix[row];
      if (!rowData) break;
      
      const record = [];
      let hasDataInRow = false;
      
      for (let col = 0; col < columnCount; col++) {
        const cell = rowData[col];
        if (cell) {
          if (cell.f) hasFormulas = true;
          record.push(cell.v || cell.f);
          hasDataInRow = true;
        } else {
          record.push('');
        }
      }
      
      if (hasDataInRow) {
        recordCount++;
        if (sampleData.length < 5) {
          sampleData.push(record);
        }
      } else {
        break; // Empty row indicates end of table
      }
    }
    
    // Determine data types for each column
    if (sampleData.length > 0) {
      for (let col = 0; col < columnCount; col++) {
        const columnValues = sampleData.map(row => row[col]).filter(v => v !== '');
        dataTypes.push(this.detectColumnDataType(columnValues));
      }
    }
    
    const endRow = startRow + recordCount - 1;
    const endCol = String.fromCharCode(65 + columnCount - 1);
    const range = `A${tableStart.row + 1}:${endCol}${endRow + 1}`;
    
    return {
      range,
      dataTypes,
      recordCount,
      hasFormulas,
      sampleData
    };
  }

  /**
   * Detect column data type from sample values
   */
  private detectColumnDataType(values: any[]): string {
    if (values.length === 0) return 'empty';
    
    const numericCount = values.filter(v => typeof v === 'number' || !isNaN(Number(v))).length;
    const dateCount = values.filter(v => this.isDateLike(v)).length;
    const currencyCount = values.filter(v => this.isCurrencyLike(v)).length;
    
    const total = values.length;
    
    if (currencyCount / total > 0.7) return 'currency';
    if (numericCount / total > 0.7) return 'number';
    if (dateCount / total > 0.7) return 'date';
    
    return 'text';
  }

  private isDateLike(value: any): boolean {
    if (typeof value === 'string') {
      return /\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}-\d{2}-\d{4}/.test(value);
    }
    return false;
  }

  private isCurrencyLike(value: any): boolean {
    if (typeof value === 'string') {
      return /[$£€¥]/.test(value) || /^\d+\.\d{2}$/.test(value);
    }
    return false;
  }

  /**
   * Analyze ALL formulas in the sheet
   */
  private async analyzeAllFormulas(worksheetData: any): Promise<any[]> {
    const formulas = [];
    const cellMatrix = worksheetData?.cellData || {};
    
    for (const rowIndex in cellMatrix) {
      const row = parseInt(rowIndex);
      const rowData = cellMatrix[rowIndex];
      
      for (const colIndex in rowData) {
        const col = parseInt(colIndex);
        const cell = rowData[colIndex];
        
        if (cell && cell.f) {
          const cellRef = `${String.fromCharCode(65 + col)}${row + 1}`;
          formulas.push({
            cell: cellRef,
            formula: cell.f,
            result: cell.v,
            dependencies: this.extractFormulaDependencies(cell.f),
            isDynamic: this.isFormulaExternallyReferenced(cell.f)
          });
        }
      }
    }
    
    return formulas;
  }

  private extractFormulaDependencies(formula: string): string[] {
    const dependencies = [];
    // Extract cell references like A1, B2:B10, etc.
    const cellRefs = formula.match(/[A-Z]+\d+(?::[A-Z]+\d+)?/g);
    if (cellRefs) {
      dependencies.push(...cellRefs);
    }
    return dependencies;
  }

  private isFormulaExternallyReferenced(formula: string): boolean {
    // Check if formula references other sheets or external sources
    return formula.includes('!') || formula.includes('INDIRECT') || formula.includes('OFFSET');
  }

  /**
   * Analyze formatting and styling
   */
  private async analyzeFormatting(worksheetData: any): Promise<any> {
    return {
      conditionalFormats: [], // Would need conditional formatting plugin
      namedRanges: [], // Would need defined names plugin  
      mergedCells: worksheetData?.mergeData || [],
      customNumberFormats: [] // Would analyze number formats
    };
  }

  /**
   * Analyze data validation rules
   */
  private async analyzeDataValidation(worksheetData: any): Promise<any[]> {
    // Would need data validation plugin to implement
    return [];
  }

  /**
   * Minimal context fallback
   */
  private getMinimalContext(): ComprehensiveSheetContext {
    return {
      workbooks: [{
        id: 'default',
        name: 'Workbook',
        activeSheetId: 'sheet1',
        sheets: [{
          id: 'sheet1',
          name: 'Sheet1',
          isActive: true,
          isHidden: false
        }]
      }],
      currentSheet: {
        id: 'sheet1',
        name: 'Sheet1',
        dimensions: {
          totalRows: 1000,
          totalColumns: 26,
          usedRange: 'A1:A1',
          frozenRows: 0,
          frozenColumns: 0
        },
        tables: [],
        formulas: [],
        formatting: {
          conditionalFormats: [],
          namedRanges: [],
          mergedCells: [],
          customNumberFormats: []
        },
        validationRules: [],
        charts: [],
        comments: []
      }
    };
  }
}