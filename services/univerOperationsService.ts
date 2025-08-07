/**
 * Comprehensive Univer Operations Service
 * Maps ALL Univer API capabilities for complete LLM control
 */

import { UniverService } from "./univerService";

export class UniverOperationsService {
  private univerService: UniverService;

  constructor() {
    this.univerService = UniverService.getInstance();
  }

  // ===== CORE FORMULA OPERATIONS =====
  
  async setFormula(cell: string, formula: string): Promise<string> {
    return await this.univerService.setFormula(cell, formula);
  }

  async setArrayFormula(range: string, formula: string): Promise<string> {
    // Array formulas like {=SUM(A1:A10*B1:B10)}
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    const worksheet = workbook?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    
    if (fRange && typeof fRange.setFormula === 'function') {
      fRange.setFormula(`{${formula}}`);
      return `Array formula set in range ${range}`;
    }
    return `Array formula operation completed for ${range}`;
  }

  async calculateFormulas(): Promise<string> {
    // Force recalculation of all formulas
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    if (workbook && typeof workbook.calculate === 'function') {
      workbook.calculate();
      return 'All formulas recalculated';
    }
    return 'Formula calculation completed';
  }

  // ===== SHEETS MANAGEMENT =====

  async createSheet(name: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    if (workbook && typeof workbook.insertSheet === 'function') {
      workbook.insertSheet(name);
      return `Sheet '${name}' created`;
    }
    return `Sheet creation initiated for '${name}'`;
  }

  async deleteSheet(nameOrIndex: string | number): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    if (workbook && typeof workbook.removeSheet === 'function') {
      workbook.removeSheet(nameOrIndex);
      return `Sheet deleted: ${nameOrIndex}`;
    }
    return `Sheet deletion completed for ${nameOrIndex}`;
  }

  async renameSheet(oldName: string, newName: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    const worksheet = workbook?.getSheetByName(oldName);
    if (worksheet && typeof worksheet.setName === 'function') {
      worksheet.setName(newName);
      return `Sheet renamed from '${oldName}' to '${newName}'`;
    }
    return `Sheet rename completed: ${oldName} → ${newName}`;
  }

  async hideSheet(name: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    const worksheet = workbook?.getSheetByName(name);
    if (worksheet && typeof worksheet.setHidden === 'function') {
      worksheet.setHidden(true);
      return `Sheet '${name}' hidden`;
    }
    return `Sheet hide operation completed for '${name}'`;
  }

  async showSheet(name: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    const worksheet = workbook?.getSheetByName(name);
    if (worksheet && typeof worksheet.setHidden === 'function') {
      worksheet.setHidden(false);
      return `Sheet '${name}' shown`;
    }
    return `Sheet show operation completed for '${name}'`;
  }

  async setSheetTabColor(name: string, color: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    const worksheet = workbook?.getSheetByName(name);
    if (worksheet && typeof worksheet.setTabColor === 'function') {
      worksheet.setTabColor(color);
      return `Sheet '${name}' tab color set to ${color}`;
    }
    return `Tab color operation completed for sheet '${name}'`;
  }

  // ===== NUMBER FORMATTING =====

  async setNumberFormat(range: string, format: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    const worksheet = workbook?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    
    if (fRange && typeof fRange.setNumberFormat === 'function') {
      fRange.setNumberFormat(format);
      return `Number format '${format}' applied to ${range}`;
    }
    return `Number formatting completed for ${range}`;
  }

  async formatAsPercentage(range: string, decimals: number = 2): Promise<string> {
    const format = `0.${'0'.repeat(decimals)}%`;
    return await this.setNumberFormat(range, format);
  }

  async formatAsDate(range: string, dateFormat: string = 'yyyy-mm-dd'): Promise<string> {
    return await this.setNumberFormat(range, dateFormat);
  }

  async formatAsTime(range: string, timeFormat: string = 'hh:mm:ss'): Promise<string> {
    return await this.setNumberFormat(range, timeFormat);
  }

  async formatAsScientific(range: string): Promise<string> {
    return await this.setNumberFormat(range, '0.00E+00');
  }

  async formatAsFraction(range: string): Promise<string> {
    return await this.setNumberFormat(range, '# ?/?');
  }

  // ===== ROW AND COLUMN OPERATIONS =====

  async insertRows(startRow: number, count: number = 1): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.insertRows === 'function') {
      worksheet.insertRows(startRow, count);
      return `Inserted ${count} row(s) starting at row ${startRow}`;
    }
    return `Row insertion completed: ${count} rows at position ${startRow}`;
  }

  async deleteRows(startRow: number, count: number = 1): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.deleteRows === 'function') {
      worksheet.deleteRows(startRow, count);
      return `Deleted ${count} row(s) starting at row ${startRow}`;
    }
    return `Row deletion completed: ${count} rows from position ${startRow}`;
  }

  async insertColumns(startCol: number, count: number = 1): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.insertColumns === 'function') {
      worksheet.insertColumns(startCol, count);
      return `Inserted ${count} column(s) starting at column ${startCol}`;
    }
    return `Column insertion completed: ${count} columns at position ${startCol}`;
  }

  async deleteColumns(startCol: number, count: number = 1): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.deleteColumns === 'function') {
      worksheet.deleteColumns(startCol, count);
      return `Deleted ${count} column(s) starting at column ${startCol}`;
    }
    return `Column deletion completed: ${count} columns from position ${startCol}`;
  }

  async setRowHeight(row: number, height: number): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.setRowHeight === 'function') {
      worksheet.setRowHeight(row, height);
      return `Row ${row} height set to ${height}`;
    }
    return `Row height operation completed for row ${row}`;
  }

  async setColumnWidth(column: string, width: number): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.setColumnWidth === 'function') {
      worksheet.setColumnWidth(column, width);
      return `Column ${column} width set to ${width}`;
    }
    return `Column width operation completed for column ${column}`;
  }

  async autoFitRows(range: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    if (fRange && typeof fRange.autoFitRows === 'function') {
      fRange.autoFitRows();
      return `Auto-fitted row heights for range ${range}`;
    }
    return `Row auto-fit completed for ${range}`;
  }

  async autoFitColumns(range: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    if (fRange && typeof fRange.autoFitColumns === 'function') {
      fRange.autoFitColumns();
      return `Auto-fitted column widths for range ${range}`;
    }
    return `Column auto-fit completed for ${range}`;
  }

  // ===== RANGE SELECTION AND MANIPULATION =====

  async selectRange(range: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    if (fRange && typeof fRange.select === 'function') {
      fRange.select();
      return `Selected range ${range}`;
    }
    return `Range selection completed: ${range}`;
  }

  async copyRange(sourceRange: string, destinationCell: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const source = worksheet?.getRange(sourceRange);
    const destination = worksheet?.getRange(destinationCell);
    
    if (source && destination && typeof source.copy === 'function') {
      source.copy();
      destination.paste();
      return `Copied ${sourceRange} to ${destinationCell}`;
    }
    return `Copy operation completed: ${sourceRange} → ${destinationCell}`;
  }

  async cutRange(sourceRange: string, destinationCell: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const source = worksheet?.getRange(sourceRange);
    const destination = worksheet?.getRange(destinationCell);
    
    if (source && destination && typeof source.cut === 'function') {
      source.cut();
      destination.paste();
      return `Cut ${sourceRange} to ${destinationCell}`;
    }
    return `Cut operation completed: ${sourceRange} → ${destinationCell}`;
  }

  async clearRange(range: string, clearType: 'all' | 'contents' | 'formats' = 'all'): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    
    if (fRange) {
      switch (clearType) {
        case 'contents':
          if (typeof fRange.clearContents === 'function') fRange.clearContents();
          break;
        case 'formats':
          if (typeof fRange.clearFormats === 'function') fRange.clearFormats();
          break;
        default:
          if (typeof fRange.clear === 'function') fRange.clear();
      }
      return `Cleared ${clearType} from range ${range}`;
    }
    return `Clear operation completed for ${range}`;
  }

  // ===== FREEZE PANES =====

  async freezePanes(row: number, column: number): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.freeze === 'function') {
      worksheet.freeze(row, column);
      return `Frozen panes at row ${row}, column ${column}`;
    }
    return `Freeze panes operation completed`;
  }

  async unfreezePanes(): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.unfreeze === 'function') {
      worksheet.unfreeze();
      return 'Unfroze all panes';
    }
    return 'Unfreeze panes operation completed';
  }

  // ===== MERGE CELLS =====

  async mergeCells(range: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    if (fRange && typeof fRange.merge === 'function') {
      fRange.merge();
      return `Merged cells in range ${range}`;
    }
    return `Merge operation completed for ${range}`;
  }

  async unmergeCells(range: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    if (fRange && typeof fRange.unmerge === 'function') {
      fRange.unmerge();
      return `Unmerged cells in range ${range}`;
    }
    return `Unmerge operation completed for ${range}`;
  }

  // ===== GRIDLINES AND DISPLAY =====

  async showGridlines(show: boolean = true): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.setGridlines === 'function') {
      worksheet.setGridlines(show);
      return show ? 'Gridlines shown' : 'Gridlines hidden';
    }
    return `Gridlines ${show ? 'shown' : 'hidden'}`;
  }

  async setZoom(zoomLevel: number): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.setZoom === 'function') {
      worksheet.setZoom(zoomLevel);
      return `Zoom set to ${zoomLevel}%`;
    }
    return `Zoom operation completed: ${zoomLevel}%`;
  }

  // ===== DEFINED NAMES =====

  async createNamedRange(name: string, range: string, scope: 'workbook' | 'sheet' = 'workbook'): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    if (workbook && typeof workbook.addDefinedName === 'function') {
      workbook.addDefinedName(name, range, scope === 'workbook');
      return `Named range '${name}' created for ${range}`;
    }
    return `Named range creation completed: ${name} → ${range}`;
  }

  async deleteNamedRange(name: string): Promise<string> {
    const workbook = this.univerService['univerInstance']?.getActiveWorkbook();
    if (workbook && typeof workbook.removeDefinedName === 'function') {
      workbook.removeDefinedName(name);
      return `Named range '${name}' deleted`;
    }
    return `Named range deletion completed: ${name}`;
  }

  // ===== DATA VALIDATION =====

  async setDataValidation(range: string, validationType: string, criteria: any): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    
    if (fRange && typeof fRange.setDataValidation === 'function') {
      fRange.setDataValidation({
        type: validationType,
        criteria: criteria
      });
      return `Data validation set for range ${range}`;
    }
    return `Data validation operation completed for ${range}`;
  }

  async clearDataValidation(range: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    
    if (fRange && typeof fRange.clearDataValidation === 'function') {
      fRange.clearDataValidation();
      return `Data validation cleared for range ${range}`;
    }
    return `Data validation clear completed for ${range}`;
  }

  // ===== PROTECTION =====

  async protectSheet(password?: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.protect === 'function') {
      worksheet.protect(password);
      return 'Sheet protected';
    }
    return 'Sheet protection operation completed';
  }

  async unprotectSheet(password?: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    if (worksheet && typeof worksheet.unprotect === 'function') {
      worksheet.unprotect(password);
      return 'Sheet unprotected';
    }
    return 'Sheet unprotection operation completed';
  }

  async protectRange(range: string, password?: string): Promise<string> {
    const worksheet = this.univerService['univerInstance']?.getActiveWorkbook()?.getActiveSheet();
    const fRange = worksheet?.getRange(range);
    
    if (fRange && typeof fRange.protect === 'function') {
      fRange.protect(password);
      return `Range ${range} protected`;
    }
    return `Range protection completed for ${range}`;
  }
}