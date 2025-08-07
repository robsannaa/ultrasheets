// types/univer.ts
// TODO: Replace with official Univer types when available

export interface IWorkbookData {
  id: string;
  name: string;
  appVersion: string;
  locale: string;
  styles?: Record<string, any>;
  sheetOrder: string[];
  sheets: Record<string, IWorksheetData>;
  resources?: Record<string, any>;
}

export interface IWorksheetData {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cellData?: ICellMatrix;
  mergeData?: Array<{
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
  }>;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
}

export interface ICellMatrix {
  [row: number]: {
    [col: number]: ICellData;
  };
}

export interface ICellData {
  v?: string | number | boolean; // Cell value
  f?: string; // Formula
  t?: CellType; // Cell type
  s?: string | IStyleData; // Style
  p?: IRichTextData; // Rich text
  si?: string; // Formula ID
  custom?: Record<string, any>; // Custom metadata
}

export enum CellType {
  STRING = 1,
  NUMBER = 2,
  BOOLEAN = 3,
  FORCE_STRING = 4,
}

export interface IStyleData {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  backgroundColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
}

export interface IRichTextData {
  body?: {
    dataStream: string;
    textRuns?: Array<{
      st: number;
      ed: number;
      ts?: IStyleData;
    }>;
  };
}

// Facade API interfaces (approximated)
export interface IUniverAPI {
  getActiveWorkbook(): IFacadeWorkbook | null;
  createWorkbook(data?: Partial<IWorkbookData>): IFacadeWorkbook;
  onBeforeCommandExecute?(callback: (command: any) => void): void;
  // Add more methods as needed
}

export interface IFacadeWorkbook {
  getActiveSheet(): IFacadeWorksheet | null;
  getSnapshot(): IWorkbookData;
  create?(data?: Partial<IWorksheetData>): IFacadeWorksheet;
  deleteSheet?(sheetId: string): boolean;
  setActiveSheet?(sheetId: string): boolean;
}

export interface IFacadeWorksheet {
  getRange(range: string): IFacadeRange;
  getSnapshot(): IWorksheetData;
  // These methods may not exist in actual Univer Facade API
  insertRow?(index: number): boolean;
  deleteRow?(index: number): boolean;
  insertColumn?(index: number): boolean;
  deleteColumn?(index: number): boolean;
}

export interface IFacadeRange {
  setValue(value: string | number | boolean): void;
  setBackground(color: string): void;
  setFontWeight?(weight: 'normal' | 'bold'): void;
  setFontStyle?(style: 'normal' | 'italic'): void;
  setFontSize?(size: number): void;
  setFontColor?(color: string): void;
  // Add more formatting methods as they become available
}