import { GetWorkbookSnapshotTool } from "../modern-tools";
import { FindCellTool } from "../additional-tools";
import { SwitchSheetTool, CreateSheetTool } from "../migrated-tools";

export const NAVIGATION_TOOLS = [
  GetWorkbookSnapshotTool,
  FindCellTool,
  SwitchSheetTool,
  CreateSheetTool,
];
