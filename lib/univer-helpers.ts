export type FormatCellsAction = {
  range: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string;
  backgroundColor?: string;
  underline?: boolean;
  numberFormat?: string;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  textRotation?: number;
  textWrap?: "overflow" | "truncate" | "wrap";
};

/**
 * Apply common cell formatting via Univer Facade API with safe feature detection.
 * Keeps implementation minimal and high-level; silently skips unsupported operations.
 */
export async function applyFormatting(
  univerAPI: any,
  action: FormatCellsAction
): Promise<void> {
  if (!univerAPI) throw new Error("Univer API not available");
  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) throw new Error("No active workbook available");
  const worksheet = workbook.getActiveSheet();

  // Use A1 range string directly when available
  const range = worksheet.getRange(action.range);

  // Basic text styles
  if (action.bold !== undefined) {
    range.setFontWeight(action.bold ? "bold" : "normal");
  }
  if (action.italic !== undefined) {
    range.setFontStyle(action.italic ? "italic" : "normal");
  }
  if (action.fontSize !== undefined) {
    range.setFontSize(action.fontSize);
  }
  if (action.fontColor !== undefined) {
    range.setFontColor(action.fontColor);
  }
  if (action.backgroundColor !== undefined) {
    range.setBackgroundColor(action.backgroundColor);
  }
  if (action.numberFormat !== undefined) {
    range.setNumberFormat(action.numberFormat);
  }

  // Optional features (silently no-op if not supported by current facade)
  const anyRange = range as any;
  if (
    action.underline !== undefined &&
    typeof anyRange.setUnderline === "function"
  ) {
    anyRange.setUnderline(!!action.underline);
  }
  // Robust horizontal alignment across differing Facade APIs
  if (action.textAlign) {
    const horizontal = action.textAlign; // 'left' | 'center' | 'right'
    const candidates: Array<[string, any[]]> = [
      ["setTextAlign", [horizontal]],
      ["setHorizontalAlign", [horizontal]],
      ["setHorizontalAlignment", [horizontal]],
      ["setAlign", ["horizontal", horizontal]],
      ["setAlignment", [{ horizontal }]],
      ["setStyle", [{ textAlign: horizontal }]],
    ];
    for (const [method, args] of candidates) {
      if (typeof (anyRange as any)[method] === "function") {
        try {
          (anyRange as any)[method](...args);
          break;
        } catch {}
      }
    }
  }

  // Robust vertical alignment (normalize 'middle' to 'center' for APIs expecting 'center')
  if (action.verticalAlign) {
    const vRaw = action.verticalAlign;
    const vertical = (vRaw === "middle" ? "center" : vRaw) as
      | "top"
      | "center"
      | "bottom";
    const candidates: Array<[string, any[]]> = [
      ["setVerticalAlign", [vertical]],
      ["setVerticalAlignment", [vertical]],
      ["setAlign", ["vertical", vertical]],
      ["setAlignment", [{ vertical }]],
      ["setStyle", [{ verticalAlign: vertical }]],
    ];
    for (const [method, args] of candidates) {
      if (typeof (anyRange as any)[method] === "function") {
        try {
          (anyRange as any)[method](...args);
          break;
        } catch {}
      }
    }
  }
  if (action.textWrap && typeof anyRange.setTextWrap === "function") {
    anyRange.setTextWrap(action.textWrap);
  }
  if (
    action.textRotation !== undefined &&
    typeof anyRange.setTextRotation === "function"
  ) {
    anyRange.setTextRotation(action.textRotation);
  }
}
