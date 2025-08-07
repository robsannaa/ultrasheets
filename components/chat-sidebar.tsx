"use client";

import { useRef, useEffect } from "react";
import * as React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Send, Loader2, CheckCircle2, Wrench } from "lucide-react";
import { cn } from "@/lib/utils";
import { useChat } from "ai/react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";

function ChatMessages({
  messages,
  isLoading,
}: {
  messages: any[];
  isLoading: boolean;
}) {
  const scrollAreaRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (scrollAreaRef.current) {
      const scrollContainer = scrollAreaRef.current.querySelector(
        "[data-radix-scroll-area-viewport]"
      );
      if (scrollContainer) {
        scrollContainer.scrollTop = scrollContainer.scrollHeight;
      }
    }
  }, [messages]);

  // Merge adjacent assistant messages with identical text to avoid repeated blocks
  const mergedMessages = React.useMemo(() => {
    const result: any[] = [];
    const normalize = (s: string) => (s || "").replace(/\s+/g, " ").trim();
    const getText = (m: any) => {
      if (Array.isArray(m.parts)) {
        try {
          return m.parts
            .filter((p: any) => p.type === "text" && typeof p.text === "string")
            .map((p: any) => p.text)
            .join("");
        } catch {
          return m.content || "";
        }
      }
      return m.content || "";
    };

    for (const m of messages) {
      if (
        result.length > 0 &&
        m.role === "assistant" &&
        result[result.length - 1].role === "assistant"
      ) {
        const prev = result[result.length - 1];
        if (normalize(getText(prev)) === normalize(getText(m))) {
          // merge tool parts; keep one text block
          const merged = {
            ...prev,
            parts: [...(prev.parts || []), ...(m.parts || [])],
          };
          result[result.length - 1] = merged;
          continue;
        }
      }
      result.push(m);
    }
    return result;
  }, [messages]);

  return (
    <ScrollArea className="h-full px-1 py-4" ref={scrollAreaRef}>
      <div className="space-y-4">
        {mergedMessages.map((message) => (
          <div
            key={message.id}
            className={cn(
              "flex items-start gap-3",
              message.role === "user" && "flex-row-reverse"
            )}
          >
            <Avatar className="w-8 h-8">
              <AvatarFallback>
                {message.role === "user" ? "U" : "A"}
              </AvatarFallback>
            </Avatar>
            <div
              className={cn(
                "max-w-[70%] rounded-lg p-3 text-sm",
                message.role === "user"
                  ? "bg-primary text-primary-foreground whitespace-pre-wrap"
                  : "bg-muted"
              )}
            >
              {message.parts ? (
                <>
                  {(() => {
                    try {
                      const text = message.parts
                        .filter(
                          (p: any) =>
                            p.type === "text" && typeof p.text === "string"
                        )
                        .map((p: any) => p.text)
                        .join("");
                      return text ? (
                        <CollapsibleMarkdown key="assistant-text" text={text} />
                      ) : null;
                    } catch {
                      return null;
                    }
                  })()}
                  {message.parts.map((part: any, index: number) => {
                    switch (part.type) {
                      case "text":
                        return null; // already rendered once above
                      case "tool-invocation": {
                        const callId = part.toolInvocation.toolCallId;
                        const toolName = part.toolInvocation.toolName;
                        const state = part.toolInvocation.state;
                        const args = part.toolInvocation.args;
                        const label = describeTool(toolName, args);

                        if (state === "call") {
                          if (toolName === "get_sheet_context") return null;
                          return (
                            <ToolBadge
                              key={callId}
                              label={label}
                              toolName={toolName}
                              state="call"
                            />
                          );
                        } else if (state === "result") {
                          // Only show result for important operations, hide technical details
                          const result = part.toolInvocation.result;
                          if (toolName === "get_sheet_context") {
                            return null; // Hide sheet context results - too technical
                          }

                          // Handle new error response format
                          if (
                            result &&
                            typeof result === "object" &&
                            result.error
                          ) {
                            return (
                              <div
                                key={callId}
                                className="text-red-600 text-sm"
                              >
                                ⚠️ {result.message || result.error}
                              </div>
                            );
                          }
                          return (
                            <ToolBadge
                              key={callId}
                              label={label}
                              toolName={toolName}
                              state="result"
                            />
                          );
                        }
                        return null;
                      }
                      default:
                        if (message.role === "assistant" && message.content) {
                          return (
                            <div
                              key={index}
                              className="leading-6 prose prose-sm max-w-none dark:prose-invert"
                            >
                              <ReactMarkdown remarkPlugins={[remarkGfm]}>
                                {message.content}
                              </ReactMarkdown>
                            </div>
                          );
                        }
                        return null;
                    }
                  })}
                </>
              ) : message.role === "assistant" ? (
                <div className="leading-6 prose prose-sm max-w-none dark:prose-invert">
                  <ReactMarkdown remarkPlugins={[remarkGfm]}>
                    {message.content}
                  </ReactMarkdown>
                </div>
              ) : (
                message.content
              )}
            </div>
          </div>
        ))}
        {/* Typing indicator removed: the send button already shows progress */}
      </div>
    </ScrollArea>
  );
}

function ChatInput({
  input,
  handleInputChange,
  handleSubmit,
  isLoading,
}: {
  input: string;
  handleInputChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  handleSubmit: (e: React.FormEvent<HTMLFormElement>) => void;
  isLoading: boolean;
}) {
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      const form = e.currentTarget.closest("form");
      if (form) {
        form.requestSubmit();
      }
    }
  };

  const handleFormSubmit = handleSubmit;

  return (
    <form onSubmit={handleFormSubmit} className="flex gap-2">
      <Input
        placeholder="Ask about your spreadsheet data..."
        value={input}
        onChange={handleInputChange}
        onKeyDown={handleKeyDown}
        disabled={isLoading}
        className="flex-1"
      />
      <Button type="submit" size="icon" disabled={isLoading || !input.trim()}>
        {isLoading ? (
          <Loader2 className="h-4 w-4 animate-spin" />
        ) : (
          <Send className="h-4 w-4" />
        )}
      </Button>
    </form>
  );
}

// Rich workbook context with multi-table detection and recent action log
function extractWorkbookData() {
  try {
    if (typeof window === "undefined" || !(window as any).univerAPI)
      return null;
    const w: any = window as any;
    const univerAPI = w.univerAPI;
    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook || typeof workbook.save !== "function") return null;

    const activeSheet = workbook.getActiveSheet();
    const activeSnapshot = activeSheet?.getSheet()?.getSnapshot();
    const activeName = activeSnapshot?.name;

    const wbData = workbook.save();
    const sheetOrder: string[] = wbData.sheetOrder || [];
    const sheetsData: Record<string, any> = wbData.sheets || {};

    const detectTables = (cellData: any) => {
      const tables: Array<{
        range: string;
        headers: string[];
        recordCount: number;
        numericColumns: string[];
      }> = [];
      if (!cellData) return tables;

      const rows = Object.keys(cellData)
        .map((k) => parseInt(k, 10))
        .sort((a, b) => a - b);
      const maxRow = rows.length ? rows[rows.length - 1] : 0;
      let maxCol = 0;
      for (const r of rows) {
        const cols = Object.keys(cellData[r] || {}).map((k) => parseInt(k, 10));
        if (cols.length) maxCol = Math.max(maxCol, cols[cols.length - 1]);
      }

      const headerCandidates: number[] = [];
      for (let r = 0; r <= Math.min(maxRow, 20); r++) {
        const rowData = cellData[r] || {};
        let textRun = 0;
        for (let c = 0; c <= Math.min(maxCol, 30); c++) {
          const cell = rowData[c];
          if (cell && typeof cell.v === "string" && cell.v.trim() && !cell.f)
            textRun++;
          else if (textRun > 0) break;
        }
        if (textRun >= 2) headerCandidates.push(r);
      }

      for (const headerRow of headerCandidates) {
        const rowData = cellData[headerRow] || {};
        const headers: string[] = [];
        let firstCol = -1;
        let lastCol = -1;
        for (let c = 0; c <= Math.min(maxCol, 50); c++) {
          const cell = rowData[c];
          if (cell && typeof cell.v === "string" && cell.v.trim() && !cell.f) {
            headers.push(String(cell.v).trim());
            if (firstCol === -1) firstCol = c;
            lastCol = c;
          } else if (headers.length > 0) break;
        }
        if (headers.length < 2) continue;

        let recordCount = 0;
        for (
          let r = headerRow + 1;
          r <= Math.min(headerRow + 500, maxRow);
          r++
        ) {
          const rd = cellData[r] || {};
          let hasData = false;
          for (let c = firstCol; c <= lastCol; c++) {
            const cell = rd[c];
            if (
              cell &&
              cell.v !== undefined &&
              cell.v !== null &&
              cell.v !== ""
            ) {
              hasData = true;
              break;
            }
          }
          if (hasData) recordCount++;
          else if (recordCount > 0) break;
        }

        const startColLetter = String.fromCharCode(65 + Math.max(0, firstCol));
        const endColLetter = String.fromCharCode(65 + Math.max(0, lastCol));
        const range = `${startColLetter}${headerRow + 1}:${endColLetter}${
          headerRow + 1 + recordCount
        }`;

        const numericColumns: string[] = [];
        for (let c = firstCol; c <= lastCol; c++) {
          let numericHits = 0;
          for (
            let r = headerRow + 1;
            r <= headerRow + 1 + Math.min(5, recordCount);
            r++
          ) {
            const cell = (cellData[r] || {})[c];
            if (!cell) continue;
            const v = cell.v;
            if (
              typeof v === "number" ||
              (typeof v === "string" && /^[£$€¥]?\s*[\d,.]+$/.test(v))
            )
              numericHits++;
          }
          if (numericHits >= 2)
            numericColumns.push(String.fromCharCode(65 + c));
        }

        tables.push({ range, headers, recordCount, numericColumns });
      }

      return tables;
    };

    const sheets = sheetOrder.map((sid) => {
      const s = sheetsData[sid];
      const name = s?.name || "Sheet";
      const cellData = s?.cellData || {};

      let totalCells = 0;
      for (const r in cellData) {
        for (const c in cellData[r]) {
          const cell = cellData[r][c];
          if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "")
            totalCells++;
        }
      }

      const tables = detectTables(cellData);
      const primaryHeaders = tables[0]?.headers || [];
      const dataRows = tables[0]?.recordCount || 0;

      return {
        name,
        isActive: name === activeName,
        headers: primaryHeaders,
        structure: { totalCells, dataRows },
        tables,
      };
    });

    const recentActions = Array.isArray(w.ultraActionLog)
      ? (w.ultraActionLog as any[]).slice(-10)
      : [];

    return { sheets, recentActions };
  } catch (error) {
    console.error("Error extracting workbook data:", error);
    return null;
  }
}

export function ChatSidebar() {
  const { messages, input, handleInputChange, handleSubmit, status } = useChat({
    api: "/api/chat",
    maxSteps: 5,
    initialMessages: [
      {
        id: "1",
        role: "assistant" as const,
        content:
          "Hello! I'm your advanced spreadsheet assistant. How can I help you analyze your spreadsheet data today?",
      },
    ],
    // Send workbook data with each request - this gets evaluated each time
    body: {
      workbookData: extractWorkbookData(),
    },
  });

  // Keep track of executed tool invocations to avoid duplicate runs during streaming updates
  const executedToolCallIdsRef = React.useRef<Set<string>>(new Set());

  // Listen for tool results and execute them on the frontend
  useEffect(() => {
    const executeClientSideActions = async () => {
      const lastMessage = messages[messages.length - 1];
      if (lastMessage?.role !== "assistant" || !lastMessage.parts) return;

      for (const part of lastMessage.parts) {
        if (
          part.type === "tool-invocation" &&
          part.toolInvocation.state === "result"
        ) {
          const callId: string | undefined = part.toolInvocation.toolCallId;
          // Deduplicate by toolCallId to prevent re-execution flicker
          if (callId && executedToolCallIdsRef.current.has(callId)) continue;

          const result = part.toolInvocation.result;
          if (!result || !result.clientSideAction) continue;

          const action = result.clientSideAction;
          try {
            if (
              action.type === "executeUniverTool" &&
              window.executeUniverTool
            ) {
              const execResult = await window.executeUniverTool(
                action.toolName,
                action.params
              );
              // Record recent action for richer LLM context
              try {
                const w: any = window as any;
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: action.toolName,
                  params: action.params,
                  result:
                    execResult &&
                    typeof execResult === "object" &&
                    execResult.message
                      ? execResult.message
                      : "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "formatCells") {
              await formatCells(action);
              try {
                const w: any = window as any;
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "format_cells",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "setCellValue") {
              // Direct cell value/formula set with formula recalculation
              const w: any = window as any;
              if (!(window as any).univerAPI)
                throw new Error("Univer API not available");
              const univerAPI = (window as any).univerAPI;
              const workbook = univerAPI.getActiveWorkbook();
              const worksheet = workbook.getActiveSheet();
              const cellRef: string = action.cell;
              const isFormula: boolean = !!action.formula;
              const value = action.value;
              const colLetter = cellRef.replace(/\d+/g, "");
              const rowNumber = parseInt(cellRef.replace(/\D+/g, ""), 10);
              const colIndex = colLetter.charCodeAt(0) - 65;
              const rowIndex = rowNumber - 1;
              const range = worksheet.getRange(rowIndex, colIndex, 1, 1);
              const finalValue = isFormula
                ? String(value).startsWith("=")
                  ? String(value)
                  : `=${String(value)}`
                : value;
              range.setValue(finalValue);
              try {
                const formula = univerAPI.getFormula();
                formula.executeCalculation();
              } catch {}
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "set_cell_value",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            }

            // Mark as executed only after successful run
            if (callId) executedToolCallIdsRef.current.add(callId);
          } catch (error) {
            console.error("Failed to execute client-side action:", error);
            // Do not mark executed on failure to allow retry on next render
          }
        }
      }
    };

    executeClientSideActions();
  }, [messages]);

  // Format cells function
  const formatCells = async (action: any) => {
    console.log("🚀 formatCells called with action:", action);
    try {
      if (typeof window !== "undefined" && (window as any).univerAPI) {
        const univerAPI = (window as any).univerAPI;
        console.log("🔍 UniversAPI object:", univerAPI);
        console.log(
          "🔍 UniversAPI methods:",
          Object.getOwnPropertyNames(Object.getPrototypeOf(univerAPI)).filter(
            (name) => typeof univerAPI[name] === "function"
          )
        );

        const workbook = univerAPI.getActiveWorkbook();
        console.log("🔍 Workbook object:", workbook);
        console.log(
          "🔍 Workbook methods:",
          Object.getOwnPropertyNames(Object.getPrototypeOf(workbook)).filter(
            (name) => typeof workbook[name] === "function"
          )
        );

        const worksheet = workbook.getActiveSheet();

        // Parse range (e.g., "A1:G1" or "A1")
        let startRow, startCol, endRow, endCol;

        if (action.range.includes(":")) {
          // Range format like "A1:G1"
          const [start, end] = action.range.split(":");
          startCol = start.charCodeAt(0) - 65;
          startRow = parseInt(start.slice(1)) - 1;
          endCol = end.charCodeAt(0) - 65;
          endRow = parseInt(end.slice(1)) - 1;
        } else {
          // Single cell like "A1"
          startCol = endCol = action.range.charCodeAt(0) - 65;
          startRow = endRow = parseInt(action.range.slice(1)) - 1;
        }

        // Get the range
        const range = worksheet.getRange(
          startRow,
          startCol,
          endRow - startRow + 1,
          endCol - startCol + 1
        );

        // Build comprehensive style object based on Univer documentation
        const styleUpdates: any = {};

        // Bold: bl property (0 = not bold, 1 = bold)
        if (action.bold !== undefined) {
          styleUpdates.bl = action.bold ? 1 : 0;
        }

        // Italic: it property (0 = not italic, 1 = italic)
        if (action.italic !== undefined) {
          styleUpdates.it = action.italic ? 1 : 0;
        }

        // Font Size: fs property (number in pt)
        if (action.fontSize !== undefined) {
          styleUpdates.fs = action.fontSize;
        }

        // Font Color: cl property (object with rgb)
        if (action.fontColor !== undefined) {
          styleUpdates.cl = { rgb: action.fontColor };
        }

        // Background Color: bg property (object with rgb)
        if (action.backgroundColor !== undefined) {
          styleUpdates.bg = { rgb: action.backgroundColor };
        }

        // Underline: ul property (0 = not underlined, 1 = underlined)
        if (action.underline !== undefined) {
          styleUpdates.ul = action.underline ? { s: 1 } : { s: 0 };
        }

        // Horizontal Alignment: ht property (1 = left, 2 = center, 3 = right)
        if (action.textAlign !== undefined) {
          const alignmentMap: { [key: string]: number } = {
            left: 1,
            center: 2,
            right: 3,
          };
          styleUpdates.ht = alignmentMap[action.textAlign] || 1;
        }

        // Vertical Alignment: vt property (1 = top, 2 = middle, 3 = bottom)
        if (action.verticalAlign !== undefined) {
          const verticalAlignMap: { [key: string]: number } = {
            top: 1,
            middle: 2,
            bottom: 3,
          };
          styleUpdates.vt = verticalAlignMap[action.verticalAlign] || 1;
        }

        // Text Rotation: tr property (object with angle and vertical flag)
        if (action.textRotation !== undefined) {
          styleUpdates.tr = {
            a: action.textRotation, // angle in degrees
            v: 0, // horizontal (1 for vertical)
          };
        }

        // Text Wrap: tb property (1 = overflow, 2 = truncate, 3 = wrap)
        if (action.textWrap !== undefined) {
          const wrapMap: { [key: string]: number } = {
            overflow: 1,
            truncate: 2,
            wrap: 3,
          };
          styleUpdates.tb = wrapMap[action.textWrap] || 1;
        }

        // Number Format: n property (object with pattern)
        if (action.numberFormat !== undefined) {
          styleUpdates.n = {
            pattern: action.numberFormat,
          };
        }

        // Apply formatting using available range methods
        try {
          // Basic formatting using known working methods
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

          // Text alignment - comprehensive approach
          if (action.textAlign !== undefined) {
            const alignmentValue: "left" | "center" | "right" =
              action.textAlign;
            console.log(
              `🔍 Attempting to set text alignment to: ${alignmentValue}`
            );

            // Debug: Log all available methods on the range object
            const rangeMethods = Object.getOwnPropertyNames(
              Object.getPrototypeOf(range)
            )
              .filter((name) => typeof range[name] === "function")
              .sort();
            console.log("📝 Available range methods:", rangeMethods);

            // Debug: Log alignment-related methods specifically
            const alignmentMethods = rangeMethods.filter(
              (name) =>
                name.toLowerCase().includes("align") ||
                name.toLowerCase().includes("style") ||
                name.toLowerCase().includes("format")
            );
            console.log("🎯 Alignment-related methods:", alignmentMethods);

            let alignmentApplied = false;

            // Method 1: Try setHorizontalAlignment with different variations
            const horizontalMethods = [
              "setHorizontalAlignment",
              "setTextAlign",
              "setAlign",
            ];
            for (const methodName of horizontalMethods) {
              if (typeof range[methodName] === "function") {
                try {
                  range[methodName](alignmentValue);
                  console.log(
                    `✅ Applied alignment via ${methodName}: ${alignmentValue}`
                  );
                  alignmentApplied = true;
                  break;
                } catch (error) {
                  console.warn(`⚠️ ${methodName} failed:`, error);
                }
              }
            }

            // Method 2: Try setStyle or similar methods
            if (!alignmentApplied) {
              const styleMethods = ["setStyle", "applyStyle", "updateStyle"];
              for (const methodName of styleMethods) {
                if (typeof range[methodName] === "function") {
                  try {
                    const alignmentMap: {
                      left: number;
                      center: number;
                      right: number;
                    } = { left: 1, center: 2, right: 3 };
                    const styleObj = { ht: alignmentMap[alignmentValue] || 1 };
                    range[methodName](styleObj);
                    console.log(
                      `✅ Applied alignment via ${methodName}:`,
                      styleObj
                    );
                    alignmentApplied = true;
                    break;
                  } catch (error) {
                    console.warn(`⚠️ ${methodName} failed:`, error);
                  }
                }
              }
            }

            // Method 3: Try direct worksheet manipulation
            if (!alignmentApplied) {
              try {
                console.log("🔍 Attempting worksheet direct manipulation...");
                const worksheet = workbook.getActiveSheet();

                // Check if worksheet has setCellStyle method
                if (typeof worksheet.setCellStyle === "function") {
                  const alignmentMap: {
                    left: number;
                    center: number;
                    right: number;
                  } = { left: 1, center: 2, right: 3 };
                  const styleValue = alignmentMap[alignmentValue] || 1;

                  for (let row = startRow; row <= endRow; row++) {
                    for (let col = startCol; col <= endCol; col++) {
                      worksheet.setCellStyle(row, col, { ht: styleValue });
                    }
                  }
                  console.log(
                    `✅ Applied alignment via worksheet.setCellStyle: ${alignmentValue} (ht: ${styleValue})`
                  );
                  alignmentApplied = true;
                } else {
                  console.log(
                    "📝 Available worksheet methods:",
                    Object.getOwnPropertyNames(Object.getPrototypeOf(worksheet))
                      .filter((name) => typeof worksheet[name] === "function")
                      .filter(
                        (name) =>
                          name.toLowerCase().includes("style") ||
                          name.toLowerCase().includes("cell")
                      )
                      .sort()
                  );
                }
              } catch (error) {
                console.warn("⚠️ Worksheet direct manipulation failed:", error);
              }
            }

            // Method 4: Try command execution approach
            if (!alignmentApplied) {
              try {
                console.log("🔍 Attempting command execution approach...");
                const univerAPI = (window as any).univerAPI;
                const commandManager = univerAPI.getCommandManager();

                if (
                  commandManager &&
                  typeof commandManager.executeCommand === "function"
                ) {
                  const alignmentMap = { left: 1, center: 2, right: 3 };
                  const command = {
                    id: "sheet.command.set-range-values",
                    params: {
                      unitId: workbook.getUnitId(),
                      subUnitId: worksheet.getSheetId(),
                      range: {
                        startRow,
                        startColumn: startCol,
                        endRow,
                        endColumn: endCol,
                      },
                      value: {
                        s: {
                          ht: alignmentMap[alignmentValue] || 1,
                        },
                      },
                    },
                  };

                  await commandManager.executeCommand(
                    command.id,
                    command.params
                  );
                  console.log(
                    `✅ Applied alignment via command manager: ${alignmentValue}`
                  );
                  alignmentApplied = true;
                } else {
                  console.log(
                    "📝 Command manager methods:",
                    commandManager
                      ? Object.getOwnPropertyNames(
                          Object.getPrototypeOf(commandManager)
                        )
                      : "commandManager not available"
                  );
                }
              } catch (error) {
                console.warn("⚠️ Command execution failed:", error);
              }
            }

            // Method 5: Try cell data manipulation
            if (!alignmentApplied) {
              try {
                console.log("🔍 Attempting cell data manipulation...");
                const alignmentMap: {
                  left: number;
                  center: number;
                  right: number;
                } = { left: 1, center: 2, right: 3 };
                const styleValue = alignmentMap[alignmentValue] || 1;

                for (let row = startRow; row <= endRow; row++) {
                  for (let col = startCol; col <= endCol; col++) {
                    // Try to get current cell data and update it
                    const cellData = worksheet.getCellRaw(row, col) || {};
                    const currentStyle = cellData.s || {};
                    const newStyle = { ...currentStyle, ht: styleValue };

                    // Try different methods to set cell data
                    const setCellMethods = [
                      "setCellRaw",
                      "setCell",
                      "updateCell",
                    ];
                    for (const methodName of setCellMethods) {
                      if (typeof worksheet[methodName] === "function") {
                        try {
                          worksheet[methodName](row, col, {
                            ...cellData,
                            s: newStyle,
                          });
                          console.log(
                            `✅ Applied alignment via ${methodName} at [${row},${col}]:`,
                            newStyle
                          );
                          alignmentApplied = true;
                          break;
                        } catch (error) {
                          console.warn(`⚠️ ${methodName} failed:`, error);
                        }
                      }
                    }

                    if (alignmentApplied) break;
                  }
                  if (alignmentApplied) break;
                }
              } catch (error) {
                console.warn("⚠️ Cell data manipulation failed:", error);
              }
            }

            if (!alignmentApplied) {
              console.error(
                `❌ All text alignment methods failed for value: ${alignmentValue}`
              );
              console.log("🔍 Full range object structure:", range);
              console.log("🔍 Full worksheet object structure:", worksheet);
            }
          }

          console.log(`✅ Applied available formatting to ${action.range}`);
        } catch (formattingError) {
          console.error(`❌ Formatting failed:`, formattingError);
        }

        console.log(`Applied formatting to ${action.range}`);
      } else {
        throw new Error("UniverAPI not available");
      }
    } catch (error) {
      console.error(`Failed to format cells:`, error);
      throw error;
    }
  };

  // Smart typing indicator: show only until first assistant tokens arrive
  const lastMsg: any = messages[messages.length - 1];
  const assistantHasContent =
    lastMsg?.role === "assistant" &&
    ((Array.isArray(lastMsg.parts) &&
      lastMsg.parts.some(
        (p: any) =>
          (p.type === "text" &&
            typeof p.text === "string" &&
            p.text.trim().length > 0) ||
          p.type === "tool-invocation"
      )) ||
      (typeof lastMsg?.content === "string" &&
        lastMsg.content.trim().length > 0));

  const isLoading =
    status === "submitted" || (status === "streaming" && !assistantHasContent);

  return (
    <div className="h-full flex flex-col bg-background rounded-2xl my-2 mx-1 p-2">
      <div className="p-4 border-b">
        <div className="flex items-center justify-between">
          <h2 className="text-lg font-semibold">Spreadsheet Master 🧠</h2>
        </div>
      </div>

      <div className="flex-1 overflow-hidden">
        <ChatMessages messages={messages} isLoading={isLoading} />
      </div>

      <div className="p-4 border-t">
        <ChatInput
          input={input}
          handleInputChange={handleInputChange}
          handleSubmit={handleSubmit}
          isLoading={isLoading}
        />
      </div>
    </div>
  );
}

// Collapsible Markdown to keep assistant replies short by default
function CollapsibleMarkdown({ text }: { text: string }) {
  const [expanded, setExpanded] = React.useState(false);
  const MAX_CHARS = 280;
  const isLong = text && text.length > MAX_CHARS;
  const display = expanded || !isLong ? text : text.slice(0, MAX_CHARS) + "…";

  return (
    <div className="leading-6 prose prose-sm max-w-none dark:prose-invert">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={{
          table: () => <></>,
          th: () => <></>,
          td: () => <></>,
          p: (props) => <p className="mb-1 leading-5" {...props} />,
          h1: (props) => <h1 className="mb-1 text-base" {...props} />,
          h2: (props) => <h2 className="mb-1 text-sm" {...props} />,
          h3: (props) => <h3 className="mb-1 text-sm" {...props} />,
          ul: (props) => <ul className="list-disc ml-4 mb-1" {...props} />,
          ol: (props) => <ol className="list-decimal ml-4 mb-1" {...props} />,
          code: ({ inline, children, ...props }: any) => (
            <code
              className={cn(
                inline
                  ? "px-1 py-0.5 rounded bg-black/10"
                  : "block p-3 rounded bg-black/10 overflow-x-auto"
              )}
              {...props}
            >
              {children}
            </code>
          ),
          a: (props) => (
            <a
              className="underline text-blue-600"
              target="_blank"
              rel="noreferrer"
              {...props}
            />
          ),
        }}
      >
        {display}
      </ReactMarkdown>
      {isLong && (
        <button
          type="button"
          className="mt-1 text-[11px] underline text-blue-600"
          onClick={() => setExpanded((v) => !v)}
        >
          {expanded ? "Show less" : "Show details"}
        </button>
      )}
    </div>
  );
}
function ToolBadge({
  label,
  toolName,
  state,
}: {
  label: string;
  toolName?: string;
  state: "call" | "result" | "error";
}) {
  const isRunning = state === "call";
  return (
    <div className="my-1">
      <div
        className={cn(
          "inline-flex items-center gap-2 text-xs rounded-md px-2 py-1 border",
          isRunning
            ? "border-blue-300 bg-blue-50 text-blue-700"
            : "border-green-300 bg-green-50 text-green-700"
        )}
      >
        {isRunning ? (
          <Loader2 className="h-3.5 w-3.5 animate-spin" />
        ) : (
          <CheckCircle2 className="h-3.5 w-3.5" />
        )}
        <Wrench className="h-3.5 w-3.5 opacity-70" />
        <span className="truncate max-w-[18rem]">{label}</span>
        {toolName ? (
          <span className="ml-2 text-[10px] opacity-70">({toolName})</span>
        ) : null}
      </div>
    </div>
  );
}

function describeTool(toolName: string, args: any): string {
  switch (toolName) {
    case "list_columns":
      return "List columns";
    case "calculate_total":
      return `Total for ${args?.column ?? "column"}`;
    case "create_pivot_table":
      return `Pivot by ${args?.groupBy ?? "group"}`;
    case "generate_chart":
      return `${args?.chart_type ?? "chart"} chart${
        args?.title ? `: ${args.title}` : ""
      }`;
    case "switch_sheet":
      return `${args?.action === "switch" ? "Switch to" : "Analyze"} ${
        args?.sheetName ?? "sheet"
      }`;
    case "financial_intelligence":
      return `Analyze: ${String(args?.user_request ?? "")}`;
    case "add_filter":
      return "Apply filter";
    case "conditional_formatting":
      return "Conditional formatting";
    case "ask_for_range":
      return "Need range specification";
    case "get_sheet_context":
      return "Analyze spreadsheet";
    default:
      return "Processing";
  }
}
