"use client";

import { useRef, useEffect } from "react";
import * as React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Send } from "lucide-react";
import { cn } from "@/lib/utils";
import { applyFormatting } from "@/lib/univer-helpers";
import { useChat } from "ai/react";
import { ToolBadge } from "./ToolBadge";
import Image from "next/image";
import { ChatMessages } from "./chat/chat-messages";
import { ChatInput } from "./chat/chat-input";
import { CollapsibleMarkdown } from "./chat/collapsible-markdown";

// ChatMessages now imported from ./chat/chat-messages

// ChatInput now imported from ./chat/chat-input

import { getChatContext, getAISummary } from "../lib/unified-context-system";
import type { ChatContext } from "../lib/unified-context-system";

// Rich workbook context using unified system with AI optimizations
async function extractWorkbookDataUnified() {
  try {
    // Use AI-optimized context extraction
    const chatContext = await getChatContext(undefined, {
      format: "summary",
      maxTokens: 1500
    });
    
    const aiSummary = await getAISummary();
    
    return {
      contextText: chatContext.contextText,
      metadata: chatContext.metadata,
      tables: aiSummary.tables,
      selection: aiSummary.selection,
      recommendations: aiSummary.recommendations
    };
  } catch (error) {
    console.warn("Unified context extraction failed, using fallback:", error);
    return {
      contextText: "Spreadsheet context unavailable",
      metadata: { tablesCount: 0, hasSelection: false, estimatedTokens: 0, contextType: "unknown" },
      tables: [],
      selection: { hasSelection: false },
      recommendations: []
    };
  }
}

// Legacy fallback function (kept for compatibility during transition)
function extractWorkbookData() {
  // This will be removed after full migration
  try {
    const w: any = window as any;
    return {
      sheets: [],
      cleanContext: w.__latestCleanContext || null,
    };
  } catch {
    return { sheets: [] };
  }
}

export function ChatSidebar({ onMobileClose }: { onMobileClose?: () => void }) {
  const sidebarRef = useRef<HTMLDivElement>(null);

  const [mounted, setMounted] = React.useState(false);
  React.useEffect(() => setMounted(true), []);

  // Auto-close mobile sidebar after successful actions

  // Enhanced context caching with unified system
  const [unifiedContext, setUnifiedContext] = React.useState<any>(null);
  const [contextMetadata, setContextMetadata] = React.useState<any>(null);
  
  React.useEffect(() => {
    const loadUnifiedContext = async () => {
      try {
        const w: any = window as any;
        if (!w.univerAPI) return;
        
        const contextData = await extractWorkbookDataUnified();
        setUnifiedContext(contextData);
        setContextMetadata(contextData.metadata);
        
        // Legacy compatibility
        w.__latestUnifiedContext = contextData;
      } catch (error) {
        console.warn("Failed to load unified context:", error);
      }
    };
    
    loadUnifiedContext();
    
    // Refresh context every 10 seconds or when user interacts
    const interval = setInterval(loadUnifiedContext, 10000);
    return () => clearInterval(interval);
  }, []);

  // Parent (ResizablePanel) controls width; no internal resizer

  const getClientEnv = React.useCallback(() => {
    try {
      const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
      const locale = navigator.language;
      // userAgentData is not available in all browsers
      const platform =
        (navigator as any).userAgentData?.platform || navigator.platform;
      return { timeZone, locale, platform };
    } catch {
      return null;
    }
  }, []);
  const {
    messages,
    input,
    handleInputChange,
    handleSubmit,
    status,
    isLoading: aiLoading,
  } = useChat({
    api: "/api/chat",
    maxSteps: 5,
    initialMessages: [
      {
        id: "1",
        role: "assistant" as const,
        content: "**Hey 👋** \nHow can I help you with your spreadsheet today?",
      },
    ],
    // Send only client environment - LLM will use tools to get current spreadsheet state
    body: {
      clientEnv: mounted ? getClientEnv() : null,
      // AI-optimized context with unified system
      workbookData: mounted
        ? (() => {
            try {
              // Use cached unified context or fallback
              const contextData = unifiedContext || extractWorkbookData();
              return {
                ...contextData,
                // Include context metadata for AI optimization
                contextMetadata,
                // ultraActionLog is populated by client-side tool executions
                recentActions:
                  typeof window !== "undefined" &&
                  (window as any).ultraActionLog
                    ? (window as any).ultraActionLog
                    : [],
                // Performance hint for AI
                optimizedContext: !!unifiedContext
              };
            } catch (error) {
              console.warn("Failed to extract unified context:", error);
              return {
                sheets: [],
                contextText: "Context extraction failed",
                recentActions:
                  typeof window !== "undefined" &&
                  (window as any).ultraActionLog
                    ? (window as any).ultraActionLog
                    : [],
                optimizedContext: false
              };
            }
          })()
        : null,
    },
  });

  // Keep track of executed tool invocations to avoid duplicate runs during streaming updates
  const executedToolCallIdsRef = React.useRef<Set<string>>(new Set());
  // Semantic deduplication based on tool + parameters to prevent multiple identical calls
  const recentExecutionsRef = React.useRef<Map<string, number>>(new Map());

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

          // SMART deduplication: prevent IDENTICAL tool calls (same params) within rapid succession
          if (action.type === "executeUniverTool") {
            const semanticKey = `${action.toolName}:${JSON.stringify(
              action.params
            )}`;
            const now = Date.now();
            const lastExecution = recentExecutionsRef.current.get(semanticKey);

            // Only block if IDENTICAL parameters within rapid timeframe (likely accidental duplicates)
            const rapidDuplicateWindow = 2000; // 2 seconds for identical calls

            if (lastExecution && now - lastExecution < rapidDuplicateWindow) {
              console.warn(
                `🚫 BLOCKED: Identical ${
                  action.toolName
                } call with same parameters within ${
                  rapidDuplicateWindow / 1000
                }s - likely duplicate`
              );
              if (callId) executedToolCallIdsRef.current.add(callId);
              continue;
            }

            // Record this execution
            recentExecutionsRef.current.set(semanticKey, now);

            // Clean up old entries (keep only last 10 seconds)
            for (const entry of Array.from(recentExecutionsRef.current.entries())) {
              const [key, timestamp] = entry;
              if (now - timestamp > 10000) {
                recentExecutionsRef.current.delete(key);
              }
            }
          }

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
            } else if (action.type === "setRangeValues") {
              const w: any = window as any;
              if (!(window as any).univerAPI)
                throw new Error("Univer API not available");
              const univerAPI = (window as any).univerAPI;
              const workbook = univerAPI.getActiveWorkbook();
              const worksheet = workbook.getActiveSheet();
              const startCell: string = action.startCell;
              const values: any[][] = action.values || [];
              if (!Array.isArray(values) || values.length === 0)
                throw new Error("No values provided");
              const startColLetter = startCell.replace(/\d+/g, "");
              const startRowNumber = parseInt(
                startCell.replace(/\D+/g, ""),
                10
              );
              const startColIndex = startColLetter.charCodeAt(0) - 65;
              const startRowIndex = startRowNumber - 1;
              const numRows = values.length;
              const numCols = Math.max(
                0,
                ...values.map((r: any[]) => r.length)
              );
              const range = worksheet.getRange(
                startRowIndex,
                startColIndex,
                numRows,
                numCols
              );
              range.setValues(values);
              try {
                const formula = univerAPI.getFormula();
                formula.executeCalculation();
              } catch {}
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "bulk_set_values",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "setRangeValuesByRange") {
              const w: any = window as any;
              if (!(window as any).univerAPI)
                throw new Error("Univer API not available");
              const univerAPI = (window as any).univerAPI;
              const workbook = univerAPI.getActiveWorkbook();
              const worksheet = workbook.getActiveSheet();
              const rangeStr: string = action.range;
              const values: any[][] = action.values || [];
              if (!Array.isArray(values) || values.length === 0)
                throw new Error("No values provided");
              const range = worksheet.getRange(rangeStr);
              range.setValues(values);
              try {
                const formula = univerAPI.getFormula();
                formula.executeCalculation();
              } catch {}
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "set_range_values",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "clearRange") {
              const w: any = window as any;
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");
              const worksheet = univerAPI.getActiveWorkbook().getActiveSheet();
              worksheet.getRange(action.range).clear();
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "clear_range",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "clearRangeContents") {
              const w: any = window as any;
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");
              const worksheet = univerAPI.getActiveWorkbook().getActiveSheet();
              worksheet.getRange(action.range).clearContents?.();
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "clear_range_contents",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "moveRange") {
              const w: any = window as any;
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");
              const workbook = univerAPI.getActiveWorkbook();
              const currentSheet = workbook.getActiveSheet();
              const currentSheetName = currentSheet
                ?.getSheet?.()
                .getSnapshot?.().name;

              // Read from source on current active sheet
              const source = currentSheet.getRange(action.sourceRange);
              const values = source.getValues();
              const destTopLeft = action.destStartCell;
              const destColLetter = destTopLeft.replace(/\d+/g, "");
              const destRowNumber = parseInt(
                destTopLeft.replace(/\D+/g, ""),
                10
              );
              const destColIndex = destColLetter.charCodeAt(0) - 65;
              const destRowIndex = destRowNumber - 1;
              const numRows = values.length;
              const numCols = Math.max(
                0,
                ...values.map((r: any[]) => r.length)
              );

              // Resolve destination sheet (optionally switch)
              let destSheet = currentSheet;
              const destSheetName: string | undefined = action.destSheetName;
              if (
                destSheetName &&
                typeof workbook.getSheetByName === "function"
              ) {
                destSheet = workbook.getSheetByName(destSheetName) || destSheet;
                if (!destSheet) {
                  // try create then resolve
                  try {
                    if (typeof workbook.insertSheet === "function") {
                      workbook.insertSheet(destSheetName);
                      destSheet =
                        workbook.getSheetByName(destSheetName) || currentSheet;
                    }
                  } catch {}
                }
                // set active to destination temporarily if API requires
                try {
                  if (
                    destSheet &&
                    typeof workbook.setActiveSheet === "function"
                  ) {
                    workbook.setActiveSheet(destSheetName);
                  }
                } catch {}
              }

              const destRange = destSheet.getRange(
                destRowIndex,
                destColIndex,
                numRows,
                numCols
              );
              destRange.setValues(values);
              if (action.clearSource) source.clear();
              // restore active sheet
              try {
                if (
                  currentSheetName &&
                  typeof workbook.setActiveSheet === "function" &&
                  destSheetName &&
                  destSheetName !== currentSheetName
                ) {
                  workbook.setActiveSheet(currentSheetName);
                }
              } catch {}
              try {
                const formula = univerAPI.getFormula();
                formula.executeCalculation();
              } catch {}
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "move_range",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "transposeRange") {
              const w: any = window as any;
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");
              const worksheet = univerAPI.getActiveWorkbook().getActiveSheet();
              const source = worksheet.getRange(action.sourceRange);
              const original = source.getValues();
              const transposed: any[][] = original[0]
                ? original[0].map((_: any, colIndex: number) =>
                    original.map((row: any[]) => row[colIndex])
                  )
                : [];
              const destTopLeft = action.destStartCell;
              const destColLetter = destTopLeft.replace(/\d+/g, "");
              const destRowNumber = parseInt(
                destTopLeft.replace(/\D+/g, ""),
                10
              );
              const destColIndex = destColLetter.charCodeAt(0) - 65;
              const destRowIndex = destRowNumber - 1;
              const numRows = transposed.length;
              const numCols = Math.max(
                0,
                ...transposed.map((r: any[]) => r.length)
              );
              const destRange = worksheet.getRange(
                destRowIndex,
                destColIndex,
                numRows,
                numCols
              );
              destRange.setValues(transposed);
              try {
                const formula = univerAPI.getFormula();
                formula.executeCalculation();
              } catch {}
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "transpose_range",
                  params: action,
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "multipleActions") {
              // Handle multiple actions in sequence
              const w: any = window as any;
              for (const subAction of action.actions) {
                // Recursively handle each sub-action
                if (subAction.type === "setCellValue") {
                  if (!(window as any).univerAPI)
                    throw new Error("Univer API not available");
                  const univerAPI = (window as any).univerAPI;
                  const workbook = univerAPI.getActiveWorkbook();
                  const worksheet = workbook.getActiveSheet();
                  const cellRef: string = subAction.cell;
                  const isFormula: boolean = !!subAction.formula;
                  const value = subAction.value;
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
                }
                // Add other sub-action types as needed
              }
              // Recalculate formulas after all actions
              try {
                const formula = (window as any).univerAPI.getFormula();
                formula.executeCalculation();
              } catch {}
              // Log the multiple actions
              try {
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "add_column",
                  params: { count: action.actions.length },
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "getCleanSheetContext") {
              // Get unified context with AI optimizations
              const { getChatContext } = await import(
                "../lib/unified-context-system"
              );
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");

              const contextData = await getChatContext(undefined, {
                format: "detailed",
                maxTokens: 3000
              });

              // Return the context to the LLM
              console.log("🧠 Unified context:", contextData);
              
              // Update cached context
              setUnifiedContext(contextData);
              setContextMetadata(contextData.metadata);
            } else if (action.type === "getCleanWorkbookState") {
              // Get full workbook state with zero assumptions
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");

              const workbook = univerAPI.getActiveWorkbook();
              const sheets = workbook.getSheets();

              const workbookState = {
                totalSheets: sheets.length,
                activeSheet: workbook.getActiveSheet().getSheet().getSnapshot()
                  .name,
                sheets: sheets.map((sheet: any) => ({
                  name: sheet.getSheet().getSnapshot().name,
                  isActive: sheet === workbook.getActiveSheet(),
                })),
              };

              console.log("🔍 Clean workbook state:", workbookState);
              (window as any).__lastWorkbookState = workbookState;
            } else if (action.type === "addColumn") {
              // Simple, precise column addition using exact Univer API
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");
              
              const { sheetName, insertAtColumn, columnName, defaultValue } = action.params;
              const workbook = univerAPI.getActiveWorkbook();
              let worksheet = workbook.getActiveSheet();
              
              // Switch to specified sheet if provided
              if (sheetName && typeof workbook.getSheetByName === "function") {
                const targetSheet = workbook.getSheetByName(sheetName);
                if (targetSheet) {
                  worksheet = targetSheet;
                  workbook.setActiveSheet(sheetName);
                }
              }
              
              // Convert column letter to index for insertion
              const colIndex = insertAtColumn.charCodeAt(0) - 65;
              
              // Insert column using precise Univer API
              if (typeof worksheet.insertColumns === "function") {
                worksheet.insertColumns(colIndex, 1);
              }
              
              // Set header if provided
              if (columnName) {
                worksheet.getRange(0, colIndex, 1, 1).setValue(columnName);
              }
              
              // Set default values if provided
              if (defaultValue !== undefined) {
                // Detect actual data range instead of hardcoding 10 rows
                try {
                  // Find the last row with data by checking multiple columns
                  let lastDataRow = 0;
                  const maxColumnsToCheck = Math.min(colIndex, 10); // Check up to 10 columns or up to insertion point
                  
                  for (let checkCol = 0; checkCol < maxColumnsToCheck; checkCol++) {
                    // Check each column for data up to row 100 (reasonable limit)
                    for (let checkRow = 1; checkRow <= 100; checkRow++) {
                      const cell = worksheet.getRange(checkRow, checkCol, 1, 1);
                      const value = cell.getValue();
                      if (value !== null && value !== undefined && value !== "") {
                        lastDataRow = Math.max(lastDataRow, checkRow);
                      }
                    }
                  }
                  
                  // If we found data, fill from row 1 to lastDataRow, otherwise don't fill anything
                  if (lastDataRow > 0) {
                    for (let row = 1; row <= lastDataRow; row++) {
                      worksheet.getRange(row, colIndex, 1, 1).setValue(defaultValue);
                    }
                    console.log(`✅ Added column with default values for ${lastDataRow} data rows`);
                  } else {
                    console.log(`ℹ️ No existing data detected, skipping default value fill`);
                  }
                } catch (error) {
                  console.warn("Failed to detect data range, skipping default values:", error);
                }
              }
              
              try {
                const w: any = window as any;
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "add_column",
                  params: action.params,
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
    const univerAPI = (window as any).univerAPI;
    if (!univerAPI) throw new Error("Univer API not available");

    // All ranges must be explicitly provided - no auto-resolution

    await applyFormatting(univerAPI, action);
  };

  // Typing indicator: show only until first assistant tokens arrive
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
    aiLoading ||
    status === "submitted" ||
    (status === "streaming" && !assistantHasContent);

  return (
    <div
      ref={sidebarRef}
      className={"border-l bg-card flex flex-col h-full relative rounded-l-xl"}
    >
      {/* Only show header on desktop */}
      {!onMobileClose && (
        <div className="p-4 border-b rounded-tl-xl">
          <h2 className="font-semibold text-lg">Chat Assistant</h2>
          <p className="text-sm text-muted-foreground">
            AI-powered conversation
          </p>
        </div>
      )}

      <div className="flex-1 relative overflow-hidden" suppressHydrationWarning>
        <div className="absolute top-0 left-0 right-0 h-8 bg-gradient-to-b from-card to-transparent z-10 pointer-events-none rounded-tl-xl" />
        {mounted ? (
          <ChatMessages 
            messages={messages} 
            isLoading={isLoading} 
            aiPerformanceMode={contextMetadata?.contextType === 'analysis' || contextMetadata?.estimatedTokens > 1500}
          />
        ) : (
          <div className="h-full" />
        )}
      </div>

      <div className="p-2 border-t bg-background rounded-bl-xl">
        <ChatInput
          input={input}
          handleInputChange={handleInputChange}
          handleSubmit={handleSubmit}
          isLoading={mounted ? isLoading : true}
          contextMetadata={contextMetadata}
        />
      </div>
    </div>
  );
}

// CollapsibleMarkdown now imported from ./chat/collapsible-markdown
// describeTool function moved to chat-messages.tsx to avoid duplication
