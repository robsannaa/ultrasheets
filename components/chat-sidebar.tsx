"use client";

import { useRef, useEffect } from "react";
import * as React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Send, Loader2 } from "lucide-react";
import { cn } from "@/lib/utils";
import { applyFormatting } from "@/lib/univer-helpers";
import { useChat } from "ai/react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { ToolBadge } from "./ToolBadge";
import Image from "next/image";

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

  // Use messages as-is; no de-duplication of adjacent assistant frames
  const mergedMessages = messages;

  // Show all assistant text output without aggressive de-duplication

  return (
    <ScrollArea
      className="h-full px-1 sm:px-2 py-2 sm:py-4"
      ref={scrollAreaRef}
    >
      <div className="space-y-3">
        {mergedMessages.map((message) => {
          // Skip rendering empty assistant frames (no text and no tool badges)
          const hasParts = Array.isArray(message.parts);
          const hasTextPart = hasParts
            ? message.parts.some(
                (p: any) =>
                  p.type === "text" &&
                  typeof p.text === "string" &&
                  p.text.trim().length > 0
              )
            : false;
          const hasToolPart = hasParts
            ? message.parts.some((p: any) => p.type === "tool-invocation")
            : false;
          const hasContent =
            (typeof message.content === "string" &&
              message.content.trim().length > 0) ||
            hasTextPart ||
            hasToolPart;
          if (message.role === "assistant" && !hasContent) return null;

          return (
            <div
              key={message.id}
              className={cn(
                "flex items-start gap-2",
                message.role === "user" && "flex-row-reverse"
              )}
            >
              <Avatar className="w-7 h-7">
                <AvatarFallback>
                  {message.role === "user" ? "U" : "A"}
                </AvatarFallback>
              </Avatar>
              <div
                className={cn(
                  "max-w-full sm:max-w-[85%] md:max-w-[75%] rounded-lg p-2 sm:p-3 text-sm break-words whitespace-pre-wrap",
                  message.role === "user"
                    ? "bg-primary text-primary-foreground"
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
                          <CollapsibleMarkdown
                            key="assistant-text"
                            text={text}
                          />
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
                              <div
                                key={callId}
                                className="break-words max-w-full overflow-y-hidden"
                              >
                                <ToolBadge
                                  label={label}
                                  toolName={toolName}
                                  state="result"
                                />
                              </div>
                            );
                          }
                          return null;
                        }
                        default:
                          // When parts exist, we do not render message.content fallback to avoid duplicates
                          return null;
                      }
                    })}
                  </>
                ) : message.role === "assistant" ? (
                  <div className="leading-6 prose prose-sm max-w-none dark:prose-invert break-words">
                    <ReactMarkdown remarkPlugins={[remarkGfm]}>
                      {message.content}
                    </ReactMarkdown>
                  </div>
                ) : (
                  message.content
                )}
              </div>
            </div>
          );
        })}
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

  const handleFormSubmit = (e: React.FormEvent<HTMLFormElement>) => {
    if (isLoading || !input.trim()) {
      e.preventDefault();
      return;
    }
    handleSubmit(e);
  };

  return (
    <form onSubmit={handleFormSubmit} className="flex gap-2">
      <Input
        placeholder="Ask about your spreadsheet data..."
        value={input}
        onChange={handleInputChange}
        onKeyDown={handleKeyDown}
        className="flex-1 text-sm"
      />
      <Button
        type="submit"
        size="icon"
        disabled={isLoading || !input.trim()}
        aria-busy={isLoading}
        aria-disabled={isLoading || !input.trim()}
        className="shrink-0"
      >
        {isLoading ? (
          <Loader2 className="h-4 w-4 animate-spin" />
        ) : (
          <Send className="h-4 w-4" />
        )}
      </Button>
    </form>
  );
}

import { extractEnhancedWorkbookData } from "../lib/enhanced-context";
import type { CleanSheetContext } from "../lib/clean-context-tools";

// Rich workbook context with multi-table detection and recent action log
function extractWorkbookData() {
  // Use the enhanced context system
  return extractEnhancedWorkbookData();
}

export function ChatSidebar({ onMobileClose }: { onMobileClose?: () => void }) {
  const [mounted, setMounted] = React.useState(false);
  React.useEffect(() => setMounted(true), []);

  // Auto-close mobile sidebar after successful actions

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
      // Provide lightweight recent action history to improve pronoun resolution (e.g., "format it")
      workbookData: mounted
        ? (() => {
            try {
              const extracted = extractWorkbookData();
              return {
                ...extracted,
                // ultraActionLog is populated by client-side tool executions
                recentActions:
                  typeof window !== "undefined" &&
                  (window as any).ultraActionLog
                    ? (window as any).ultraActionLog
                    : [],
              };
            } catch (error) {
              console.warn("Failed to extract workbook data:", error);
              return {
                sheets: [],
                recentActions:
                  typeof window !== "undefined" &&
                  (window as any).ultraActionLog
                    ? (window as any).ultraActionLog
                    : [],
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
            const semanticKey = `${action.toolName}:${JSON.stringify(action.params)}`;
            const now = Date.now();
            const lastExecution = recentExecutionsRef.current.get(semanticKey);
            
            // Only block if IDENTICAL parameters within rapid timeframe (likely accidental duplicates)
            const rapidDuplicateWindow = 2000; // 2 seconds for identical calls
            
            if (lastExecution && (now - lastExecution) < rapidDuplicateWindow) {
              console.warn(`🚫 BLOCKED: Identical ${action.toolName} call with same parameters within ${rapidDuplicateWindow/1000}s - likely duplicate`);
              if (callId) executedToolCallIdsRef.current.add(callId);
              continue;
            }
            
            // Record this execution
            recentExecutionsRef.current.set(semanticKey, now);
            
            // Clean up old entries (keep only last 10 seconds)
            for (const [key, timestamp] of recentExecutionsRef.current.entries()) {
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
              const worksheet = univerAPI.getActiveWorkbook().getActiveSheet();
              const source = worksheet.getRange(action.sourceRange);
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
              const destRange = worksheet.getRange(
                destRowIndex,
                destColIndex,
                numRows,
                numCols
              );
              destRange.setValues(values);
              if (action.clearSource) source.clear();
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
                  tool: "smart_add_column",
                  params: { count: action.actions.length },
                  result: "ok",
                });
                if (w.ultraActionLog.length > 50) w.ultraActionLog.shift();
              } catch {}
            } else if (action.type === "getCleanSheetContext") {
              // Get real sheet context with zero assumptions
              const { getCleanSheetContext } = await import(
                "../lib/clean-context-tools"
              );
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");

              const cleanContext = await getCleanSheetContext(univerAPI);

              // Return the context to the LLM by updating the message
              console.log("🧠 Clean sheet context:", cleanContext);
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
            } else if (action.type === "smartAddColumnWithContext") {
              // Add column using SEMANTIC intelligence
              const { getCleanSheetContext } = await import(
                "../lib/clean-context-tools"
              );
              const univerAPI = (window as any).univerAPI;
              if (!univerAPI) throw new Error("Univer API not available");

              const cleanContext = await getCleanSheetContext(univerAPI);
              const { columnName, formulaPattern, defaultValue, insertBefore, insertAfter, between } =
                action.params;

              // Find the best table to add column to
              const targetRegion = cleanContext.analysis.dataRegions[0];
              if (!targetRegion) throw new Error("No data regions found");

              // Smart column positioning logic
              let targetColumnLetter: string;
              
              if (between && Array.isArray(between) && between.length === 2) {
                // Insert between two columns (e.g., between D and E should insert at E)
                const [leftCol, rightCol] = between;
                // Convert column names to letters if needed
                const leftLetter = typeof leftCol === 'string' && leftCol.length === 1 ? leftCol : 
                  (targetRegion as any).headers?.find((h: any) => h.name === leftCol)?.letter || leftCol;
                const rightLetter = typeof rightCol === 'string' && rightCol.length === 1 ? rightCol :
                  (targetRegion as any).headers?.find((h: any) => h.name === rightCol)?.letter || rightCol;
                
                targetColumnLetter = rightLetter; // Insert at the position of the right column
                console.log(`📍 Positioning: Between ${leftLetter} and ${rightLetter} → Insert at ${targetColumnLetter}`);
              } else if (insertBefore) {
                // Insert before a specific column
                const beforeLetter = typeof insertBefore === 'string' && insertBefore.length === 1 ? insertBefore :
                  (targetRegion as any).headers?.find((h: any) => h.name === insertBefore)?.letter || insertBefore;
                targetColumnLetter = beforeLetter;
                console.log(`📍 Positioning: Before ${beforeLetter} → Insert at ${targetColumnLetter}`);
              } else if (insertAfter) {
                // Insert after a specific column
                const afterLetter = typeof insertAfter === 'string' && insertAfter.length === 1 ? insertAfter :
                  (targetRegion as any).headers?.find((h: any) => h.name === insertAfter)?.letter || insertAfter;
                const nextColIndex = afterLetter.charCodeAt(0) - 65 + 1;
                targetColumnLetter = String.fromCharCode(65 + nextColIndex);
                console.log(`📍 Positioning: After ${afterLetter} → Insert at ${targetColumnLetter}`);
              } else {
                // Default: use next available column
                const nextColumn = cleanContext.analysis.emptyAreas.nextColumns[0];
                if (!nextColumn) throw new Error("No space for new column");
                targetColumnLetter = nextColumn;
                console.log(`📍 Positioning: Default next available column → ${targetColumnLetter}`);
              }

              const headerRow = parseInt(targetRegion.range.match(/\d+/)![0]);
              const worksheet = univerAPI.getActiveWorkbook().getActiveSheet();

              // Insert column if position is specified (between, before, after)
              if (between || insertBefore || insertAfter) {
                const targetColIndex = targetColumnLetter.charCodeAt(0) - 65;
                console.log(`📍 Inserting new column at index ${targetColIndex} (letter ${targetColumnLetter})`);
                
                // Use correct Univer.js API: insertColumns (plural) not insertColumn
                if (typeof worksheet.insertColumns === 'function') {
                  worksheet.insertColumns(targetColIndex, 1);
                  console.log(`✅ Successfully inserted column using insertColumns API`);
                } else {
                  console.warn('⚠️ insertColumns method not available, column insertion skipped');
                  // Fallback: proceed without actual insertion (data will be added to target position)
                }
              }

              // 🧠 SEMANTIC INTELLIGENCE: Auto-detect if this column needs a calculation
              let smartFormula = formulaPattern;

              if (!smartFormula && (targetRegion as any).semanticAnalysis) {
                // Look for matching calculation in semantic analysis
                const matchingCalc = (
                  targetRegion as any
                ).semanticAnalysis.possibleCalculations.find(
                  (calc: any) =>
                    calc.newColumnName.toLowerCase() ===
                      columnName.toLowerCase() ||
                    columnName
                      .toLowerCase()
                      .includes(calc.newColumnName.toLowerCase()) ||
                    calc.newColumnName
                      .toLowerCase()
                      .includes(columnName.toLowerCase())
                );

                if (matchingCalc) {
                  smartFormula = matchingCalc.formula;
                  console.log(
                    `🧠 SEMANTIC MATCH: "${columnName}" → ${matchingCalc.description}`
                  );
                  console.log(`📊 Auto-generated formula: ${smartFormula}`);
                }
              }

              // Add header  
              const headerColIndex = targetColumnLetter.charCodeAt(0) - 65;
              const headerRowIndex = headerRow - 1;
              worksheet
                .getRange(headerRowIndex, headerColIndex, 1, 1)
                .setValue(columnName);

              // Add data with intelligent formula or fallback
              if (smartFormula) {
                for (let i = 1; i <= targetRegion.rowCount; i++) {
                  const dataRow = headerRow + i;
                  const formula = smartFormula.replace(
                    /\{row\}/g,
                    dataRow.toString()
                  );
                  const finalFormula = formula.startsWith("=")
                    ? formula
                    : `=${formula}`;
                  worksheet
                    .getRange(headerRowIndex + i, headerColIndex, 1, 1)
                    .setValue(finalFormula);
                }
                console.log(
                  `✨ Added column "${columnName}" with intelligent formula: ${smartFormula}`
                );
              } else if (defaultValue !== undefined) {
                for (let i = 1; i <= targetRegion.rowCount; i++) {
                  worksheet
                    .getRange(headerRowIndex + i, headerColIndex, 1, 1)
                    .setValue(defaultValue);
                }
                console.log(
                  `✅ Added column "${columnName}" with default value: ${defaultValue}`
                );
              } else {
                console.log(
                  `ℹ️ Added empty column "${columnName}" - no semantic match found`
                );
              }

              // Record a rich recent action so follow-ups like "format it" can target this column
              try {
                const headerCell = `${targetColumnLetter}${headerRow}`;
                const dataRange = `${targetColumnLetter}${
                  headerRow + 1
                }:${targetColumnLetter}$${headerRow + targetRegion.rowCount}`.replace(
                  /\$+/g,
                  ""
                );
                const w: any = window as any;
                w.ultraActionLog = w.ultraActionLog || [];
                w.ultraActionLog.push({
                  at: new Date().toISOString(),
                  tool: "smart_add_column",
                  params: {
                    columnName,
                    headerCell,
                    dataRange,
                    tableRange: targetRegion.range,
                  },
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

    // Allow special, explicit target tokens resolved via Univer context
    // __AUTO_HEADERS__: resolve to header row of the primary data region (first detected region)
    if (action?.range === "__AUTO_HEADERS__") {
      const { getCleanSheetContext } = await import(
        "../lib/clean-context-tools"
      );
      const clean = await getCleanSheetContext(univerAPI);
      const region = clean?.analysis?.dataRegions?.[0];
      if (!region?.range)
        throw new Error("No data region found to resolve headers range");
      const match = region.range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!match) throw new Error(`Invalid region range: ${region.range}`);
      const startCol = match[1];
      const startRow = parseInt(match[2], 10);
      const endCol = match[3];
      // Header row is the first row of the region
      action = {
        ...action,
        range: `${startCol}${startRow}:${endCol}${startRow}`,
      };
    }

    await applyFormatting(univerAPI, action);
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
    aiLoading ||
    status === "submitted" ||
    (status === "streaming" && !assistantHasContent);

  return (
    <div
      className={`h-full flex flex-col bg-background border rounded-2xl m-1`}
    >
      {/* Only show header on desktop */}
      {!onMobileClose && (
        <div className="px-2 border-b">
          <div className="flex items-center justify-between">
            <h2 className="text-base font-semibold">CellChat</h2>
            <Image src="/logo.png" alt="CellChat" width={64} height={64} />
          </div>
        </div>
      )}

      <div className="flex-1 overflow-hidden" suppressHydrationWarning>
        {mounted ? (
          <ChatMessages messages={messages} isLoading={isLoading} />
        ) : (
          <div className="h-full" />
        )}
      </div>

      <div className={`${onMobileClose ? "p-3" : "px-3 py-3"}`}>
        <ChatInput
          input={input}
          handleInputChange={handleInputChange}
          handleSubmit={handleSubmit}
          isLoading={mounted ? isLoading : true}
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
                  ? "px-1 py-0.5 rounded bg-muted"
                  : "block p-3 rounded bg-muted overflow-x-auto"
              )}
              {...props}
            >
              {children}
            </code>
          ),
          a: (props) => (
            <a
              className="underline text-primary"
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
          className="mt-1 text-[11px] underline text-primary"
          onClick={() => setExpanded((v) => !v)}
        >
          {expanded ? "Show less" : "Show details"}
        </button>
      )}
    </div>
  );
}

function describeTool(toolName: string, args: any): string {
  switch (toolName) {
    case "list_columns":
      return "List columns";
    case "calculate_total":
      return `Total for ${args?.column ?? "column"}`;
    case "add_smart_totals":
      return `Smart totals${
        args?.columns
          ? ` for ${args.columns.join(", ")}`
          : " for all calculable columns"
      }`;
    case "format_recent_totals":
      return `Format recent totals as ${args?.currency || "USD"}`;
    case "format_currency_column":
      return `Format currency column as ${args?.currency || "USD"}`;
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
    case "set_cell_formula":
      return `Set formula`;
    case "find_cell":
      return "Find cell";
    case "format_as_table":
      return `Format as table`;
    case "ask_for_range":
      return "Need range specification";
    case "get_sheet_context":
      return "Analyze spreadsheet";
    case "get_workbook_snapshot":
      return "Analyze workbook";
    default:
      return "Processing";
  }
}
