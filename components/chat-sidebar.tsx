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

import { extractEnhancedWorkbookData } from "../lib/enhanced-context";
import type { CleanSheetContext } from "../lib/clean-context-tools";

// Rich workbook context with multi-table detection and recent action log
function extractWorkbookData() {
  // Use the enhanced context system
  const enhanced = extractEnhancedWorkbookData();
  try {
    const w: any = window as any;
    return {
      ...enhanced,
      cleanContext: w.__latestCleanContext || null,
    };
  } catch {
    return enhanced;
  }
}

export function ChatSidebar({ onMobileClose }: { onMobileClose?: () => void }) {
  const [mounted, setMounted] = React.useState(false);
  React.useEffect(() => setMounted(true), []);

  // Auto-close mobile sidebar after successful actions

  // Capture a live clean sheet context for the LLM to use in deciding tool args
  React.useEffect(() => {
    const loadClean = async () => {
      try {
        const w: any = window as any;
        if (!w.univerAPI) return;
        const { getCleanSheetContext } = await import(
          "../lib/clean-context-tools"
        );
        const clean = await getCleanSheetContext(w.univerAPI);
        w.__latestCleanContext = clean;
      } catch {}
    };
    loadClean();
  }, []);

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
            for (const [
              key,
              timestamp,
            ] of recentExecutionsRef.current.entries()) {
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
              const { columnName, formulaPattern, defaultValue } =
                action.params;

              // Find the best table to add column to
              const targetRegion = cleanContext.analysis.dataRegions[0];
              if (!targetRegion) throw new Error("No data regions found");

              // Smart column positioning logic
              // Robust A1 helpers and anchor resolution (supports AA, AB, ...)
              const letterToIndex = (letters: string): number => {
                let idx = 0;
                const up = letters.toUpperCase();
                for (let i = 0; i < up.length; i++) {
                  idx = idx * 26 + (up.charCodeAt(i) - 64);
                }
                return idx - 1; // 0-based
              };
              const indexToLetter = (index: number): string => {
                let n = index + 1;
                let s = "";
                while (n > 0) {
                  const rem = (n - 1) % 26;
                  s = String.fromCharCode(65 + rem) + s;
                  n = Math.floor((n - 1) / 26);
                }
                return s;
              };
              const resolveAnchorToLetter = (
                anchor: string | number
              ): string | null => {
                if (typeof anchor === "number")
                  return indexToLetter(anchor - 1);
                const text = String(anchor).trim();
                if (/^[A-Za-z]+$/.test(text)) return text.toUpperCase();
                if (/^\d+$/.test(text))
                  return indexToLetter(parseInt(text, 10) - 1);
                const headers: string[] = (targetRegion as any).headers || [];
                const headerIdx = headers.findIndex(
                  (h) => String(h).toLowerCase() === text.toLowerCase()
                );
                if (headerIdx >= 0) {
                  const match = (targetRegion as any).range.match(
                    /([A-Z]+)\d+:([A-Z]+)\d+/
                  );
                  if (match) {
                    const startLetter = match[1];
                    const startIndex = letterToIndex(startLetter);
                    return indexToLetter(startIndex + headerIdx);
                  }
                }
                return null;
              };
              let targetColumnLetter: string;
              {
                // Default: insert immediately AFTER the table's last column
                const match = (targetRegion as any).range.match(
                  /([A-Z]+)\d+:([A-Z]+)\d+/
                );
                if (!match) throw new Error("Invalid region range");
                const endLetter = match[2];
                const endIndex = letterToIndex(endLetter);
                targetColumnLetter = indexToLetter(endIndex + 1);
                console.log(
                  `📍 Positioning: Default after table end → ${targetColumnLetter}`
                );
              }

              const headerRow = parseInt(targetRegion.range.match(/\d+/)![0]);
              const worksheet = univerAPI.getActiveWorkbook().getActiveSheet();

              // Always insert a new column at the resolved position (keeps grid consistent)
              {
                const targetColIndex = letterToIndex(targetColumnLetter);
                console.log(
                  `📍 Inserting new column at index ${targetColIndex} (letter ${targetColumnLetter})`
                );

                // Use correct Univer.js API: insertColumns (plural) not insertColumn
                if (typeof worksheet.insertColumns === "function") {
                  worksheet.insertColumns(targetColIndex, 1);
                  console.log(
                    `✅ Successfully inserted column using insertColumns API`
                  );
                } else {
                  console.warn(
                    "⚠️ insertColumns method not available, column insertion skipped"
                  );
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
                }:${targetColumnLetter}$${
                  headerRow + targetRegion.rowCount
                }`.replace(/\$+/g, "");
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

// CollapsibleMarkdown now imported from ./chat/collapsible-markdown

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
