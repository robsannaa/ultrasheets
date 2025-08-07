"use client";

import { useRef, useEffect } from "react";
import * as React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Send, Loader2, CheckCircle2, Wrench } from "lucide-react";
import { cn } from "@/lib/utils";
import { applyFormatting } from "@/lib/univer-helpers";
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
  const [isHydrated, setIsHydrated] = React.useState(false);

  // Avoid SSR/CSR markup mismatch by rendering message list only after hydration
  useEffect(() => {
    setIsHydrated(true);
  }, []);

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

  // Track assistant texts we've already rendered during this pass to avoid duplicates
  const seenAssistantTexts = React.useRef<Set<string>>(new Set());

  // During SSR, render an empty container to keep markup stable
  if (!isHydrated) {
    return (
      <ScrollArea className="h-full px-1 py-4" ref={scrollAreaRef}>
        <div className="space-y-4" />
      </ScrollArea>
    );
  }

  return (
    <ScrollArea className="h-full px-1 py-4" ref={scrollAreaRef}>
      <div className="space-y-4">
        {mergedMessages.map((message) => {
          // Skip rendering empty assistant frames (no text and no tool badges)
          const hasParts = Array.isArray(message.parts);
          const hasTextPart = hasParts
            ? message.parts.some(
                (p: any) => p.type === "text" && typeof p.text === "string" && p.text.trim().length > 0
              )
            : false;
          const hasToolPart = hasParts
            ? message.parts.some((p: any) => p.type === "tool-invocation")
            : false;
          const hasContent =
            (typeof message.content === "string" && message.content.trim().length > 0) ||
            hasTextPart ||
            hasToolPart;
          if (message.role === "assistant" && !hasContent) return null;

          return (
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
                      // Render only the final assistant text part (post-tool result),
                      // instead of concatenating all text parts which can duplicate intent text.
                      const lastTextPart = [...message.parts]
                        .reverse()
                        .find(
                          (p: any) =>
                            p.type === "text" && typeof p.text === "string"
                        );
                      const text = lastTextPart?.text ?? "";
                      // Deduplicate identical assistant texts within the same render pass
                      const normalized = (text || "")
                        .replace(/\s+/g, " ")
                        .trim();
                      const shouldHideText =
                        message.role === "assistant" &&
                        normalized.length > 0 &&
                        seenAssistantTexts.current.has(normalized);
                      if (
                        message.role === "assistant" &&
                        normalized.length > 0 &&
                        !shouldHideText
                      ) {
                        seenAssistantTexts.current.add(normalized);
                      }
                      return text && !shouldHideText ? (
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
                        // When parts exist, we do not render message.content fallback to avoid duplicates
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
        const univerAPI = (window as any).univerAPI;
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
    default:
      return "Processing";
  }
}
