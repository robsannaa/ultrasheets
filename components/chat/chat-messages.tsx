"use client";

import * as React from "react";
import { useEffect, useRef } from "react";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { cn } from "@/lib/utils";
import { ToolBadge } from "@/components/ToolBadge";
import { CollapsibleMarkdown } from "./collapsible-markdown";
import { Clock, CheckCircle, XCircle, AlertTriangle } from "lucide-react";

const consultingVerbs = [
  "Analyzing",
  "Arbitraging",
  "Auditing",
  "Benchmarking",
  "Budgeting",
  "Capitalizing",
  "Closing",
  "Consolidating",
  "Deleveraging",
  "De-risking",
  "Diversifying",
  "Drafting",
  "Engaging",
  "Evaluating",
  "Executing",
  "Expanding",
  "Facilitating",
  "Forecasting",
  "Hedging",
  "Implementing",
  "Integrating",
  "Investing",
  "Leveraging",
  "Mapping",
  "Maximizing",
  "Measuring",
  "Mediating",
  "Mitigating",
  "Modeling",
  "Negotiating",
  "Optimizing",
  "Outsourcing",
  "Planning",
  "Positioning",
  "Prioritizing",
  "Procuring",
  "Processing",
  "Rebalancing",
  "Rebranding",
  "Refinancing",
  "Refining",
  "Researching",
  "Restructuring",
  "Scaling",
  "Securing",
  "Simplifying",
  "Standardizing",
  "Streamlining",
  "Strengthening",
  "Structuring",
  "Synthesizing",
  "Targeting",
  "Testing",
  "Tracking",
  "Trading",
  "Transforming",
  "Validating",
  "Valuating",
  "Visualizing",
];

export function ChatMessages({
  messages,
  isLoading,
  aiPerformanceMode = false,
}: {
  messages: any[];
  isLoading: boolean;
  aiPerformanceMode?: boolean;
}) {
  const scrollAreaRef = useRef<HTMLDivElement>(null);

  // Loading verb + animated dots state
  const [loadingVerb, setLoadingVerb] = React.useState<string>(
    consultingVerbs[Math.floor(Math.random() * consultingVerbs.length)]
  );
  const [dotCount, setDotCount] = React.useState<number>(1);

  // Rotate word every 3s and dots every 500ms while loading
  useEffect(() => {
    if (!isLoading) return;

    const verbTimer = setInterval(() => {
      const next =
        consultingVerbs[Math.floor(Math.random() * consultingVerbs.length)];
      setLoadingVerb(next);
    }, 3000);

    const dotsTimer = setInterval(() => {
      setDotCount((c) => (c % 3) + 1);
    }, 500);

    return () => {
      clearInterval(verbTimer);
      clearInterval(dotsTimer);
    };
  }, [isLoading]);

  useEffect(() => {
    if (scrollAreaRef.current) {
      const scrollContainer = scrollAreaRef.current.querySelector(
        "[data-radix-scroll-area-viewport]"
      );
      if (scrollContainer) {
        (scrollContainer as HTMLElement).scrollTop = (
          scrollContainer as HTMLElement
        ).scrollHeight;
      }
    }
  }, [messages]);

  const mergedMessages = messages;

  return (
    <ScrollArea className="h-full p-4" ref={scrollAreaRef}>
      <div className="space-y-3">
        {/* AI Performance indicator at top */}
        {aiPerformanceMode && messages.length > 1 && (
          <div className="text-xs text-center text-muted-foreground py-2 border-b">
            <div className="flex items-center justify-center gap-2">
              <div className="h-2 w-2 bg-green-500 rounded-full animate-pulse" />
              AI Performance Mode Active
            </div>
          </div>
        )}

        {mergedMessages.map((message, idx) => {
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
          const isLast = idx === mergedMessages.length - 1;
          if (
            message.role === "assistant" &&
            !hasContent &&
            !(isLoading && isLast)
          )
            return null;

          return (
            <div
              key={message.id}
              className={cn(
                message.role === "assistant"
                  ? "rounded-xl p-3 transition-all duration-300 bg-muted/50"
                  : message.role === "user"
                  ? "rounded-xl p-3 transition-all duration-300 bg-primary/5 ml-6"
                  : "rounded-full px-3 py-1.5 transition-all duration-300 text-center bg-orange-50/80 dark:bg-orange-950/30 border border-orange-200/60 dark:border-orange-800/40 mr-auto max-w-fit"
              )}
            >
              <div className="text-sm leading-relaxed text-foreground">
                {message.parts ? (
                  <>
                    {message.parts.map((part: any) => {
                      if (part.type === "tool-invocation") {
                        const callId = part.toolInvocation.toolCallId;
                        const toolName = part.toolInvocation.toolName;
                        const state = part.toolInvocation.state;
                        const args = part.toolInvocation.args;
                        const label = describeTool(toolName, args);

                        if (state === "call") {
                          return (
                            <ToolBadge
                              key={callId}
                              label={label}
                              toolName={toolName}
                              state="call"
                            />
                          );
                        } else if (state === "result") {
                          const result = part.toolInvocation.result;
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
                              <div className="mt-1 text-xs text-muted-foreground break-words space-y-1">
                                <div className="flex items-center gap-2">
                                  <span>
                                    Action:{" "}
                                    {result?.clientSideAction?.type || "N/A"}
                                  </span>
                                  {result?.success !== false ? (
                                    <CheckCircle className="h-3 w-3 text-green-500" />
                                  ) : (
                                    <XCircle className="h-3 w-3 text-red-500" />
                                  )}
                                </div>

                                {/* Enhanced result information */}
                                {aiPerformanceMode && result?._metadata && (
                                  <div className="text-xs text-muted-foreground">
                                    <div className="flex items-center gap-1">
                                      <Clock className="h-3 w-3" />
                                      <span>
                                        {result._metadata.executionTime}ms
                                      </span>
                                      {result._metadata.attempts > 1 && (
                                        <span className="text-amber-500">
                                          ({result._metadata.attempts} attempts)
                                        </span>
                                      )}
                                      {result._metadata.cacheHit && (
                                        <span className="text-green-500">
                                          (cached)
                                        </span>
                                      )}
                                    </div>
                                    {result._metadata.tablesAnalyzed > 0 && (
                                      <div className="text-xs">
                                        Tables analyzed:{" "}
                                        {result._metadata.tablesAnalyzed}
                                      </div>
                                    )}
                                  </div>
                                )}

                                {/* Show any error information */}
                                {result?.error && (
                                  <div className="flex items-center gap-1 text-red-500">
                                    <AlertTriangle className="h-3 w-3" />
                                    <span className="text-xs">
                                      {result.error}
                                    </span>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        }
                        return null;
                      }
                      return null;
                    })}
                    {(() => {
                      try {
                        const text = message.parts
                          .filter(
                            (p: any) =>
                              p.type === "text" && typeof p.text === "string"
                          )
                          .map((p: any) => p.text)
                          .join("");
                        if (text && text.trim().length > 0) {
                          return (
                            <CollapsibleMarkdown
                              key="assistant-text"
                              text={text}
                            />
                          );
                        }
                        if (
                          message.role === "assistant" &&
                          isLast &&
                          isLoading
                        ) {
                          return (
                            <div
                              key="assistant-loading"
                              className="text-sm leading-relaxed text-foreground"
                              role="status"
                              aria-live="polite"
                            >
                              {loadingVerb} {".".repeat(dotCount)}
                            </div>
                          );
                        }
                        return null;
                      } catch {
                        return null;
                      }
                    })()}
                  </>
                ) : message.role === "assistant" ? (
                  (() => {
                    const content = String(message.content || "");
                    if (content && content.trim().length > 0) {
                      return <CollapsibleMarkdown text={content} />;
                    }
                    if (isLast && isLoading) {
                      return (
                        <div
                          className="text-sm leading-relaxed text-foreground"
                          role="status"
                          aria-live="polite"
                        >
                          {loadingVerb} {".".repeat(dotCount)}
                        </div>
                      );
                    }
                    return null;
                  })()
                ) : (
                  message.content
                )}
              </div>
            </div>
          );
        })}
      </div>
    </ScrollArea>
  );
}

function describeTool(toolName: string, args: any): string {
  switch (toolName) {
    case "list_columns":
      return "List columns";
    case "calculate_total":
      return `Total for ${args?.column ?? "column"}`;
    case "add_totals":
      return `Add totals${
        args?.columns
          ? ` for ${args.columns.join(", ")}`
          : " for specified columns"
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
