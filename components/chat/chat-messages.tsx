"use client";

import * as React from "react";
import { useEffect, useRef } from "react";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { cn } from "@/lib/utils";
import { ToolBadge } from "@/components/ToolBadge";
import { CollapsibleMarkdown } from "./collapsible-markdown";

export function ChatMessages({
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
        (scrollContainer as HTMLElement).scrollTop = (
          scrollContainer as HTMLElement
        ).scrollHeight;
      }
    }
  }, [messages]);

  const mergedMessages = messages;

  return (
    <ScrollArea
      className="h-full px-1 sm:px-2 py-2 sm:py-4"
      ref={scrollAreaRef}
    >
      <div className="space-y-3">
        {mergedMessages.map((message) => {
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
                              <div className="mt-1 text-xs text-muted-foreground break-words">
                                Action type: {result?.clientSideAction?.type}
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
                  </>
                ) : message.role === "assistant" ? (
                  <CollapsibleMarkdown text={String(message.content || "")} />
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
