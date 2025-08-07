"use client";

import { useRef, useEffect } from "react";
import * as React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Send, Loader2 } from "lucide-react";
import { cn } from "@/lib/utils";
import { useChat } from "ai/react";

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

  return (
    <ScrollArea className="h-full px-1 py-4" ref={scrollAreaRef}>
      <div className="space-y-4">
        {messages.map((message) => (
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
                "max-w-[70%] rounded-lg p-3 text-sm whitespace-pre-wrap",
                message.role === "user"
                  ? "bg-primary text-primary-foreground"
                  : "bg-muted"
              )}
            >
              {message.parts
                ? message.parts.map((part: any, index: number) => {
                    switch (part.type) {
                      case "text":
                        return <span key={index}>{part.text}</span>;
                      case "tool-invocation": {
                        const callId = part.toolInvocation.toolCallId;
                        const toolName = part.toolInvocation.toolName;
                        const state = part.toolInvocation.state;
                        const args = part.toolInvocation.args;

                        if (state === "call") {
                          switch (toolName) {
                            case "get_sheet_context":
                              return (
                                <div
                                  key={callId}
                                  className="text-slate-600 text-sm italic"
                                >
                                  Analyzing spreadsheet...
                                </div>
                              );
                            case "list_columns":
                              return (
                                <div
                                  key={callId}
                                  className="text-slate-600 text-sm italic"
                                >
                                  Getting column information...
                                </div>
                              );
                            case "calculate_total":
                              return (
                                <div
                                  key={callId}
                                  className="text-blue-600 text-sm italic"
                                >
                                  Calculating total for column {args.column}...
                                </div>
                              );
                            case "create_pivot_table":
                              return (
                                <div
                                  key={callId}
                                  className="text-blue-600 text-sm italic"
                                >
                                  Creating pivot table grouped by {args.groupBy}
                                  ...
                                </div>
                              );
                            case "generate_chart":
                              return (
                                <div
                                  key={callId}
                                  className="text-blue-600 text-sm italic"
                                >
                                  Generating {args.chart_type} chart...
                                </div>
                              );
                            case "switch_sheet":
                              return (
                                <div
                                  key={callId}
                                  className="text-blue-600 text-sm italic"
                                >
                                  {args.action === "switch"
                                    ? "Switching to"
                                    : "Analyzing"}{" "}
                                  sheet {args.sheetName}...
                                </div>
                              );
                            case "financial_intelligence":
                              return (
                                <div
                                  key={callId}
                                  className="text-blue-600 text-sm italic"
                                >
                                  Processing: {args.user_request}...
                                </div>
                              );
                            case "ask_for_range":
                              return (
                                <div
                                  key={callId}
                                  className="text-amber-600 text-sm italic"
                                >
                                  Need range specification...
                                </div>
                              );
                            case "add_filter":
                              return (
                                <div
                                  key={callId}
                                  className="text-blue-600 text-sm italic"
                                >
                                  Applying filter...
                                </div>
                              );
                            default:
                              return (
                                <div
                                  key={callId}
                                  className="text-slate-600 text-sm italic"
                                >
                                  Processing...
                                </div>
                              );
                          }
                        } else if (state === "result") {
                          // Only show result for important operations, hide technical details
                          const result = part.toolInvocation.result;
                          if (toolName === "get_sheet_context") {
                            return null; // Hide sheet context results - too technical
                          }

                          if (toolName === "financial_intelligence") {
                            return (
                              <div
                                key={callId}
                                className="text-green-700 text-sm font-medium"
                              >
                                ✓ Calculation complete
                              </div>
                            );
                          }

                          if (toolName === "list_columns") {
                            return (
                              <div
                                key={callId}
                                className="text-green-700 text-sm"
                              >
                                ✓ Found {result.columns?.length || 0} columns
                              </div>
                            );
                          }

                          if (toolName === "calculate_total") {
                            return (
                              <div
                                key={callId}
                                className="text-green-700 text-sm"
                              >
                                ✓ Total calculated
                              </div>
                            );
                          }

                          if (toolName === "create_pivot_table") {
                            return (
                              <div
                                key={callId}
                                className="text-green-700 text-sm"
                              >
                                ✓ Pivot table created
                              </div>
                            );
                          }

                          if (toolName === "generate_chart") {
                            return (
                              <div
                                key={callId}
                                className="text-green-700 text-sm"
                              >
                                ✓ Chart generated
                              </div>
                            );
                          }

                          if (toolName === "switch_sheet") {
                            return (
                              <div
                                key={callId}
                                className="text-green-700 text-sm"
                              >
                                ✓ Sheet {result.action || "analyzed"}
                              </div>
                            );
                          }

                          if (toolName === "add_filter") {
                            return (
                              <div
                                key={callId}
                                className={
                                  result.success
                                    ? "text-green-700 text-sm"
                                    : "text-amber-600 text-sm"
                                }
                              >
                                {result.success
                                  ? "✓ Filter applied"
                                  : "⚠️ Filter simulation"}
                              </div>
                            );
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

                          // Show clean success messages for other operations
                          return (
                            <div
                              key={callId}
                              className="text-green-700 text-sm"
                            >
                              {typeof result === "string" &&
                              result.includes("Successfully")
                                ? "✓ Complete"
                                : "✓ Done"}
                            </div>
                          );
                        }
                        return null;
                      }
                      default:
                        return null;
                    }
                  })
                : message.content}
            </div>
          </div>
        ))}
        {isLoading && (
          <div className="flex items-start gap-3">
            <Avatar className="w-8 h-8">
              <AvatarFallback>A</AvatarFallback>
            </Avatar>
            <div className="bg-muted rounded-lg p-3 text-sm">
              <Loader2 className="w-4 h-4 animate-spin" />
            </div>
          </div>
        )}
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

// Simple function to extract basic workbook data from Univer
function extractWorkbookData() {
  try {
    if (typeof window !== "undefined" && (window as any).univerAPI) {
      const univerAPI = (window as any).univerAPI;
      const workbook = univerAPI.getActiveWorkbook();

      if (!workbook) {
        return null;
      }

      const activeSheet = workbook.getActiveSheet();
      const sheetSnapshot = activeSheet.getSheet().getSnapshot();
      const cellData = sheetSnapshot.cellData || {};

      // Find headers in first few rows
      let headers: string[] = [];
      for (let row = 0; row < 5; row++) {
        const rowData = cellData[row] || {};
        const rowHeaders: string[] = [];

        for (let col = 0; col < 20; col++) {
          const cell = rowData[col];
          if (cell && typeof cell.v === "string" && cell.v.trim()) {
            rowHeaders.push(cell.v.trim());
          } else {
            break;
          }
        }

        if (rowHeaders.length > headers.length) {
          headers = rowHeaders;
        }
      }

      // Count total cells with data
      let totalCells = 0;
      for (const rowIndex in cellData) {
        for (const colIndex in cellData[rowIndex]) {
          const cell = cellData[rowIndex][colIndex];
          if (
            cell &&
            cell.v !== undefined &&
            cell.v !== null &&
            cell.v !== ""
          ) {
            totalCells++;
          }
        }
      }

      return {
        sheets: [
          {
            name: sheetSnapshot.name || "Sheet1",
            isActive: true,
            headers,
            structure: {
              totalCells,
              dataRows: Object.keys(cellData).length,
            },
          },
        ],
      };
    }

    return null;
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

  // Listen for tool results and execute them on the frontend
  useEffect(() => {
    const executeClientSideActions = async () => {
      const lastMessage = messages[messages.length - 1];
      if (lastMessage?.role === "assistant" && lastMessage.parts) {
        for (const part of lastMessage.parts) {
          if (
            part.type === "tool-invocation" &&
            part.toolInvocation.state === "result"
          ) {
            const result = part.toolInvocation.result;

            if (result && result.clientSideAction) {
              const action = result.clientSideAction;

              try {
                if (
                  action.type === "executeUniverTool" &&
                  window.executeUniverTool
                ) {
                  await window.executeUniverTool(
                    action.toolName,
                    action.params
                  );
                } else if (action.type === "formatCells") {
                  await formatCells(action);
                }
              } catch (error) {
                console.error("Failed to execute client-side action:", error);
              }
            }
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
            const alignmentValue = action.textAlign;
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
                    const alignmentMap = { left: 1, center: 2, right: 3 };
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
                  const alignmentMap = { left: 1, center: 2, right: 3 };
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
                const alignmentMap = { left: 1, center: 2, right: 3 };
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

  const isLoading = status === "submitted" || status === "streaming";

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
