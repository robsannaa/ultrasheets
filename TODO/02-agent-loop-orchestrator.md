# 02. Agent Loop Orchestrator (CRITICAL)

## What?

Implement the multi-step agent loop that allows the LLM to call multiple tools in sequence until reaching a final answer, as specified in Section 4 of the integration spec.

## Where?

- **File**: `app/api/chat/route.ts` (modify existing)
- **Dependencies**: All tools from TODO-01
- **Configuration**: Update `maxSteps` and add proper loop logic

## How?

### 2.1 Update Chat Route

```typescript
// Current: Single tool execution
// Target: Multi-step loop with tool chaining

const result = streamText({
  model: openai("gpt-4o-mini"),
  system: `You are a spreadsheet analysis assistant for finance professionals.
You have function-calling access to precise tools.
Always call a tool for any answer that requires data from the sheet.
Never fabricate numbers. Summarize insights clearly.

AVAILABLE TOOLS:
- list_columns: Get column names and row count
- create_pivot_table: Create pivot tables with grouping
- calculate_total: Calculate totals for columns
- generate_chart: Generate charts from data
- financial_intelligence: Advanced financial analysis

EXECUTION STRATEGY:
1. Use list_columns to understand data structure
2. Use appropriate tools for calculations
3. Create charts for visualization
4. Provide clear summary of results`,
  messages,
  temperature: 0.1, // Lower for deterministic operations
  maxSteps: 10, // Allow multiple tool calls
  tools: {
    // All tools from TODO-01
  },
});
```

### 2.2 Add Tool Result Handling

```typescript
// In chat-sidebar.tsx, enhance tool execution
useEffect(() => {
  const executeToolCalls = async () => {
    const lastMessage = messages[messages.length - 1];
    if (lastMessage?.role === "assistant" && lastMessage.parts) {
      for (const part of lastMessage.parts) {
        if (part.type === "tool-invocation") {
          // Execute tool and wait for result
          // Handle multiple tool calls in sequence
          // Update UI with progress
        }
      }
    }
  };

  executeToolCalls();
}, [messages]);
```

### 2.3 Add Progress Tracking

```typescript
// Add to chat-sidebar.tsx
const [toolExecutionProgress, setToolExecutionProgress] = useState<{
  current: number;
  total: number;
  description: string;
} | null>(null);
```

## How to Test?

1. **Multi-step Scenarios**:

   ```bash
   # Test: "Analyze the sales data"
   # Should call: list_columns → create_pivot_table → generate_chart → summary
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"Analyze the sales data"}]}'
   ```

2. **Tool Chaining Tests**:

   - Verify tools can reference previous tool results
   - Test error handling in multi-step sequences
   - Validate final summary includes all results

3. **Performance Tests**:
   - Test with 5+ tool calls in sequence
   - Verify timeout handling
   - Test memory usage with large datasets

## Important Dependencies to Not Break

- **Existing single-tool calls** - Must still work
- **Message history** - Preserve conversation context
- **Error handling** - Don't break on tool failures
- **UI responsiveness** - Keep chat interface responsive

## Dependencies That Will Work Thanks to This

- **Complex analysis workflows** - Multi-step financial analysis
- **Chart generation** - Can use pivot table results
- **User experience** - Single request can perform complex operations
- **Future tools** - Framework for adding more complex tools

## Implementation Strategy (Least Disruptive)

1. **Start with existing tools** - Don't add new tools yet
2. **Test with simple scenarios** - 2-3 tool calls first
3. **Add progress indicators** - Keep users informed
4. **Implement error recovery** - Handle tool failures gracefully
5. **Optimize performance** - Cache tool results when possible

## Priority Order

1. Update system prompt and temperature
2. Implement multi-step execution logic
3. Add progress tracking UI
4. Test with existing tools
5. Optimize performance and error handling
