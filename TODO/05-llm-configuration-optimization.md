# 05. LLM Configuration Optimization (MEDIUM)

## What?

Optimize the LLM configuration to match the integration spec requirements: lower temperature for deterministic operations, add max_tokens configuration, and enhance system prompt with schema awareness.

## Where?

- **File**: `app/api/chat/route.ts`
- **Dependencies**: All tools from previous TODOs
- **Configuration**: Model parameters and system prompt

## How?

### 5.1 Update Model Parameters

```typescript
// Current configuration
const result = streamText({
  model: openai("gpt-4o-mini"),
  system: `You are a professional financial analyst assistant...`,
  messages,
  temperature: 0.7, // TOO HIGH
  maxSteps: 5,
  tools: {
    /* tools */
  },
});

// Target configuration
const result = streamText({
  model: openai("gpt-4o-mini"),
  system: await generateSchemaAwareSystemPrompt(),
  messages,
  temperature: 0.1, // Deterministic for numeric accuracy
  maxTokens: 4096, // Plenty for multi-step loops
  maxSteps: 10, // Allow more tool calls
  tools: {
    /* all tools */
  },
});
```

### 5.2 Create Schema-Aware System Prompt

```typescript
// New function in app/api/chat/route.ts
async function generateSchemaAwareSystemPrompt(): Promise<string> {
  try {
    // Get current sheet context
    const univerService = UniverService.getInstance();
    const context = await univerService.getSheetContext();

    // Build schema-aware prompt
    let prompt = `You are a spreadsheet analysis assistant for finance professionals.
You have function-calling access to precise tools.
Always call a tool for any answer that requires data from the sheet.
Never fabricate numbers. Summarize insights clearly.

CURRENT SHEET CONTEXT:
- Workbook: ${context.workbookName}
- Sheet: ${context.worksheetName}
- Data Range: ${context.dataRange}
- Total Rows: ${context.totalRows}
- Total Columns: ${context.totalColumns}

AVAILABLE DATA STRUCTURE:`;

    if (context.tables.length > 0) {
      const table = context.tables[0];
      prompt += `
- Table Range: ${table.range}
- Headers: ${table.headers.join(", ")}
- Sample Data: ${table.dataPreview
        .slice(0, 3)
        .map((row) => row.join(" | "))
        .join("\n  ")}`;
    }

    prompt += `

AVAILABLE TOOLS:
- list_columns: Get column names and row count
- create_pivot_table: Create pivot tables with grouping and aggregation
- calculate_total: Calculate totals for specific columns
- generate_chart: Generate charts from data and insert to Dashboard sheet
- financial_intelligence: Advanced financial analysis

EXECUTION STRATEGY:
1. Use list_columns to understand data structure (if not already known)
2. Use appropriate tools for calculations and analysis
3. Create charts for visualization when requested
4. Provide clear summary of results with specific numbers

IMPORTANT:
- Always use actual data from the sheet
- Reference specific column names when available
- Provide precise numeric results
- Suggest visualizations when appropriate`;

    return prompt;
  } catch (error) {
    // Fallback to basic prompt if context unavailable
    return `You are a spreadsheet analysis assistant for finance professionals.
You have function-calling access to precise tools.
Always call a tool for any answer that requires data from the sheet.
Never fabricate numbers. Summarize insights clearly.`;
  }
}
```

### 5.3 Add Context Injection to Chat

```typescript
// In app/api/chat/route.ts, before streamText call
const systemPrompt = await generateSchemaAwareSystemPrompt();

const result = streamText({
  model: openai("gpt-4o-mini"),
  system: systemPrompt,
  messages,
  temperature: 0.1,
  maxTokens: 4096,
  maxSteps: 10,
  tools: {
    // All tools from previous TODOs
  },
});
```

## How to Test?

1. **Temperature Testing**:

   ```bash
   # Test deterministic responses
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"sum column B"}]}'
   # Should get consistent results across multiple calls
   ```

2. **Schema Awareness Testing**:

   ```bash
   # Test with specific column names
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"calculate total for Sales column"}]}'
   # Should reference actual column names from sheet
   ```

3. **Multi-step Testing**:
   ```bash
   # Test complex analysis
   curl -X POST /api/chat -d '{"messages":[{"role":"user","content":"analyze sales data and create charts"}]}'
   # Should execute multiple tools in sequence
   ```

## Important Dependencies to Not Break

- **Existing tool execution** - Don't break current tool functionality
- **Message handling** - Preserve existing chat flow
- **Error handling** - Maintain error response patterns
- **Context extraction** - Don't break UniverService.getSheetContext()

## Dependencies That Will Work Thanks to This

- **Multi-step analysis** - Better tool selection with context
- **User experience** - More accurate and relevant responses
- **Financial analysis** - Precise numeric calculations
- **Chart generation** - Better data selection for charts

## Implementation Strategy (Least Disruptive)

1. **Start with temperature change** - Test deterministic behavior
2. **Add maxTokens configuration** - Ensure sufficient response length
3. **Implement schema-aware prompt** - Add context gradually
4. **Test with existing tools** - Verify no regression
5. **Optimize prompt content** - Refine based on testing

## Priority Order

1. Lower temperature to 0.1
2. Add maxTokens: 4096 configuration
3. Increase maxSteps to 10
4. Implement basic schema-aware prompt
5. Add comprehensive context injection
6. Test with all tools
7. Optimize prompt content based on testing
