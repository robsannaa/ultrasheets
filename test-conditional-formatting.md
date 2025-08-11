# Conditional Formatting Test Plan

## Root Cause Identified
The original ConditionalFormattingTool was using the wrong API approach:
- **Wrong approach**: Calling `setBackgroundColor()` and `setFontColor()` directly on individual cells
- **Correct approach**: Using Univer.js's dedicated conditional formatting API with `newConditionalFormattingRule()`

## Fix Implementation
The updated ConditionalFormattingTool now:

1. **Uses proper CF API**: Creates rules using `context.fWorksheet.newConditionalFormattingRule()`
2. **Sets ranges correctly**: Uses `rule.setRanges([fRange])` to define the target range
3. **Applies conditions properly**: Uses methods like `whenNumberGreaterThan()`, `whenTextContains()`, etc.
4. **Sets formatting on rule**: Uses `rule.setBackground()` and `rule.setFontColor()` on the rule itself
5. **Builds and applies rule**: Uses `rule.build()` and `setConditionalFormattingRule(builtRule)`
6. **Has fallback**: Falls back to direct cell formatting if CF API fails

## Expected Behavior Changes
- Visual conditional formatting should now appear in the spreadsheet interface
- Rules should be persistent and reactive to cell value changes
- Console should show detailed logging of the CF rule creation process

## API Method Changes
| Old Method | New Method | Purpose |
|------------|------------|---------|
| `cellRange.setBackgroundColor()` | `rule.setBackground()` | Set background color on rule instead of cell |
| `cellRange.setFontColor()` | `rule.setFontColor()` | Set font color on rule instead of cell |
| Direct formatting | `newConditionalFormattingRule()` | Create proper CF rule |

## Test Scenarios to Verify

1. **Greater Than Condition**
   - Create range with numbers (e.g., A1:A10)
   - Apply conditional formatting for values > 50 with red background
   - Expect: All values > 50 should have red background that updates dynamically

2. **Text Contains Condition**  
   - Create range with text values
   - Apply conditional formatting for cells containing "test" with blue background
   - Expect: Matching cells should have blue background

3. **Empty/Not Empty Conditions**
   - Create range with mixed empty and filled cells
   - Apply conditional formatting for non-empty cells with green background
   - Expect: Only non-empty cells should be highlighted

## Console Logging Added
The fixed implementation adds detailed console logging:
- `🎨 Creating conditional formatting rule for range: X`
- `🎨 Applied background color: Y` 
- `🎨 Built conditional formatting rule:` (with rule object)
- `🎨 Applied conditional formatting rule with ID: Z`

If CF API fails, it logs:
- `❌ Failed to apply conditional formatting:` (with error)
- `🔄 Attempting fallback to direct cell formatting...`

## Key Univer.js CF API Methods Used
- `newConditionalFormattingRule()` - Create new rule
- `setRanges([range])` - Set target ranges
- `whenNumberGreaterThan(value)` - Number greater than condition
- `whenNumberLessThan(value)` - Number less than condition  
- `whenTextContains(text)` - Text contains condition
- `whenCellNotEmpty()` - Cell not empty condition
- `whenCellEmpty()` - Cell empty condition
- `setBackground(color)` - Set background color on rule
- `setFontColor(color)` - Set font color on rule
- `setBold(true)` - Set bold formatting on rule
- `setItalic(true)` - Set italic formatting on rule
- `build()` - Build the rule
- `setConditionalFormattingRule(rule)` - Apply rule to worksheet