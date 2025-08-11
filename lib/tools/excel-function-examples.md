# Excel Function Tools - Usage Examples

The Excel Function Tools provide comprehensive support for all common Excel functions without any hardcoded assumptions. Here are examples of how to use each function:

## Mathematical Functions

### SUM
```javascript
// Sum a specific column
{
  "formula_type": "sum",
  "target_column": "Sales"
}

// Sum with custom range
{
  "formula_type": "sum",
  "source_range": "B2:B10"
}
```

### AVERAGE
```javascript
{
  "formula_type": "average",
  "target_column": "Ratings"
}
```

### COUNT
```javascript
{
  "formula_type": "count",
  "target_column": "Items"
}
```

### MAX/MIN
```javascript
{
  "formula_type": "max",
  "target_column": "Prices"
}

{
  "formula_type": "min",
  "target_column": "Costs"
}
```

## Conditional Functions

### IF
```javascript
{
  "formula_type": "if",
  "target_column": "Sales",
  "condition": "B2>1000",
  "output_cell": "C2"
}
```

### SUMIF
```javascript
{
  "formula_type": "sumif",
  "target_column": "Revenue",
  "condition": ">500"
}
```

### COUNTIF
```javascript
{
  "formula_type": "countif",
  "target_column": "Status",
  "condition": "Complete"
}
```

## Lookup Functions

### VLOOKUP
```javascript
{
  "formula_type": "vlookup",
  "lookup_value": "A2",
  "lookup_table": "Products!A:D",
  "target_column": "Price"
}
```

## Text Functions

### CONCATENATE
```javascript
{
  "formula_type": "concatenate",
  "target_column": "FirstName",
  "source_range": "A2:B2"
}
```

### LEFT/RIGHT/MID
```javascript
{
  "formula_type": "left",
  "target_column": "ProductCode",
  "parameters": { "numChars": 3 }
}

{
  "formula_type": "right",
  "target_column": "SKU",
  "parameters": { "numChars": 4 }
}

{
  "formula_type": "mid",
  "target_column": "Description",
  "parameters": { "startPos": 3, "numChars": 5 }
}
```

### Text Manipulation
```javascript
{
  "formula_type": "upper",
  "target_column": "Name"
}

{
  "formula_type": "lower",
  "target_column": "Email"
}

{
  "formula_type": "trim",
  "target_column": "Address"
}

{
  "formula_type": "len",
  "target_column": "Description"
}
```

## Mathematical Functions

### ROUND/ABS/SQRT
```javascript
{
  "formula_type": "round",
  "target_column": "Price",
  "parameters": { "decimals": 2 }
}

{
  "formula_type": "abs",
  "target_column": "Variance"
}

{
  "formula_type": "sqrt",
  "target_column": "Area"
}

{
  "formula_type": "power",
  "target_column": "Base",
  "parameters": { "power": 3 }
}
```

## Bulk Formula Application

Apply multiple formulas at once:

```javascript
{
  "formulas": [
    {
      "type": "sum",
      "target_column": "Sales",
      "output_cell": "B15"
    },
    {
      "type": "average",
      "target_column": "Ratings",
      "output_cell": "C15"
    },
    {
      "type": "count",
      "target_column": "Items",
      "output_cell": "D15"
    }
  ]
}
```

## Key Features

✅ **No Hardcoding**: All ranges are detected intelligently based on actual table structure
✅ **Context Aware**: Automatically finds optimal placement for formula results  
✅ **Type Intelligence**: Prefers numeric columns for math functions, text columns for string functions
✅ **Spatial Analysis**: Uses spatial mapping to avoid overwriting existing data
✅ **Error Handling**: Graceful fallbacks when columns or ranges aren't found
✅ **Batch Processing**: Efficient bulk formula application

## Supported Functions (20+)

- **Math**: SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ABS, SQRT, POWER
- **Logic**: IF, SUMIF, COUNTIF  
- **Lookup**: VLOOKUP
- **Text**: CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM

All functions work without requiring hardcoded cell references or ranges!