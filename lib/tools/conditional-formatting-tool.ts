/**
 * Enhanced Conditional Formatting Tool using proper Univer Rule Builder Pattern
 * 
 * This tool follows the official Univer API documentation for conditional formatting:
 * - Uses method chaining with rule builders
 * - Properly handles different condition types
 * - Applies formatting through official Univer APIs
 */

import { createSimpleTool } from "../tool-executor";
import type { UniversalToolContext } from "../tool-executor";

export const ConditionalFormattingTool = createSimpleTool(
  {
    name: "conditional_formatting",
    description: "Add conditional formatting rules using proper Univer rule builder pattern",
    category: "format",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (context: UniversalToolContext, params: {
    range: string;
    ruleType?: string;
    min?: number;
    max?: number;
    equals?: number;
    contains?: string;
    startsWith?: string;
    endsWith?: string;
    formula?: string;
    format?: { 
      backgroundColor?: string; 
      fontColor?: string;
      bold?: boolean;
      italic?: boolean;
    };
  }) => {
    const { 
      range, 
      ruleType = "number_lt", 
      format = {},
      min,
      max,
      equals,
      contains,
      startsWith,
      endsWith,
      formula
    } = params;

    console.log(`🎨 ConditionalFormattingTool: Applying ${ruleType} to ${range}`);

    try {
      // Get the range object
      const fRange = context.fWorksheet.getRange(range);
      
      // Start building the conditional formatting rule
      let ruleBuilder = context.fWorksheet.newConditionalFormattingRule();

      // Apply condition based on rule type
      switch (ruleType) {
        case "number_lt":
          const threshold = min !== undefined ? min : 0;
          ruleBuilder = ruleBuilder.whenNumberLessThan(threshold);
          console.log(`🔢 Applying number less than ${threshold}`);
          break;

        case "number_gt":
          ruleBuilder = ruleBuilder.whenNumberGreaterThan(min || 0);
          console.log(`🔢 Applying number greater than ${min || 0}`);
          break;

        case "number_gte":
          ruleBuilder = ruleBuilder.whenNumberGreaterThanOrEqualTo(min || 0);
          break;

        case "number_lte":
          ruleBuilder = ruleBuilder.whenNumberLessThanOrEqualTo(max || 0);
          break;

        case "number_eq":
          ruleBuilder = ruleBuilder.whenNumberEqualTo(equals || 0);
          break;

        case "number_neq":
          ruleBuilder = ruleBuilder.whenNumberNotEqualTo(equals || 0);
          break;

        case "number_between":
          if (min !== undefined && max !== undefined) {
            ruleBuilder = ruleBuilder.whenNumberBetween(min, max);
          } else {
            throw new Error("Both min and max values required for number_between rule");
          }
          break;

        case "text_contains":
          if (!contains) {
            throw new Error("'contains' parameter required for text_contains rule");
          }
          ruleBuilder = ruleBuilder.whenTextContains(contains);
          break;

        case "text_not_contains":
          if (!contains) {
            throw new Error("'contains' parameter required for text_not_contains rule");
          }
          ruleBuilder = ruleBuilder.whenTextDoesNotContain(contains);
          break;

        case "text_starts_with":
          if (!startsWith) {
            throw new Error("'startsWith' parameter required for text_starts_with rule");
          }
          ruleBuilder = ruleBuilder.whenTextStartsWith(startsWith);
          break;

        case "text_ends_with":
          if (!endsWith) {
            throw new Error("'endsWith' parameter required for text_ends_with rule");
          }
          ruleBuilder = ruleBuilder.whenTextEndsWith(endsWith);
          break;

        case "not_empty":
          ruleBuilder = ruleBuilder.whenCellNotEmpty();
          break;

        case "empty":
          ruleBuilder = ruleBuilder.whenCellEmpty();
          break;

        case "formula":
          if (!formula) {
            throw new Error("'formula' parameter required for formula rule");
          }
          ruleBuilder = ruleBuilder.whenFormulaSatisfied(formula);
          break;

        case "color_scale":
          // For color scale, create a gradient rule
          ruleBuilder = ruleBuilder
            .whenCellNotEmpty()
            .setGradientMinpoint("#FF0000")   // Red for minimum
            .setGradientMidpoint("#FFFF00")   // Yellow for middle  
            .setGradientMaxpoint("#00FF00");  // Green for maximum
          break;

        case "data_bar":
          // Create data bar visualization
          ruleBuilder = ruleBuilder
            .whenCellNotEmpty()
            .setDataBar("#4472C4");  // Blue data bars
          break;

        case "unique":
          ruleBuilder = ruleBuilder.whenCellValueUnique();
          break;

        case "duplicate":
          ruleBuilder = ruleBuilder.whenCellValueDuplicate();
          break;

        default:
          console.warn(`⚠️ Unknown rule type: ${ruleType}, defaulting to number_lt 0`);
          ruleBuilder = ruleBuilder.whenNumberLessThan(0);
      }

      // Apply formatting styles (skip for special rules like color_scale, data_bar)
      if (!['color_scale', 'data_bar'].includes(ruleType)) {
        if (format.backgroundColor) {
          ruleBuilder = ruleBuilder.setBackground(format.backgroundColor);
        }
        if (format.fontColor) {
          ruleBuilder = ruleBuilder.setFontColor(format.fontColor);
        }
        if (format.bold) {
          ruleBuilder = ruleBuilder.setBold(format.bold);
        }
        if (format.italic) {
          ruleBuilder = ruleBuilder.setItalic(format.italic);
        }

        // Default formatting for negative values if no format specified
        if (!format.backgroundColor && !format.fontColor && ruleType === "number_lt") {
          ruleBuilder = ruleBuilder.setFontColor("#FF0000"); // Red for negative
        }
      }

      // Set the range and build the rule
      const rule = ruleBuilder
        .setRanges([fRange.getRange()])
        .build();

      if (!rule) {
        throw new Error("Failed to build conditional formatting rule");
      }

      // Apply the rule to the worksheet
      const success = await context.fWorksheet.addConditionalFormattingRule(rule);
      
      if (!success) {
        throw new Error("Failed to add conditional formatting rule to worksheet");
      }

      console.log(`✅ ConditionalFormattingTool: Successfully applied ${ruleType} to ${range}`);

      return {
        range,
        ruleType,
        format,
        parameters: { min, max, equals, contains, startsWith, endsWith, formula },
        message: `Applied ${ruleType} conditional formatting to ${range}`,
        success: true
      };

    } catch (error) {
      console.error(`❌ ConditionalFormattingTool failed:`, error);
      
      // Fallback: Try using direct range formatting if rule builder fails
      if (format.backgroundColor || format.fontColor) {
        try {
          console.log(`🔄 Trying fallback range formatting...`);
          const fRange = context.fWorksheet.getRange(range);
          
          if (format.backgroundColor) {
            fRange.setBackground(format.backgroundColor);
          }
          if (format.fontColor) {
            fRange.setFontColor(format.fontColor);
          }
          
          return {
            range,
            ruleType: "fallback_direct_format",
            format,
            message: `Applied direct formatting to ${range} (fallback method)`,
            success: true,
            fallback: true
          };
        } catch (fallbackError) {
          console.error(`❌ Fallback formatting also failed:`, fallbackError);
        }
      }

      throw error;
    }
  }
);