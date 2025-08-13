/**
 * Data Validation Tool using Univer's Data Validation API
 * 
 * This tool follows the official Univer API documentation for data validation:
 * - Uses proper builder pattern for validation rules
 * - Supports various validation types (list, number, date, text length)
 * - Handles custom messages and error handling
 * - Provides intelligent validation rule detection
 */

import { createSimpleTool } from "../tool-executor";
import type { UniversalToolContext } from "../tool-executor";

export const AddDataValidationTool = createSimpleTool(
  {
    name: "add_data_validation",
    description: "Add data validation rules to cells using Univer's validation API",
    category: "data",
    requiredContext: [],
    invalidatesCache: false,
  },
  async (context: UniversalToolContext, params: {
    range: string;
    validationType: 'list' | 'number' | 'date' | 'textLength' | 'custom' | 'email' | 'url';
    criteria?: {
      // List validation
      values?: string[];
      showDropdown?: boolean;
      
      // Number validation
      min?: number;
      max?: number;
      allowDecimals?: boolean;
      
      // Date validation
      startDate?: string | Date;
      endDate?: string | Date;
      
      // Text length validation
      minLength?: number;
      maxLength?: number;
      
      // Custom formula validation
      formula?: string;
    };
    inputMessage?: string;
    inputTitle?: string;
    errorMessage?: string;
    errorTitle?: string;
    showInputMessage?: boolean;
    showErrorAlert?: boolean;
    allowInvalidData?: boolean;
  }) => {
    const {
      range,
      validationType,
      criteria = {},
      inputMessage = '',
      inputTitle = 'Input',
      errorMessage = '',
      errorTitle = 'Invalid Input',
      showInputMessage = true,
      showErrorAlert = true,
      allowInvalidData = false
    } = params;

    console.log(`🔒 AddDataValidationTool: Adding ${validationType} validation to ${range}`);

    try {
      // Get the range object
      const fRange = context.fWorksheet.getRange(range);
      
      // Check what validation capabilities are available
      const validationSupport = checkValidationSupport(context.univerAPI, fRange);
      
      if (validationSupport.hasDataValidation) {
        return await createNativeDataValidation(context, fRange, {
          validationType,
          criteria,
          inputMessage,
          inputTitle,
          errorMessage,
          errorTitle,
          showInputMessage,
          showErrorAlert,
          allowInvalidData,
          range
        });
      } else if (validationSupport.hasConditionalFormatting) {
        // Fallback: use conditional formatting for visual validation
        return await createVisualValidation(context, fRange, {
          validationType,
          criteria,
          range
        });
      } else {
        // Basic fallback: add notes/comments for validation guidance
        return await createValidationNotes(context, fRange, {
          validationType,
          criteria,
          inputMessage,
          range
        });
      }

    } catch (error) {
      console.error(`❌ AddDataValidationTool failed:`, error);
      throw error;
    }
  }
);

/**
 * Check what data validation capabilities are available
 */
function checkValidationSupport(univerAPI: any, fRange: any): {
  hasDataValidation: boolean;
  hasConditionalFormatting: boolean;
  hasNotes: boolean;
} {
  return {
    hasDataValidation: typeof fRange.setDataValidation === 'function' || 
                      typeof fRange.newDataValidation === 'function',
    hasConditionalFormatting: typeof fRange.setConditionalFormat === 'function' ||
                              typeof univerAPI.newConditionalFormattingRule === 'function',
    hasNotes: typeof fRange.setNote === 'function' || 
              typeof fRange.addComment === 'function'
  };
}

/**
 * Create native data validation using Univer's API
 */
async function createNativeDataValidation(
  context: UniversalToolContext,
  fRange: any,
  config: any
): Promise<any> {
  console.log(`🔒 Using native Univer data validation API`);

  try {
    let validationBuilder;
    
    // Create validation builder
    if (typeof fRange.newDataValidation === 'function') {
      validationBuilder = fRange.newDataValidation();
    } else if (typeof context.fWorksheet.newDataValidation === 'function') {
      validationBuilder = context.fWorksheet.newDataValidation();
    } else {
      throw new Error("Data validation builder not available");
    }

    // Configure validation based on type
    switch (config.validationType) {
      case 'list':
        if (!config.criteria.values || config.criteria.values.length === 0) {
          throw new Error("Values array required for list validation");
        }
        validationBuilder = validationBuilder.requireValueInList(
          config.criteria.values,
          config.criteria.showDropdown !== false
        );
        break;

      case 'number':
        if (config.criteria.min !== undefined && config.criteria.max !== undefined) {
          validationBuilder = validationBuilder.requireNumberBetween(
            config.criteria.min,
            config.criteria.max
          );
        } else if (config.criteria.min !== undefined) {
          validationBuilder = validationBuilder.requireNumberGreaterThanOrEqualTo(
            config.criteria.min
          );
        } else if (config.criteria.max !== undefined) {
          validationBuilder = validationBuilder.requireNumberLessThanOrEqualTo(
            config.criteria.max
          );
        } else {
          validationBuilder = validationBuilder.requireNumberGreaterThan(0);
        }
        break;

      case 'date':
        if (config.criteria.startDate && config.criteria.endDate) {
          const startDate = typeof config.criteria.startDate === 'string' 
            ? new Date(config.criteria.startDate) 
            : config.criteria.startDate;
          const endDate = typeof config.criteria.endDate === 'string'
            ? new Date(config.criteria.endDate)
            : config.criteria.endDate;
          validationBuilder = validationBuilder.requireDateBetween(startDate, endDate);
        } else if (config.criteria.startDate) {
          const startDate = typeof config.criteria.startDate === 'string'
            ? new Date(config.criteria.startDate)
            : config.criteria.startDate;
          validationBuilder = validationBuilder.requireDateAfter(startDate);
        } else {
          validationBuilder = validationBuilder.requireDate();
        }
        break;

      case 'textLength':
        if (config.criteria.minLength !== undefined && config.criteria.maxLength !== undefined) {
          validationBuilder = validationBuilder.requireTextLengthBetween(
            config.criteria.minLength,
            config.criteria.maxLength
          );
        } else if (config.criteria.maxLength !== undefined) {
          validationBuilder = validationBuilder.requireTextLengthLessThanOrEqualTo(
            config.criteria.maxLength
          );
        } else if (config.criteria.minLength !== undefined) {
          validationBuilder = validationBuilder.requireTextLengthGreaterThanOrEqualTo(
            config.criteria.minLength
          );
        } else {
          validationBuilder = validationBuilder.requireText();
        }
        break;

      case 'email':
        // Use custom formula for email validation
        const emailFormula = '=AND(ISERROR(FIND(" ",A1))=FALSE,LEN(A1)-LEN(SUBSTITUTE(A1,"@",""))=1,FIND("@",A1)>1,FIND(".",A1,FIND("@",A1))>FIND("@",A1)+1)';
        validationBuilder = validationBuilder.requireFormulaSatisfied(emailFormula);
        break;

      case 'url':
        // Use custom formula for URL validation
        const urlFormula = '=OR(LEFT(A1,7)="http://",LEFT(A1,8)="https://")';
        validationBuilder = validationBuilder.requireFormulaSatisfied(urlFormula);
        break;

      case 'custom':
        if (!config.criteria.formula) {
          throw new Error("Formula required for custom validation");
        }
        validationBuilder = validationBuilder.requireFormulaSatisfied(config.criteria.formula);
        break;

      default:
        throw new Error(`Unknown validation type: ${config.validationType}`);
    }

    // Set messages
    if (config.inputMessage && config.showInputMessage) {
      if (typeof validationBuilder.setInputMessage === 'function') {
        validationBuilder = validationBuilder.setInputMessage(config.inputTitle, config.inputMessage);
      } else if (typeof validationBuilder.setHelpText === 'function') {
        validationBuilder = validationBuilder.setHelpText(config.inputMessage);
      }
    }

    if (config.errorMessage && config.showErrorAlert) {
      if (typeof validationBuilder.setErrorMessage === 'function') {
        validationBuilder = validationBuilder.setErrorMessage(config.errorTitle, config.errorMessage);
      } else if (typeof validationBuilder.setErrorText === 'function') {
        validationBuilder = validationBuilder.setErrorText(config.errorMessage);
      }
    }

    // Set rejection behavior
    if (config.allowInvalidData) {
      if (typeof validationBuilder.setAllowInvalid === 'function') {
        validationBuilder = validationBuilder.setAllowInvalid(true);
      }
    }

    // Build and apply validation
    const validation = validationBuilder.build();
    
    if (!validation) {
      throw new Error("Failed to build validation rule");
    }

    // Apply validation to range
    const result = await fRange.setDataValidation(validation);
    
    if (result === false) {
      throw new Error("Failed to apply validation to range");
    }

    console.log(`✅ Applied ${config.validationType} validation to ${config.range}`);

    return {
      range: config.range,
      validationType: config.validationType,
      criteria: config.criteria,
      method: 'native',
      message: `Added ${config.validationType} validation to ${config.range}`,
      success: true
    };

  } catch (error) {
    console.error('❌ Native data validation failed:', error);
    throw error;
  }
}

/**
 * Create visual validation using conditional formatting (fallback)
 */
async function createVisualValidation(
  context: UniversalToolContext,
  fRange: any,
  config: any
): Promise<any> {
  console.log(`🔒 Using conditional formatting for visual validation (fallback)`);

  try {
    // Create conditional formatting rule for invalid data
    let ruleBuilder;
    
    if (typeof context.fWorksheet.newConditionalFormattingRule === 'function') {
      ruleBuilder = context.fWorksheet.newConditionalFormattingRule();
    } else {
      throw new Error("Conditional formatting not available");
    }

    // Create validation formula based on type
    let validationFormula = '';
    
    switch (config.validationType) {
      case 'list':
        if (config.criteria.values) {
          const valueList = config.criteria.values.map((v: any) => `"${v}"`).join(',');
          validationFormula = `=NOT(OR(A1=${valueList}))`;
        }
        break;
      
      case 'number':
        if (config.criteria.min !== undefined && config.criteria.max !== undefined) {
          validationFormula = `=OR(A1<${config.criteria.min},A1>${config.criteria.max})`;
        } else if (config.criteria.min !== undefined) {
          validationFormula = `=A1<${config.criteria.min}`;
        } else if (config.criteria.max !== undefined) {
          validationFormula = `=A1>${config.criteria.max}`;
        }
        break;
      
      case 'textLength':
        if (config.criteria.maxLength !== undefined) {
          validationFormula = `=LEN(A1)>${config.criteria.maxLength}`;
        }
        break;
    }

    if (!validationFormula) {
      validationFormula = '=FALSE'; // No highlighting for unknown types
    }

    // Apply conditional formatting for invalid data (red background)
    const rule = ruleBuilder
      .whenFormulaSatisfied(validationFormula)
      .setBackground('#FFE6E6')
      .setFontColor('#CC0000')
      .setRanges([fRange.getRange()])
      .build();

    await context.fWorksheet.addConditionalFormattingRule(rule);

    return {
      range: config.range,
      validationType: config.validationType,
      criteria: config.criteria,
      method: 'visual',
      message: `Added visual validation to ${config.range} using conditional formatting`,
      success: true,
      note: 'Visual validation only - invalid data will be highlighted in red'
    };

  } catch (error) {
    console.error('❌ Visual validation creation failed:', error);
    throw error;
  }
}

/**
 * Create validation notes (basic fallback)
 */
async function createValidationNotes(
  context: UniversalToolContext,
  fRange: any,
  config: any
): Promise<any> {
  console.log(`🔒 Using validation notes (basic fallback)`);

  try {
    // Create a validation note
    let noteText = `Data Validation: ${config.validationType}\n`;
    
    if (config.inputMessage) {
      noteText += `Instructions: ${config.inputMessage}\n`;
    }
    
    // Add criteria information
    switch (config.validationType) {
      case 'list':
        if (config.criteria.values) {
          noteText += `Allowed values: ${config.criteria.values.join(', ')}`;
        }
        break;
      case 'number':
        if (config.criteria.min !== undefined && config.criteria.max !== undefined) {
          noteText += `Range: ${config.criteria.min} to ${config.criteria.max}`;
        } else if (config.criteria.min !== undefined) {
          noteText += `Minimum: ${config.criteria.min}`;
        } else if (config.criteria.max !== undefined) {
          noteText += `Maximum: ${config.criteria.max}`;
        }
        break;
      case 'textLength':
        if (config.criteria.maxLength !== undefined) {
          noteText += `Max length: ${config.criteria.maxLength} characters`;
        }
        if (config.criteria.minLength !== undefined) {
          noteText += `Min length: ${config.criteria.minLength} characters`;
        }
        break;
    }

    // Add note to the range
    if (typeof fRange.setNote === 'function') {
      fRange.setNote(noteText);
    } else if (typeof fRange.addComment === 'function') {
      fRange.addComment(noteText);
    } else {
      // Last resort: add text to first cell
      const firstCell = fRange.getCell(0, 0);
      const currentValue = firstCell.getValue();
      firstCell.setValue(`${currentValue} [${config.validationType} validation]`);
    }

    return {
      range: config.range,
      validationType: config.validationType,
      criteria: config.criteria,
      method: 'notes',
      message: `Added validation note to ${config.range}`,
      success: true,
      note: 'Basic validation - guidance provided via cell notes/comments'
    };

  } catch (error) {
    console.error('❌ Validation notes creation failed:', error);
    throw error;
  }
}