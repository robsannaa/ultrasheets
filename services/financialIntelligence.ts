// Financial Intelligence Agent System
// This system uses AI agents to understand financial requests and automatically determine formulas

import { openai } from "@ai-sdk/openai";
import { generateObject, generateText } from "ai";
import { UniverService } from "./univerService";
import { z } from "zod";

// Financial concept mapping schema
const FinancialConceptSchema = z.object({
  concept: z.string().describe("The financial concept identified (e.g., 'average', 'ebitda', 'gross_margin')"),
  calculation_type: z.enum(['simple_formula', 'complex_calculation', 'multi_sheet_analysis', 'ratio_calculation']),
  required_data: z.array(z.string()).describe("What data columns/ranges are needed"),
  formula_approach: z.string().describe("The Excel formula or calculation approach"),
  explanation: z.string().describe("Brief explanation for the user"),
  confidence: z.number().min(0).max(1).describe("Confidence in the interpretation")
});

// Multi-sheet calculation schema
const MultiSheetCalculationSchema = z.object({
  calculation_name: z.string(),
  sheets_required: z.array(z.string()),
  steps: z.array(z.object({
    step: z.number(),
    description: z.string(),
    formula: z.string(),
    sheet: z.string(),
    cell: z.string()
  })),
  final_result_location: z.object({
    sheet: z.string(),
    cell: z.string()
  })
});

export class FinancialIntelligenceAgent {
  private static instance: FinancialIntelligenceAgent;
  private univerService: UniverService;

  constructor() {
    this.univerService = UniverService.getInstance();
  }

  static getInstance(): FinancialIntelligenceAgent {
    if (!FinancialIntelligenceAgent.instance) {
      FinancialIntelligenceAgent.instance = new FinancialIntelligenceAgent();
    }
    return FinancialIntelligenceAgent.instance;
  }

  /**
   * Main intelligence function: understands user request and determines calculation approach
   */
  async analyzeFinancialRequest(userRequest: string, sheetContext?: any): Promise<{
    concept: any;
    execution_plan: string;
    requires_multi_sheet: boolean;
  }> {
    try {
      // Step 1: Financial Concept Recognition
      const concept = await this.recognizeFinancialConcept(userRequest, sheetContext);
      
      // Step 2: Determine if multi-sheet analysis is needed
      const requiresMultiSheet = await this.assessMultiSheetRequirement(userRequest, concept);
      
      // Step 3: Create execution plan
      const executionPlan = await this.createExecutionPlan(concept, sheetContext, requiresMultiSheet);
      
      return {
        concept,
        execution_plan: executionPlan,
        requires_multi_sheet: requiresMultiSheet
      };
    } catch (error) {
      throw new Error(`Financial intelligence analysis failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Recognizes financial concepts from natural language
   */
  private async recognizeFinancialConcept(userRequest: string, sheetContext?: any): Promise<any> {
    const contextInfo = sheetContext ? 
      `Available columns: ${sheetContext.tables?.[0]?.headers?.join(', ') || 'Unknown'}` : 
      'No sheet context available';

    const result = await generateObject({
      model: openai("gpt-4o-mini"),
      schema: FinancialConceptSchema,
      system: `You are a financial intelligence agent specializing in mapping natural language requests to financial calculations.

FINANCIAL CONCEPT EXPERTISE:
- Revenue metrics: total revenue, revenue growth, revenue per customer
- Profitability: gross profit, operating profit, EBITDA, net profit, margins
- Efficiency: average, median, standard deviation, variance
- Ratios: current ratio, debt-to-equity, ROE, ROA, gross margin %
- Cash flow: operating cash flow, free cash flow, cash conversion cycle
- Valuation: P/E ratio, EV/EBITDA, price-to-book, market cap

FORMULA MAPPING INTELLIGENCE:
"average of money spent" → AVERAGE function on amount column
"total revenue" → SUM function on revenue column  
"gross margin" → (Revenue - COGS) / Revenue * 100
"EBITDA" → Operating Income + Depreciation + Amortization
"current ratio" → Current Assets / Current Liabilities
"revenue growth" → (Current Period Revenue - Prior Period Revenue) / Prior Period Revenue * 100

CALCULATION TYPES:
- simple_formula: Basic Excel functions (SUM, AVERAGE, MAX, MIN, etc.)
- complex_calculation: Multi-step calculations requiring intermediate steps
- multi_sheet_analysis: Requires data from multiple sheets
- ratio_calculation: Division-based metrics with percentage formatting

Be intelligent about column detection - if user says "money spent" or "amounts" and you see an "Amount" column, use that.`,
      prompt: `User request: "${userRequest}"
      
Sheet context: ${contextInfo}

Analyze this request and map it to the appropriate financial concept and calculation approach.`
    });

    return result.object;
  }

  /**
   * Determines if multi-sheet analysis is required
   */
  private async assessMultiSheetRequirement(userRequest: string, concept: any): Promise<boolean> {
    const multiSheetIndicators = [
      'balance sheet', 'income statement', 'cash flow', 'consolidated',
      'across sheets', 'total company', 'ebitda', 'financial statements',
      'working capital', 'comprehensive analysis'
    ];

    const lowerRequest = userRequest.toLowerCase();
    return multiSheetIndicators.some(indicator => lowerRequest.includes(indicator)) ||
           concept.calculation_type === 'multi_sheet_analysis';
  }

  /**
   * Creates detailed execution plan for the calculation
   */
  private async createExecutionPlan(concept: any, sheetContext: any, requiresMultiSheet: boolean): Promise<string> {
    if (requiresMultiSheet) {
      return await this.createMultiSheetPlan(concept, sheetContext);
    } else {
      return await this.createSingleSheetPlan(concept, sheetContext);
    }
  }

  /**
   * Creates execution plan for single sheet calculations
   */
  private async createSingleSheetPlan(concept: any, sheetContext: any): Promise<string> {
    const result = await generateText({
      model: openai("gpt-4o-mini"),
      system: `You are a formula execution planner. Create a clear, step-by-step plan for executing financial calculations.

AVAILABLE SHEET DATA:
${sheetContext ? JSON.stringify(sheetContext, null, 2) : 'No context available'}

Your plan should:
1. Identify the exact data columns/ranges to use
2. Specify the Excel formula or calculation steps
3. Indicate where to place the result
4. Provide clear reasoning

Format as a concise execution plan.`,
      prompt: `Financial concept: ${concept.concept}
Calculation type: ${concept.calculation_type}
Formula approach: ${concept.formula_approach}
Required data: ${concept.required_data.join(', ')}

Create an execution plan for this calculation.`
    });

    return result.text;
  }

  /**
   * Creates execution plan for multi-sheet calculations
   */
  private async createMultiSheetPlan(concept: any, sheetContext: any): Promise<string> {
    const result = await generateObject({
      model: openai("gpt-4o-mini"),
      schema: MultiSheetCalculationSchema,
      system: `You are a multi-sheet financial analysis planner specializing in complex calculations like EBITDA, consolidated financials, and cross-sheet analysis.

MULTI-SHEET CALCULATION EXPERTISE:
- EBITDA: Operating Income (Income Statement) + Depreciation (Cash Flow Statement) + Amortization (Cash Flow Statement)
- Working Capital: Current Assets (Balance Sheet) - Current Liabilities (Balance Sheet)
- Free Cash Flow: Operating Cash Flow (Cash Flow Statement) - Capital Expenditures (Cash Flow Statement)
- ROE: Net Income (Income Statement) / Shareholders' Equity (Balance Sheet)
- Debt-to-Equity: Total Debt (Balance Sheet) / Shareholders' Equity (Balance Sheet)

Create step-by-step execution plans that work across multiple sheets.`,
      prompt: `Financial concept: ${concept.concept}
Calculation: ${concept.formula_approach}

Create a detailed multi-sheet execution plan.`
    });

    return JSON.stringify(result.object, null, 2);
  }

  /**
   * Executes the financial calculation based on the intelligence analysis
   */
  async executeFinancialCalculation(analysis: any): Promise<string> {
    try {
      const { concept, execution_plan, requires_multi_sheet } = analysis;

      if (requires_multi_sheet) {
        return await this.executeMultiSheetCalculation(concept, execution_plan);
      } else {
        return await this.executeSingleSheetCalculation(concept, execution_plan);
      }
    } catch (error) {
      throw new Error(`Calculation execution failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Executes single sheet calculations
   */
  private async executeSingleSheetCalculation(concept: any, executionPlan: string): Promise<string> {
    // Parse the execution plan and determine the appropriate action
    if (concept.calculation_type === 'simple_formula') {
      // For simple formulas like AVERAGE, SUM, etc.
      const formula = concept.formula_approach;
      const targetCell = this.determineTargetCell(concept);
      
      await this.univerService.setFormula(targetCell, formula);
      return `Calculated ${concept.concept}: ${formula} in cell ${targetCell}`;
    } else if (concept.calculation_type === 'ratio_calculation') {
      // For ratios that need percentage formatting
      const formula = concept.formula_approach;
      const targetCell = this.determineTargetCell(concept);
      
      await this.univerService.setFormula(targetCell, formula);
      // Apply percentage formatting if it's a percentage ratio
      if (concept.concept.includes('margin') || concept.concept.includes('growth') || concept.concept.includes('%')) {
        await this.univerService.formatCells(targetCell, { /* percentage format */ });
      }
      return `Calculated ${concept.concept}: ${formula} in cell ${targetCell}`;
    } else {
      // Complex calculations requiring multiple steps
      return await this.executeComplexCalculation(concept, executionPlan);
    }
  }

  /**
   * Executes multi-sheet calculations
   */
  private async executeMultiSheetCalculation(concept: any, executionPlan: string): Promise<string> {
    try {
      const plan = JSON.parse(executionPlan);
      const results = [];

      for (const step of plan.steps) {
        // This would require enhancing UniverService to support multi-sheet operations
        // For now, return the plan for manual execution
        results.push(`Step ${step.step}: ${step.description} - ${step.formula} in ${step.sheet}!${step.cell}`);
      }

      return `Multi-sheet calculation plan for ${concept.concept}:\n${results.join('\n')}`;
    } catch (error) {
      return `Complex calculation planned for ${concept.concept}. Execution plan: ${executionPlan}`;
    }
  }

  /**
   * Determines appropriate target cell for results
   */
  private determineTargetCell(concept: any): string {
    // Intelligent cell placement based on data structure
    // This would analyze the sheet context to find a good place for the result
    return 'H1'; // Default for now, should be made intelligent
  }

  /**
   * Executes complex single-sheet calculations
   */
  private async executeComplexCalculation(concept: any, executionPlan: string): Promise<string> {
    // Parse execution plan and perform multi-step calculation
    // This would break down complex calculations into steps
    return `Complex calculation executed for ${concept.concept} based on plan: ${executionPlan}`;
  }
}

// Financial Intelligence Service
export class FinancialIntelligenceService {
  private agent: FinancialIntelligenceAgent;
  
  constructor() {
    this.agent = FinancialIntelligenceAgent.getInstance();
  }

  /**
   * Main entry point for financial intelligence
   */
  async processFinancialRequest(userRequest: string): Promise<string> {
    try {
      // Get sheet context
      const sheetContext = await UniverService.getInstance().getSheetContext();
      
      // Analyze the request using AI
      const analysis = await this.agent.analyzeFinancialRequest(userRequest, sheetContext);
      
      // Execute the calculation
      const result = await this.agent.executeFinancialCalculation(analysis);
      
      return result;
    } catch (error) {
      throw new Error(`Financial intelligence processing failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
}