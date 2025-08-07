import { useCallback } from "react";

export interface UniverToolResult {
  success: boolean;
  data?: any;
  error?: string;
  message?: string;
}

export const useUniverTools = () => {
  const executeTool = useCallback(
    async (toolName: string, params?: any): Promise<UniverToolResult> => {
      try {
        console.log(`🔧 useUniverTools: Executing ${toolName}`, params);

        // Check if the global function exists
        if (typeof window.executeUniverTool !== "function") {
          return {
            success: false,
            error: "Univer tools not available. Please refresh the page.",
          };
        }

        // Execute the tool
        const result = await window.executeUniverTool(toolName, params);

        console.log(
          `✅ useUniverTools: ${toolName} executed successfully`,
          result
        );

        return {
          success: true,
          data: result,
          message: result.message || `Successfully executed ${toolName}`,
        };
      } catch (error) {
        console.error(`❌ useUniverTools: ${toolName} execution failed`, error);

        return {
          success: false,
          error:
            error instanceof Error ? error.message : "Unknown error occurred",
        };
      }
    },
    []
  );

  const handleToolResponse = useCallback(
    async (response: any): Promise<UniverToolResult> => {
      try {
        // Check if this is a client-side tool response
        if (response.clientSide && response.toolName) {
          console.log(
            `🔄 useUniverTools: Handling client-side tool response`,
            response
          );

          return await executeTool(response.toolName, response.params);
        }

        // Not a client-side tool, return as is
        return {
          success: true,
          data: response,
        };
      } catch (error) {
        console.error("❌ useUniverTools: Error handling tool response", error);

        return {
          success: false,
          error:
            error instanceof Error ? error.message : "Unknown error occurred",
        };
      }
    },
    [executeTool]
  );

  return {
    executeTool,
    handleToolResponse,
  };
};
