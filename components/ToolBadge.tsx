import React from "react";
import { cn } from "@/lib/utils";
import { Badge } from "@/components/ui/badge";

interface ToolBadgeProps {
  label: string;
  toolName: string;
  state: "call" | "result";
}

export function ToolBadge({ label, toolName, state }: ToolBadgeProps) {
  const isCall = state === "call";
  
  return (
    <Badge
      variant={isCall ? "secondary" : "default"}
      className={cn(
        "text-xs max-w-full truncate",
        isCall 
          ? "bg-blue-100 text-blue-800 border-blue-200" 
          : "bg-green-100 text-green-800 border-green-200"
      )}
    >
      {isCall ? "🔄" : "✅"} {label}
    </Badge>
  );
}