import { cn } from "@/lib/utils";
import { CheckCircle2, Loader2, XCircle } from "lucide-react";

export function ToolBadge({
  label,
  toolName,
  state,
}: {
  label: string;
  toolName?: string;
  state: "call" | "result" | "error";
}) {
  const isRunning = state === "call";
  const isError = state === "error";
  return (
    <div className="my-0.5">
      <div
        className={cn(
          "inline-flex items-center gap-1.5 text-xs rounded-md px-2 py-1 max-w-full break-words whitespace-normal",
          isError
            ? "bg-red-50 text-red-700"
            : isRunning
            ? "bg-blue-50 text-blue-700"
            : "bg-green-50 text-green-700"
        )}
      >
        {isRunning ? (
          <Loader2 className="h-3 w-3 animate-spin shrink-0" />
        ) : isError ? (
          <XCircle className="h-3 w-3 shrink-0" />
        ) : (
          <CheckCircle2 className="h-3 w-3 shrink-0" />
        )}
        <span className="truncate text-xs font-medium">
          {toolName ? (
            <span className="font-mono text-[10px]">{toolName}</span>
          ) : (
            label
          )}
        </span>
      </div>
    </div>
  );
}
