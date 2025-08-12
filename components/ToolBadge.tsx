import { cn } from "@/lib/utils";
import { CheckCircle2, Loader2 } from "lucide-react";

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
  return (
    <div className="my-0.5">
      <div
        className={cn(
          "inline-flex items-center gap-1.5 text-xs rounded-md px-2 py-1 max-w-full break-words whitespace-normal",
          isRunning ? "bg-blue-50 text-blue-700" : "bg-green-50 text-green-700"
        )}
      >
        {isRunning ? (
          <Loader2 className="h-3 w-3 animate-spin shrink-0" />
        ) : (
          <CheckCircle2 className="h-3 w-3 shrink-0" />
        )}
        <span className="truncate text-xs font-medium">{label}</span>
      </div>
    </div>
  );
}
