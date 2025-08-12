"use client";

import * as React from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Send, Loader2, Bot, MessageSquare } from "lucide-react";

export function ChatInput({
  input,
  handleInputChange,
  handleSubmit,
  isLoading,
}: {
  input: string;
  handleInputChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  handleSubmit: (e: React.FormEvent<HTMLFormElement>) => void;
  isLoading: boolean;
}) {
  const [isAgentMode, setIsAgentMode] = React.useState(false);

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      const form = e.currentTarget.closest("form");
      if (form) form.requestSubmit();
    }
  };

  const onSubmit = (e: React.FormEvent<HTMLFormElement>) => {
    if (isLoading || !input.trim()) {
      e.preventDefault();
      return;
    }
    handleSubmit(e);
  };

  return (
    <form onSubmit={onSubmit} className="relative flex items-center max-w-full">
      <div className="relative flex-1">
        <Input
          placeholder={
            isAgentMode
              ? "Ask the AI agent..."
              : "Ask about your spreadsheet data..."
          }
          value={input}
          onChange={handleInputChange}
          onKeyDown={onKeyDown}
          className="pr-12 py-6 text-sm rounded-2xl border-0 shadow-sm bg-muted/50 focus-visible:ring-1 focus-visible:ring-offset-0"
        />
        <div className="absolute right-2 top-1/2 -translate-y-1/2 flex gap-2">
          <Button
            type="button"
            size="icon"
            variant="ghost"
            onClick={() => setIsAgentMode(!isAgentMode)}
            className="h-8 w-8 shrink-0 hover:bg-muted"
          >
            {isAgentMode ? (
              <Bot className="h-4 w-4" />
            ) : (
              <MessageSquare className="h-4 w-4" />
            )}
          </Button>
          <Button
            type="submit"
            size="icon"
            disabled={isLoading || !input.trim()}
            aria-busy={isLoading}
            aria-disabled={isLoading || !input.trim()}
            className="h-8 w-8 shrink-0 hover:bg-muted"
          >
            {isLoading ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <Send className="h-4 w-4" />
            )}
          </Button>
        </div>
      </div>
    </form>
  );
}
