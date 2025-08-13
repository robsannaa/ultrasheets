"use client";

import * as React from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Send, Loader2, Bot, MessageSquare, Settings, Database, CheckCircle, AlertCircle } from "lucide-react";
import { getCacheStats, debugAIPerformance } from "@/lib/unified-context-system";

export function ChatInput({
  input,
  handleInputChange,
  handleSubmit,
  isLoading,
  onStop,
  contextMetadata,
}: {
  input: string;
  handleInputChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  handleSubmit: (e: React.FormEvent<HTMLFormElement>) => void;
  isLoading: boolean;
  onStop?: () => void;
  contextMetadata?: {
    tablesCount: number;
    hasSelection: boolean;
    estimatedTokens: number;
    contextType: string;
  };
}) {
  const [isAgentMode, setIsAgentMode] = React.useState(false);
  const [isPerformingModel, setIsPerformingModel] = React.useState(false);
  const [settingsOpen, setSettingsOpen] = React.useState(false);
  const [cacheStats, setCacheStats] = React.useState<any>(null);
  const [showContextInfo, setShowContextInfo] = React.useState(false);

  // Monitor AI performance and cache stats
  React.useEffect(() => {
    const updateStats = async () => {
      try {
        const stats = await getCacheStats();
        setCacheStats(stats);
      } catch {
        // Ignore errors in stats collection
      }
    };
    
    updateStats();
    const interval = setInterval(updateStats, 5000); // Update every 5s
    return () => clearInterval(interval);
  }, []);
  
  // Debug performance on demand
  const handleDebugPerformance = async () => {
    try {
      await debugAIPerformance();
    } catch (error) {
      console.error('Failed to debug AI performance:', error);
    }
  };

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
    <form onSubmit={onSubmit} className="space-y-2 w-full">
      <div className="border-2 border-muted-foreground/20 rounded-xl p-3 bg-background">
        <div className="relative">
          <Input
            placeholder={
              isAgentMode
                ? "Ask the AI agent..."
                : "Ask about your spreadsheet data..."
            }
            value={input}
            onChange={handleInputChange}
            onKeyDown={onKeyDown}
            disabled={isLoading}
            className="w-full h-12 text-sm px-4 py-3 border-0 focus-visible:ring-0 focus-visible:ring-offset-0 rounded-xl bg-transparent"
          />
          <div className="absolute right-2 top-1/2 -translate-y-1/2 flex gap-2">
            {isLoading ? (
              <Button
                type="button"
                onClick={() => onStop?.()}
                variant="outline"
                size="sm"
                className="h-8 px-3 rounded-full bg-background"
              >
                <Loader2 className="h-3.5 w-3.5 animate-spin mr-1" />
                Stop
              </Button>
            ) : (
              <Button
                type="submit"
                size="icon"
                disabled={!input.trim()}
                aria-disabled={!input.trim()}
                className="h-8 w-8 shrink-0 hover:bg-muted"
              >
                <Send className="h-4 w-4" />
              </Button>
            )}
          </div>
        </div>

        <div className="flex items-center justify-between mt-3">
          <div className="flex items-center gap-2">
            <Button
              type="button"
              variant="ghost"
              size="sm"
              className="h-8 w-8 p-0 rounded-full hover:bg-muted"
              onClick={() => setShowContextInfo(v => !v)}
              title={`Context: ${contextMetadata?.tablesCount || 0} tables`}
            >
              <Database className={`h-4 w-4 ${
                contextMetadata?.tablesCount ? 'text-green-600' : 'text-muted-foreground'
              }`} />
            </Button>

            <div className="relative">
              <Button
                type="button"
                variant="ghost"
                size="sm"
                className="h-8 w-8 p-0 rounded-full hover:bg-muted"
                onClick={() => setSettingsOpen((v) => !v)}
                aria-expanded={settingsOpen}
                aria-haspopup="menu"
              >
                <Settings className="h-4 w-4" />
              </Button>
              {settingsOpen && (
                <div
                  className="absolute z-20 bottom-full mb-2 right-0 w-72 max-h-64 overflow-auto p-3 rounded-md border bg-popover text-popover-foreground shadow-sm"
                  role="menu"
                  aria-label="Model settings"
                >
                  <div>
                    <h4 className="font-medium text-sm mb-2">
                      Model Selection
                    </h4>
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="text-sm font-medium text-foreground">
                          {isPerformingModel ? "CC-25" : "CC-01"}
                        </span>
                        {isPerformingModel ? (
                          <span className="text-xs text-blue-600 dark:text-blue-400 font-medium">
                            Turbo
                          </span>
                        ) : (
                          <span className="text-xs text-muted-foreground">
                            Basic
                          </span>
                        )}
                      </div>
                      <Button
                        type="button"
                        variant="outline"
                        size="sm"
                        className="h-7 px-2"
                        onClick={() => setIsPerformingModel((v) => !v)}
                      >
                        {isPerformingModel ? "On" : "Off"}
                      </Button>
                    </div>
                    <p className="text-xs text-muted-foreground mt-2">
                      {isPerformingModel
                        ? "High-performance model with advanced capabilities"
                        : "Standard model for everyday tasks"}
                    </p>
                    
                    {/* Context info section */}
                    {contextMetadata && (
                      <div className="mt-4 pt-3 border-t">
                        <h4 className="font-medium text-sm mb-2">Context Status</h4>
                        <div className="space-y-1 text-xs">
                          <div className="flex justify-between">
                            <span>Tables:</span>
                            <span className="font-mono">{contextMetadata.tablesCount}</span>
                          </div>
                          <div className="flex justify-between">
                            <span>Selection:</span>
                            <span className={contextMetadata.hasSelection ? 'text-green-600' : 'text-muted-foreground'}>
                              {contextMetadata.hasSelection ? 'Active' : 'None'}
                            </span>
                          </div>
                          <div className="flex justify-between">
                            <span>Context Type:</span>
                            <span className="font-mono">{contextMetadata.contextType}</span>
                          </div>
                          <div className="flex justify-between">
                            <span>Tokens:</span>
                            <span className={`font-mono ${
                              contextMetadata.estimatedTokens > 2000 ? 'text-amber-600' : 'text-green-600'
                            }`}>
                              {contextMetadata.estimatedTokens}
                            </span>
                          </div>
                        </div>
                      </div>
                    )}
                    
                    {/* Cache performance info */}
                    {cacheStats && (
                      <div className="mt-3 pt-3 border-t">
                        <h4 className="font-medium text-sm mb-2">AI Performance</h4>
                        <div className="space-y-1 text-xs">
                          <div className="flex justify-between">
                            <span>Cache Strategy:</span>
                            <span className="font-mono">{cacheStats.strategy}</span>
                          </div>
                          <div className="flex justify-between">
                            <span>Hit Rate:</span>
                            <span className={`font-mono ${
                              cacheStats.hitRate > 80 ? 'text-green-600' : 
                              cacheStats.hitRate > 60 ? 'text-amber-600' : 'text-red-600'
                            }`}>
                              {cacheStats.hitRate}%
                            </span>
                          </div>
                          <div className="flex justify-between">
                            <span>Cache Age:</span>
                            <span className="font-mono">{Math.round(cacheStats.cacheAge / 1000)}s</span>
                          </div>
                        </div>
                        <Button
                          type="button"
                          variant="outline"
                          size="sm"
                          className="w-full mt-2"
                          onClick={handleDebugPerformance}
                        >
                          Debug Performance
                        </Button>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>

            <Button
              type="button"
              variant="ghost"
              size="sm"
              onClick={() => setIsAgentMode((m) => !m)}
              className={
                isAgentMode
                  ? "px-2 py-1 h-8 text-xs font-medium rounded-md bg-teal-100 text-teal-700 hover:bg-teal-200 dark:bg-teal-900/30 dark:text-teal-400 dark:hover:bg-teal-900/50"
                  : "px-2 py-1 h-8 text-xs font-medium rounded-md bg-muted text-muted-foreground hover:bg-muted/80"
              }
            >
              {isAgentMode ? "Agent" : "Ask"}
            </Button>
          </div>

          {/* Context and performance indicators */}
          <div className="flex items-center gap-2">
            {/* Context indicator */}
            {contextMetadata && (
              <div className="text-xs text-muted-foreground flex items-center gap-1">
                {contextMetadata.hasSelection && (
                  <CheckCircle className="h-3 w-3 text-green-500" />
                )}
                <span>{contextMetadata.tablesCount} tables</span>
                {contextMetadata.estimatedTokens > 1000 && (
                  <AlertCircle className="h-3 w-3 text-amber-500" />
                )}
              </div>
            )}
            
            {/* Cache performance indicator */}
            {cacheStats && (
              <div className="text-xs text-muted-foreground flex items-center gap-1">
                <div 
                  className={`h-2 w-2 rounded-full ${
                    cacheStats.hitRate > 80 ? 'bg-green-500' : 
                    cacheStats.hitRate > 60 ? 'bg-amber-500' : 'bg-red-500'
                  }`} 
                  title={`Cache hit rate: ${cacheStats.hitRate}%`}
                />
                <span className="cursor-pointer" onClick={handleDebugPerformance}>
                  {cacheStats.strategy}
                </span>
              </div>
            )}
          </div>
        </div>
      </div>
    </form>
  );
}
