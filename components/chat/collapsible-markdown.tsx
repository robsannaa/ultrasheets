"use client";

import * as React from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { cn } from "@/lib/utils";

export function CollapsibleMarkdown({ text }: { text: string }) {
  const normalize = React.useCallback((t: string) => {
    if (!t) return "";
    let s = t.replace(/\r\n?/g, "\n").trim();
    s = s.replace(/\n{3,}/g, "\n\n");
    return s;
  }, []);
  const normalized = normalize(text);

  return (
    <div className="chat-markdown leading-5 text-sm max-w-none break-words whitespace-normal">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={{
          table: () => <></>,
          th: () => <></>,
          td: () => <></>,
          p: (props) => <p className="mb-0.5 leading-5" {...props} />,
          h1: (props) => <h1 className="mb-0.5 text-base" {...props} />,
          h2: (props) => <h2 className="mb-0.5 text-sm" {...props} />,
          h3: (props) => <h3 className="mb-0.5 text-sm" {...props} />,
          ul: (props) => (
            <ul className="list-disc ml-4 my-1 space-y-1" {...props} />
          ),
          ol: (props) => (
            <ol className="list-decimal ml-4 my-1 space-y-1" {...props} />
          ),
          li: (props) => <li className="mb-0.5" {...props} />,
          code: ({ inline, children, ...props }: any) => (
            <code
              className={cn(
                inline
                  ? "px-1 py-0.5 rounded bg-muted"
                  : "block p-3 rounded bg-muted overflow-x-auto"
              )}
              {...props}
            >
              {children}
            </code>
          ),
          a: (props) => (
            <a
              className="underline text-primary"
              target="_blank"
              rel="noreferrer"
              {...props}
            />
          ),
        }}
      >
        {normalized}
      </ReactMarkdown>
    </div>
  );
}
