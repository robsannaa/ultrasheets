"use client";

import { Univer } from "@/components/univer";
import { ChatSidebar } from "@/components/chat-sidebar";
import { SidebarProvider } from "@/components/ui/sidebar";
import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from "@/components/ui/resizable";
import { useIsMobile } from "@/hooks/use-mobile";
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { MessageSquare, X } from "lucide-react";
import Image from "next/image";

export default function Home() {
  const isMobile = useIsMobile();
  const [showSidebar, setShowSidebar] = useState(false);

  if (isMobile) {
    return (
      <SidebarProvider>
        <div className="h-svh w-svw min-h-0 bg-background relative">
          {/* Main Univer Component */}
          <div className="h-full w-full p-2">
            <Univer />
          </div>

          {/* Mobile Chat Toggle Button */}
          <Button
            onClick={() => setShowSidebar(true)}
            size="icon"
            className="fixed bottom-4 right-4 z-50 h-12 w-12 rounded-full shadow-lg"
            aria-label="Open chat"
          >
            <MessageSquare className="h-5 w-5" />
          </Button>

          {/* Mobile Chat Overlay */}
          {showSidebar && (
            <>
              {/* Backdrop */}
              <div
                className="fixed inset-0 bg-black/50 z-40"
                onClick={() => setShowSidebar(false)}
              />

              {/* Sliding Chat Panel */}
              <div className="fixed inset-y-0 right-0 w-full max-w-sm z-50">
                <div className="h-full flex flex-col">
                  {/* Header with close button */}
                  <div className="flex items-center justify-between p-4 rounded-t-2xl bg-background">
                    <h2 className="text-lg font-semibold">CellChat</h2>
                    <Image
                      src="/logo.png"
                      alt="CellChat"
                      width={24}
                      height={24}
                    />
                    <Button
                      onClick={() => setShowSidebar(false)}
                      size="icon"
                      variant="ghost"
                      className="h-8 w-8"
                      aria-label="Close chat"
                    >
                      <X className="h-4 w-4" />
                    </Button>
                  </div>

                  {/* Chat Content */}
                  <div className="flex-1 overflow-hidden rounded-b-2xl m-2">
                    <ChatSidebar onMobileClose={() => setShowSidebar(false)} />
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </SidebarProvider>
    );
  }

  // Desktop Layout
  return (
    <SidebarProvider>
      <div className="bg-background h-screen w-screen border-0 m-1">
        <ResizablePanelGroup direction="horizontal">
          <ResizablePanel
            defaultSize={82}
            minSize={60}
            className="my-1 mx-1 border rounded-2xl"
          >
            <Univer />
          </ResizablePanel>
          <ResizableHandle withHandle />
          <ResizablePanel defaultSize={30} minSize={15} maxSize={40}>
            <ChatSidebar />
          </ResizablePanel>
        </ResizablePanelGroup>
      </div>
    </SidebarProvider>
  );
}
