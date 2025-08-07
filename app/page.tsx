import { Univer } from "@/components/univer";
import { ChatSidebar } from "@/components/chat-sidebar";
import { SidebarProvider } from "@/components/ui/sidebar";
import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from "@/components/ui/resizable";

export default function Home() {
  return (
    <SidebarProvider>
      <div className="max-h-screen w-screen bg-gray-200">
        <ResizablePanelGroup direction="horizontal">
          <ResizablePanel
            defaultSize={75}
            minSize={30}
            className="bg-background my-2 mx-1 rounded-2xl"
          >
            <Univer />
          </ResizablePanel>
          <ResizableHandle withHandle />
          <ResizablePanel defaultSize={25} minSize={15} maxSize={50}>
            <ChatSidebar />
          </ResizablePanel>
        </ResizablePanelGroup>
      </div>
    </SidebarProvider>
  );
}
