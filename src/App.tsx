import { ChatWorkspace } from "@/components/chat/ChatWorkspace";
import { QueryProvider } from "@/providers/QueryProvider";
import { TeamsAccountProvider } from "@/providers/TeamsAccountProvider";
import { ThemeProvider } from "@/providers/ThemeProvider";

export function App() {
  return (
    <ThemeProvider>
      <QueryProvider>
        <TeamsAccountProvider>
          <div className="flex h-screen flex-col overflow-hidden">
            <ChatWorkspace />
          </div>
        </TeamsAccountProvider>
      </QueryProvider>
    </ThemeProvider>
  );
}
