import { ChatWorkspace } from "@better-teams/app/components/chat/ChatWorkspace";
import { QueryProvider } from "@better-teams/app/providers/QueryProvider";
import { TeamsAccountProvider } from "@better-teams/app/providers/TeamsAccountProvider";
import { ThemeProvider } from "@better-teams/app/providers/ThemeProvider";

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
