import { PerfProfiler } from "@better-teams/app/platform/perf";
import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { Skeleton } from "@better-teams/ui/components/skeleton";
import { Composer } from "../composer/Composer";
import { MembersSidebar, ProfileSidebar } from "../profile/ProfileCard";
import { Sidebar } from "../sidebar/Sidebar";
import { ThreadHeader } from "../thread/ThreadHeader";
import { ThreadView } from "../thread/ThreadView";
import { useProductivityWorkspaceController } from "./workspace-controller";

function MainLoadingSkeleton() {
  return (
    <div className="flex flex-1 flex-col px-8 py-8">
      <div className="mb-8 flex items-center gap-3">
        <Skeleton className="size-11 rounded-xl" />
        <div className="space-y-2">
          <Skeleton className="h-4 w-44" />
          <Skeleton className="h-3 w-28" />
        </div>
      </div>
      <div className="space-y-6">
        {[0.74, 0.48, 0.62].map((width) => (
          <div key={width} className="flex gap-3">
            <Skeleton className="size-9 shrink-0 rounded-xl" />
            <div className="flex-1 space-y-2">
              <Skeleton className="h-3.5 w-24" />
              <Skeleton
                className="h-10 rounded-xl"
                style={{ width: `${width * 100}%` }}
              />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function EmptyWorkspaceState({
  title,
  subtitle,
}: {
  title: string;
  subtitle: string;
}) {
  return (
    <div className="flex flex-1 flex-col items-center justify-center gap-4">
      <div className="flex size-20 items-center justify-center rounded-3xl bg-accent">
        <span className="text-4xl text-muted-foreground/20">💬</span>
      </div>
      <div className="text-center">
        <p className="text-[15px] font-semibold text-muted-foreground/40">
          {title}
        </p>
        <p className="mt-1 text-[12px] text-muted-foreground/25">{subtitle}</p>
      </div>
    </div>
  );
}

function ProductivityWorkspaceContent() {
  const workspace = useProductivityWorkspaceController();

  if (workspace.errorMessage) {
    return (
      <div className="flex h-full flex-1 flex-col items-center justify-center gap-4 bg-background">
        <div className="flex size-12 items-center justify-center rounded-2xl bg-destructive/10">
          <span className="text-lg text-destructive">!</span>
        </div>
        <p className="max-w-sm text-center text-[13px] text-muted-foreground">
          {workspace.errorMessage}
        </p>
        <button
          type="button"
          onClick={() => void workspace.sessionQuery.refetch()}
          className="rounded-xl bg-primary px-4 py-2 text-[13px] font-medium text-primary-foreground transition-colors hover:bg-primary/90"
        >
          Try again
        </button>
      </div>
    );
  }

  return (
    <>
      <div aria-live="polite" className="sr-only">
        {workspace.announcement}
      </div>

      <div ref={workspace.workspaceRef} className="flex h-full min-h-0 flex-1">
        <PerfProfiler
          id="Sidebar"
          detail={{
            conversationCount: workspace.allSidebarItems.length,
            selectedConversation:
              workspace.pendingSelectedId ??
              workspace.activeConversationId ??
              "__none__",
          }}
        >
          <Sidebar
            upn={workspace.session?.upn}
            selfDisplayName={workspace.selfDisplayName}
            selfAvatarSrc={workspace.selfAvatarSrc}
            accountAvatarByTenant={workspace.accountAvatarByTenant}
            presenceByMri={workspace.presenceByMri}
            accounts={workspace.accounts}
            activeTenantId={workspace.activeTenantId}
            onSwitchAccount={workspace.switchAccount}
            switchPending={
              workspace.isSwitchingAccount || workspace.sessionQuery.isFetching
            }
            allSidebarItems={workspace.allSidebarItems}
            activeConversationId={
              workspace.pendingSelectedId ?? workspace.activeConversationId
            }
            onSelectConversation={workspace.handleSelectConversation}
            onHoverConversationStart={workspace.handleHoverConversation}
            onHoverConversationEnd={workspace.handleHoverConversationEnd}
            onToggleFavorite={workspace.handleToggleFavorite}
            searchInputRef={workspace.searchInputRef}
            accountLoading={workspace.accountLoading}
            conversationsLoading={workspace.conversationsLoading}
            avatarFallbackReady={workspace.avatarFallbackReady}
          />
        </PerfProfiler>

        <main className="flex min-h-0 min-w-0 flex-1 flex-col bg-background">
          {workspace.conversationsLoading ? (
            <MainLoadingSkeleton />
          ) : !workspace.selectedItem &&
            workspace.allSidebarItems.length === 0 ? (
            <EmptyWorkspaceState
              title="No conversations yet"
              subtitle="Start or receive a Teams chat to see it here"
            />
          ) : !workspace.selectedItem ? (
            <EmptyWorkspaceState
              title="Select a conversation"
              subtitle="Choose from your chats on the left to get started"
            />
          ) : (
            <>
              <ThreadHeader
                title={workspace.selectedItem.title}
                kind={workspace.selectedItem.kind}
                memberCount={workspace.selectedHeaderMemberCount}
                avatarMris={workspace.selectedHeaderAvatarMris}
                avatarByMri={workspace.selectedAvatarThumbByMri}
                avatarLabelByMri={workspace.selectedHeaderAvatarLabelsByMri}
                avatarFallbackReady={workspace.selectedAvatarFallbackReady}
                presenceByMri={workspace.selectedPresenceByMri}
                onOpenProfile={workspace.handleOpenSelectedProfile}
                profileButtonLabel={workspace.selectedProfileButtonLabel}
                onOpenMembers={workspace.handleOpenMembersSidebar}
                searchQuery={workspace.threadSearchQuery}
                searchResultCount={workspace.threadSearchResultCount}
                onSearchQueryChange={workspace.setThreadSearchQuery}
                onSubmitSearch={workspace.handleSubmitSearch}
                onCloseSearch={workspace.handleCloseSearch}
              />
              <PerfProfiler
                id="ThreadView"
                detail={{
                  conversationId: workspace.activeConversationId as string,
                  kind: workspace.selectedItem.kind,
                }}
              >
                <ThreadView
                  ref={workspace.threadViewRef}
                  key={`thread-${workspace.activeTenantId ?? "__default__"}-${workspace.activeConversationId}`}
                  tenantId={workspace.activeTenantId}
                  conversationId={workspace.activeConversationId as string}
                  conversationKind={workspace.selectedItem.kind}
                  liveSessionReady={workspace.liveSessionReady}
                  autoFocus={workspace.selectionFocusTarget === "thread"}
                  searchQuery={workspace.threadSearchQuery}
                  consumptionHorizon={
                    workspace.selectedItem.conversation.consumptionHorizon
                  }
                  onSearchResultCountChange={
                    workspace.setThreadSearchResultCount
                  }
                  selfSkypeId={workspace.session?.skypeId}
                  selfDisplayName={workspace.selfDisplayName}
                  avatarByMri={workspace.avatarThumbByMri}
                  avatarFullByMri={workspace.avatarFullByMri}
                  avatarFallbackReady={workspace.avatarFallbackReady}
                  displayNameByMri={workspace.displayNameByMri}
                  emailByMri={workspace.emailByMri}
                  jobTitleByMri={workspace.jobTitleByMri}
                  departmentByMri={workspace.departmentByMri}
                  companyNameByMri={workspace.companyNameByMri}
                  tenantNameByMri={workspace.tenantNameByMri}
                  locationByMri={workspace.locationByMri}
                  sharedConversationsByMri={workspace.sharedConversationsByMri}
                  onOpenProfile={workspace.handleOpenThreadProfile}
                />
              </PerfProfiler>
              <PerfProfiler
                id="Composer"
                detail={{
                  conversationId: workspace.activeConversationId as string,
                  mentionCandidateCount:
                    workspace.composerMentionCandidates.length,
                }}
              >
                <Composer
                  key={`composer-${workspace.activeTenantId ?? "__default__"}-${workspace.activeConversationId}`}
                  tenantId={workspace.activeTenantId}
                  conversationId={workspace.activeConversationId as string}
                  conversationTitle={workspace.selectedItem.title}
                  conversationMembers={
                    workspace.selectedConversationMembers
                      ?.map((member) => member.id ?? "")
                      .filter(Boolean) ?? []
                  }
                  composerRef={workspace.composerRef}
                  liveSessionReady={workspace.liveSessionReady}
                  mentionCandidates={workspace.composerMentionCandidates}
                />
              </PerfProfiler>
            </>
          )}
        </main>

        {workspace.profileSidebarData ? (
          <ProfileSidebar
            profile={workspace.profileSidebarData}
            closeLabel="Close profile sidebar"
            onClose={workspace.handleCloseProfileSidebar}
          />
        ) : workspace.membersSidebarOpen ? (
          <MembersSidebar
            members={workspace.selectedMemberProfiles}
            memberCount={workspace.selectedHeaderMemberCount}
            onOpenProfile={workspace.handleOpenMemberProfile}
            closeLabel="Close members sidebar"
            onClose={workspace.handleCloseMembersSidebar}
          />
        ) : null}
      </div>
    </>
  );
}

export function ProductivityWorkspace() {
  const { activeTenantId } = useTeamsAccountContext();

  return <ProductivityWorkspaceContent key={activeTenantId ?? "__default__"} />;
}
