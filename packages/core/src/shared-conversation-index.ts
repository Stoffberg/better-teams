import type { Conversation, ConversationMember } from "./teams/types";
import { canonAvatarMri } from "./teams-profile-avatars";

export type SharedConversationKind = "dm" | "group" | "meeting";

export type SharedConversationSource = {
  id: string;
  conversation: Conversation;
  title: string;
  kind: SharedConversationKind;
  preview: string;
  sideTime: string;
  avatarMri?: string;
  isFavorite?: boolean;
  searchText?: string;
  avatarThumbSrc?: string;
  avatarFullSrc?: string;
};

type SharedConversationEntry = {
  id: string;
  title: string;
  kind: SharedConversationKind;
  preview: string;
  sideTime: string;
};

function normalizeProfileText(value: string | undefined): string | undefined {
  const trimmed = value?.trim().toLowerCase();
  if (!trimmed) return undefined;
  return trimmed.replace(/\s+/g, " ");
}

function memberDisplayName(member: ConversationMember): string | undefined {
  const displayName =
    member.displayName ||
    member.friendlyName ||
    (typeof member.userPrincipalName === "string"
      ? member.userPrincipalName.split("@")[0]
      : undefined);
  return normalizeProfileText(displayName);
}

function conversationMembersForItem(
  item: SharedConversationSource,
  detailedMembersByConversationId: Record<string, ConversationMember[]>,
): ConversationMember[] {
  return (
    detailedMembersByConversationId[item.id] ?? item.conversation.members ?? []
  );
}

export function buildSharedConversationsByMri(
  items: SharedConversationSource[],
  detailedMembersByConversationId: Record<string, ConversationMember[]>,
  displayNameByMri: Record<string, string>,
  emailByMri: Record<string, string>,
): Record<string, SharedConversationEntry[]> {
  const safeDisplayNameByMri = displayNameByMri ?? {};
  const safeEmailByMri = emailByMri ?? {};
  const byMri = new Map<string, SharedConversationEntry[]>();
  const byEmail = new Map<string, SharedConversationEntry[]>();
  const byName = new Map<string, SharedConversationEntry[]>();

  for (const item of items) {
    const seenConversationMris = new Set<string>();
    const seenEmails = new Set<string>();
    const seenNames = new Set<string>();
    if (item.avatarMri) {
      seenConversationMris.add(canonAvatarMri(item.avatarMri));
    }

    for (const member of conversationMembersForItem(
      item,
      detailedMembersByConversationId,
    )) {
      if (member.id) {
        seenConversationMris.add(canonAvatarMri(member.id));
      }
      const email =
        typeof member.userPrincipalName === "string"
          ? normalizeProfileText(member.userPrincipalName)
          : undefined;
      if (email) {
        seenEmails.add(email);
      }
      const displayName = memberDisplayName(member);
      if (displayName) {
        seenNames.add(displayName);
      }
    }

    const sharedConversation = {
      id: item.id,
      title: item.title,
      kind: item.kind,
      preview: item.preview,
      sideTime: item.sideTime,
    } satisfies SharedConversationEntry;

    for (const mri of seenConversationMris) {
      const entries = byMri.get(mri);
      if (entries) {
        entries.push(sharedConversation);
      } else {
        byMri.set(mri, [sharedConversation]);
      }
    }

    for (const email of seenEmails) {
      const entries = byEmail.get(email);
      if (entries) {
        entries.push(sharedConversation);
      } else {
        byEmail.set(email, [sharedConversation]);
      }
    }

    for (const name of seenNames) {
      const entries = byName.get(name);
      if (entries) {
        entries.push(sharedConversation);
      } else {
        byName.set(name, [sharedConversation]);
      }
    }
  }

  for (const [mri, email] of Object.entries(safeEmailByMri)) {
    const normalizedEmail = normalizeProfileText(email);
    if (!normalizedEmail) continue;
    const entries = byEmail.get(normalizedEmail);
    if (!entries?.length) continue;
    const target = byMri.get(mri) ?? [];
    const seenConversationIds = new Set(target.map((entry) => entry.id));
    for (const entry of entries) {
      if (seenConversationIds.has(entry.id)) continue;
      target.push(entry);
      seenConversationIds.add(entry.id);
    }
    byMri.set(mri, target);
  }

  for (const [mri, displayName] of Object.entries(safeDisplayNameByMri)) {
    const normalizedName = normalizeProfileText(displayName);
    if (!normalizedName) continue;
    const entries = byName.get(normalizedName);
    if (!entries?.length) continue;
    const target = byMri.get(mri) ?? [];
    const seenConversationIds = new Set(target.map((entry) => entry.id));
    for (const entry of entries) {
      if (seenConversationIds.has(entry.id)) continue;
      target.push(entry);
      seenConversationIds.add(entry.id);
    }
    byMri.set(mri, target);
  }

  return Object.fromEntries(byMri);
}
