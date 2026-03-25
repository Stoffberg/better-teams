import type {
  ConversationChatKind,
  MessageAttachment,
  MessageInlinePart,
  MessageReference,
} from "@/lib/chat-format";
import type { Conversation, Message } from "@/services/teams/types";

export type SidebarSectionId = "meetings" | "groups" | "direct";
export type SidebarSectionState = Record<SidebarSectionId, boolean>;

export type SidebarConversationItem = {
  id: string;
  conversation: Conversation;
  title: string;
  preview: string;
  kind: ConversationChatKind;
  isFavorite: boolean;
  avatarMri?: string;
  avatarThumbSrc?: string;
  avatarFullSrc?: string;
  sideTime: string;
  searchText: string;
};

export type ReadStatus = "sending" | "sent" | "delivered" | "read";

export type ReceiptPerson = {
  mri: string;
  name: string;
  readAt?: string;
};

export type DisplayMessage = {
  message: Message;
  parts: {
    quote: MessageInlinePart[] | null;
    body: MessageInlinePart[];
    quoteRef: MessageReference | null;
    attachments: MessageAttachment[];
  };
  displayName: string;
  time: string;
  self: boolean;
  deleted: boolean;
  edited: boolean;
  readStatus?: ReadStatus;
  sentAt?: string;
  readAt?: string;
  receiptScope?: "dm" | "group";
  receiptSeenBy?: ReceiptPerson[];
  receiptUnseenBy?: ReceiptPerson[];
  bodyPreview: string;
  searchText: string;
};

export type MessageBlock =
  | { kind: "day"; label: string; key: string }
  | {
      kind: "msg";
      entry: DisplayMessage;
      messageIndex: number;
      showMeta: boolean;
      key: string;
    };

export const DEFAULT_SECTIONS: SidebarSectionState = {
  meetings: true,
  groups: true,
  direct: true,
};

export const SIDEBAR_SECTIONS_STORAGE_KEY =
  "better-teams-chat-sidebar-sections";
export const THREAD_PAGE = 80;
export const OLDER_LOAD_THROTTLE_MS = 75;
