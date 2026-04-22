import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";

type WorkspaceShellAccount = {
  upn?: string;
  tenantId?: string;
};

type WorkspaceShellSession = {
  upn?: string;
  tenantId?: string;
  skypeId?: string;
  expiresAt?: string | null;
  region?: string | null;
};

type WorkspaceShellTenant = {
  session?: WorkspaceShellSession;
  conversations?: unknown[];
};

type WorkspaceShell = {
  accounts?: WorkspaceShellAccount[];
  tenants?: Record<string, WorkspaceShellTenant>;
};

type ProfileRow = {
  mri: string;
  avatar?: string | null;
  avatar_thumb?: string | null;
  avatar_full?: string | null;
  display_name?: string | null;
  email?: string | null;
  job_title?: string | null;
  department?: string | null;
  company_name?: string | null;
  tenant_name?: string | null;
  location?: string | null;
};

const appSupportDir = path.join(os.homedir(), "Library/Application Support");
const cacheRoots = [
  path.join(appSupportDir, "com.betterteams.app"),
  path.join(appSupportDir, "Better Teams"),
];

export function getPersistedAccounts(): unknown[] {
  for (const root of cacheRoots) {
    const shell = readWorkspaceShell(
      path.join(root, "teams-workspace-shell.json"),
    );
    if (Array.isArray(shell?.accounts) && shell.accounts.length > 0) {
      return shell.accounts;
    }
  }
  return [];
}

export function getPersistedSession(
  tenantId?: string | null,
): WorkspaceShellSession | null {
  for (const root of cacheRoots) {
    const shell = readWorkspaceShell(
      path.join(root, "teams-workspace-shell.json"),
    );
    const tenants = shell?.tenants;
    if (!tenants) continue;
    if (tenantId && tenants[tenantId]?.session) {
      return tenants[tenantId].session;
    }
    if (!tenantId) {
      const session = Object.values(tenants).find(
        (tenant) => tenant.session,
      )?.session;
      if (session) return session;
    }
  }
  return null;
}

export function getPersistedConversations(tenantId?: string | null): unknown[] {
  if (!tenantId) return [];
  for (const root of cacheRoots) {
    const shell = readWorkspaceShell(
      path.join(root, "teams-workspace-shell.json"),
    );
    const conversations = shell?.tenants?.[tenantId]?.conversations;
    if (Array.isArray(conversations) && conversations.length > 0) {
      const session = shell?.tenants?.[tenantId]?.session;
      return conversations.map((conversation) =>
        hydrateConversationParticipants(
          conversation,
          session,
          path.join(root, "better-teams-cache.db"),
        ),
      );
    }
  }
  return [];
}

export function getPersistedMessages(
  tenantId?: string | null,
  conversationId?: string | null,
): unknown | null {
  if (!tenantId || !conversationId) return null;
  for (const root of cacheRoots) {
    const value = readThreadCache(
      path.join(root, "better-teams-cache.db"),
      tenantId,
      conversationId,
    );
    if (value) return value;
  }
  return null;
}

export function getPersistedProfilePresentation(mris: string[]): unknown {
  const unique = [
    ...new Set(mris.map((mri) => mri.trim()).filter((mri) => mri.length > 0)),
  ];
  const avatarThumbs: Record<string, string> = {};
  const avatarFull: Record<string, string> = {};
  const displayNames: Record<string, string> = {};
  const emails: Record<string, string> = {};
  const jobTitles: Record<string, string> = {};
  const departments: Record<string, string> = {};
  const companyNames: Record<string, string> = {};
  const tenantNames: Record<string, string> = {};
  const locations: Record<string, string> = {};

  const missing = new Set(unique.map(normalizeMri));
  for (const root of cacheRoots) {
    if (missing.size === 0) break;
    const dbPath = path.join(root, "better-teams-cache.db");
    for (const mri of [...missing]) {
      const row = readProfileRow(dbPath, mri);
      if (!row) continue;
      const key = normalizeMri(row.mri);
      const thumb = nonEmpty(row.avatar_thumb) ?? nonEmpty(row.avatar);
      const full = nonEmpty(row.avatar_full) ?? thumb;
      if (thumb) avatarThumbs[key] = thumb;
      if (full) avatarFull[key] = full;
      const displayName = nonEmpty(row.display_name);
      if (displayName) displayNames[key] = displayName;
      const email = nonEmpty(row.email);
      if (email) emails[key] = email;
      const jobTitle = nonEmpty(row.job_title);
      if (jobTitle) jobTitles[key] = jobTitle;
      const department = nonEmpty(row.department);
      if (department) departments[key] = department;
      const companyName = nonEmpty(row.company_name);
      if (companyName) companyNames[key] = companyName;
      const tenantName = nonEmpty(row.tenant_name);
      if (tenantName) tenantNames[key] = tenantName;
      const location = nonEmpty(row.location);
      if (location) locations[key] = location;
      missing.delete(mri);
    }
  }

  return {
    avatarThumbs,
    avatarFull,
    displayNames,
    emails,
    jobTitles,
    departments,
    companyNames,
    tenantNames,
    locations,
  };
}

function readWorkspaceShell(filePath: string): WorkspaceShell | null {
  try {
    if (!fs.existsSync(filePath)) return null;
    const parsed = JSON.parse(fs.readFileSync(filePath, "utf8")) as unknown;
    return isWorkspaceShell(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

function readThreadCache(
  dbPath: string,
  tenantId: string,
  conversationId: string,
): unknown | null {
  if (!fs.existsSync(dbPath)) return null;
  let db: Database.Database | null = null;
  try {
    db = new Database(dbPath, { readonly: true, fileMustExist: true });
    const row = db
      .prepare(
        `SELECT value
         FROM thread_cache
         WHERE tenant_id = ? AND conversation_id = ?`,
      )
      .get(tenantId, conversationId) as { value?: string } | undefined;
    if (!row?.value) return null;
    return JSON.parse(row.value) as unknown;
  } catch {
    return null;
  } finally {
    db?.close();
  }
}

function readProfileRow(dbPath: string, mri: string): ProfileRow | null {
  if (!fs.existsSync(dbPath)) return null;
  let db: Database.Database | null = null;
  try {
    db = new Database(dbPath, { readonly: true, fileMustExist: true });
    return (
      (db
        .prepare("SELECT * FROM profile_cache WHERE lower(mri) = lower(?)")
        .get(mri) as ProfileRow | undefined) ?? null
    );
  } catch {
    return null;
  } finally {
    db?.close();
  }
}

function isWorkspaceShell(value: unknown): value is WorkspaceShell {
  return (
    typeof value === "object" &&
    value !== null &&
    "tenants" in value &&
    typeof (value as WorkspaceShell).tenants === "object"
  );
}

function nonEmpty(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function normalizeMri(mri: string): string {
  return mri.trim().toLowerCase();
}

function mriFromSkypeId(skypeId?: string | null): string | undefined {
  const value = nonEmpty(skypeId);
  if (!value) return undefined;
  return value.startsWith("8:") ? value : `8:${value}`;
}

function orgIdMriFromGuid(value: string): string | undefined {
  const guid = value.trim();
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(guid)) {
    return undefined;
  }
  return `8:orgid:${guid}`;
}

function pairMrisFromConversationId(conversationId: string): string[] {
  const decoded = (() => {
    try {
      return decodeURIComponent(conversationId.trim());
    } catch {
      return conversationId.trim();
    }
  })();
  const match = decoded.match(/^19:([^@]+)@unq\.gbl\.spaces$/i);
  if (!match) return [];
  return match[1]
    .split("_")
    .map((part) => orgIdMriFromGuid(part) ?? nonEmpty(part))
    .filter((mri): mri is string => Boolean(mri));
}

function hydrateConversationParticipants(
  conversation: unknown,
  session: WorkspaceShellSession | undefined,
  dbPath: string,
): unknown {
  if (!conversation || typeof conversation !== "object") return conversation;
  const record = conversation as Record<string, unknown>;
  if (Array.isArray(record.members) && record.members.length > 0) {
    return conversation;
  }
  const id = nonEmpty(record.id);
  if (!id) return conversation;
  const mris = [
    ...pairMrisFromConversationId(id),
    ...inferMrisFromConversation(record, session),
  ];
  if (mris.length === 0) return conversation;
  const selfMri = mriFromSkypeId(session?.skypeId);
  const ordered = selfMri
    ? [
        ...mris.filter((mri) => normalizeMri(mri) === normalizeMri(selfMri)),
        ...mris.filter((mri) => normalizeMri(mri) !== normalizeMri(selfMri)),
      ]
    : mris;
  const seen = new Set<string>();
  return {
    ...record,
    members: ordered.flatMap((mri) => {
      const key = normalizeMri(mri);
      if (seen.has(key)) return [];
      seen.add(key);
      const profile = readProfileRow(dbPath, mri);
      const fallbackName = displayNameFromConversationForMri(record, mri);
      return [
        {
          id: mri,
          role: "User",
          isMri: true,
          displayName: nonEmpty(profile?.display_name) ?? fallbackName,
        },
      ];
    }),
  };
}

function inferMrisFromConversation(
  conversation: Record<string, unknown>,
  session: WorkspaceShellSession | undefined,
): string[] {
  const out: string[] = [];
  const properties =
    conversation.properties && typeof conversation.properties === "object"
      ? (conversation.properties as Record<string, unknown>)
      : {};
  const lastMessage =
    conversation.lastMessage && typeof conversation.lastMessage === "object"
      ? (conversation.lastMessage as Record<string, unknown>)
      : {};
  const push = (value: unknown) => {
    const mri = nonEmpty(value);
    if (mri) out.push(mri);
  };
  push(mriFromSkypeId(session?.skypeId));
  push(properties.addedBy);
  push(lastMessage.fromMri);
  push(lastMessage.from);
  return out;
}

function displayNameFromConversationForMri(
  conversation: Record<string, unknown>,
  mri: string,
): string | undefined {
  const lastMessage =
    conversation.lastMessage && typeof conversation.lastMessage === "object"
      ? (conversation.lastMessage as Record<string, unknown>)
      : {};
  const from = nonEmpty(lastMessage.fromMri) ?? nonEmpty(lastMessage.from);
  if (from && normalizeMri(from) === normalizeMri(mri)) {
    return (
      nonEmpty(lastMessage.imdisplayname) ??
      nonEmpty(lastMessage.senderDisplayName)
    );
  }
  return undefined;
}
