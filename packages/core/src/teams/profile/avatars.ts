import { conversationChatKind, isSelfMessage } from "../../chat";
import { extractTeamsMri, nonEmptyTrimmedString } from "../normalize";
import {
  parseShortProfileRows,
  type ShortProfileRow,
  ShortProfileRowSchema,
} from "../schemas";
import type { Conversation, Message } from "../types";

type NormalizedShortProfile = {
  mri: string | null;
  imageUrl: string | null;
  displayName: string | null;
  email: string | null;
  jobTitle: string | null;
  department: string | null;
  companyName: string | null;
  tenantName: string | null;
  location: string | null;
};

function extractAvatarMriLike(input: string): string | null {
  return extractTeamsMri(input);
}

function guidLikeToOrgidMri(input: string): string | null {
  const t = input.trim();
  if (
    !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(t)
  ) {
    return null;
  }
  return `8:orgid:${t}`;
}

function extractAvatarMrisFromConversationId(conversationId: string): string[] {
  const out = new Set<string>();
  const trimmed = conversationId.trim();
  if (!trimmed) return [];
  const decoded = (() => {
    try {
      return decodeURIComponent(trimmed);
    } catch {
      return trimmed;
    }
  })();
  const direct = decoded.match(/8:[^/?#&\s"']+/gi) ?? [];
  for (const m of direct) out.add(m);
  const m = decoded.match(/^19:([^@]+)@unq\.gbl\.spaces$/i);
  if (m) {
    const pair = m[1]
      .split("_")
      .map((s) => s.trim())
      .filter(Boolean);
    for (const p of pair) {
      const maybeMri = extractAvatarMriLike(p) ?? guidLikeToOrgidMri(p);
      if (maybeMri) out.add(maybeMri);
    }
  }
  return [...out];
}

export function canonAvatarMri(mri: string): string {
  const extracted = extractAvatarMriLike(mri);
  return (extracted ?? mri).trim().toLowerCase();
}

export function dmConversationAvatarMri(
  c: Conversation,
  selfSkypeId?: string,
): string | undefined {
  const normalizedSelf =
    selfSkypeId == null
      ? undefined
      : canonAvatarMri(
          selfSkypeId.startsWith("8:") ? selfSkypeId : `8:${selfSkypeId}`,
        );
  if (conversationChatKind(c) !== "dm") return undefined;
  const members = c.members;
  if (members?.length) {
    for (const m of members) {
      const id = m.id ? extractAvatarMriLike(m.id) : null;
      if (!id) continue;
      if (isSelfMessage(id, normalizedSelf)) continue;
      return id;
    }
  }
  const lm = c.lastMessage;
  const from = lm?.fromMri
    ? extractAvatarMriLike(lm.fromMri)
    : lm?.from
      ? extractAvatarMriLike(lm.from)
      : null;
  if (from && !isSelfMessage(from, normalizedSelf)) {
    return from;
  }
  const fromConversationId = extractAvatarMrisFromConversationId(c.id);
  for (const id of fromConversationId) {
    if (isSelfMessage(id, normalizedSelf)) continue;
    return id;
  }
  return undefined;
}

export function collectProfileAvatarMris(input: {
  conversations: Conversation[];
  messages: Message[];
  selfSkypeId?: string;
}): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  const push = (mriLike: string | undefined) => {
    if (!mriLike) return;
    const parsed = extractAvatarMriLike(mriLike);
    if (!parsed) return;
    const c = canonAvatarMri(parsed);
    if (seen.has(c)) return;
    seen.add(c);
    out.push(parsed.trim());
  };

  if (input.selfSkypeId) {
    push(
      input.selfSkypeId.startsWith("8:")
        ? input.selfSkypeId
        : `8:${input.selfSkypeId}`,
    );
  }

  for (const c of input.conversations) {
    push(dmConversationAvatarMri(c, input.selfSkypeId));
    if (Array.isArray(c.members)) {
      for (const member of c.members) {
        push(member.id);
      }
    }
    for (const id of extractAvatarMrisFromConversationId(c.id)) {
      push(id);
    }
    push(c.lastMessage?.fromMri ?? c.lastMessage?.from);
  }

  for (const m of input.messages) {
    push(m.fromMri ?? m.from);
  }

  return out;
}

export function normalizeFetchShortProfileRows(raw: unknown): unknown[] {
  if (Array.isArray(raw)) return raw;
  if (raw && typeof raw === "object") {
    const o = raw as Record<string, unknown>;
    if (Array.isArray(o.value)) return o.value;
    if (Array.isArray(o.profiles)) return o.profiles;
    if (Array.isArray(o.users)) return o.users;
    if (Array.isArray(o.shortProfiles)) return o.shortProfiles;
  }
  return parseShortProfileRows(raw);
}

function profileRecord(row: unknown): ShortProfileRow | null {
  const top = ShortProfileRowSchema.safeParse(row);
  if (top.success) return top.data;
  if (row && typeof row === "object" && "shortProfile" in row) {
    const nested = ShortProfileRowSchema.safeParse(
      (row as { shortProfile?: unknown }).shortProfile,
    );
    if (nested.success) return nested.data;
  }
  return null;
}

function normalizeShortProfileRecord(
  row: unknown,
): NormalizedShortProfile | null {
  const o = profileRecord(row);
  if (!o) return null;
  return {
    mri: pickMri(o),
    imageUrl: pickImageUrl(o),
    displayName: pickDisplayNameFromProfile(o),
    email: pickEmailFromProfile(o),
    jobTitle: pickJobTitleFromProfile(o),
    department: pickDepartmentFromProfile(o),
    companyName: pickCompanyNameFromProfile(o),
    tenantName: pickTenantNameFromProfile(o),
    location: pickLocationFromProfile(o),
  };
}

export function shortProfileRowToMriAndImageUrl(
  row: unknown,
): { mri: string; imageUrl: string } | null {
  const normalized = normalizeShortProfileRecord(row);
  if (!normalized) return null;
  const { mri, imageUrl } = normalized;
  if (!mri) return null;
  if (!imageUrl) return null;
  return { mri, imageUrl };
}

export function mriFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.mri ?? null;
}

function pickDisplayNameFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null =>
    nonEmptyTrimmedString(v) ?? null;
  const dn =
    asTrimmed(o.displayName) ?? asTrimmed(o.displayname) ?? asTrimmed(o.name);
  if (dn) return dn;
  const gn = asTrimmed(o.givenName) ?? "";
  const sn = asTrimmed(o.surname) ?? "";
  const combined = [gn, sn].filter(Boolean).join(" ");
  if (combined) return combined;
  const upn = asTrimmed(o.userPrincipalName) ?? asTrimmed(o.mail);
  if (upn?.includes("@")) {
    const local = upn.split("@")[0]?.trim();
    if (local) return local;
  }
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickDisplayNameFromProfile(nested.data);
  }
  return null;
}

function pickEmailFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null => {
    const trimmed = nonEmptyTrimmedString(v);
    return trimmed?.includes("@") ? trimmed : null;
  };
  const email =
    asTrimmed(o.email) ?? asTrimmed(o.mail) ?? asTrimmed(o.userPrincipalName);
  if (email) return email;
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickEmailFromProfile(nested.data);
  }
  const sp = o.shortProfile;
  if (sp) {
    const nested = ShortProfileRowSchema.safeParse(sp);
    if (nested.success) return pickEmailFromProfile(nested.data);
  }
  return null;
}

function pickJobTitleFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null =>
    nonEmptyTrimmedString(v) ?? null;
  const raw = o as Record<string, unknown>;
  const title =
    asTrimmed(raw.jobTitle) ??
    asTrimmed(raw.title) ??
    asTrimmed(raw.department);
  if (title) return title;
  const st = o.skypeTeamsInfo;
  if (st && typeof st === "object") {
    const stRec = st as Record<string, unknown>;
    const nested =
      asTrimmed(stRec.jobTitle) ??
      asTrimmed(stRec.title) ??
      asTrimmed(stRec.department);
    if (nested) return nested;
  }
  const sp = o.shortProfile;
  if (sp && typeof sp === "object") {
    const spRec = sp as Record<string, unknown>;
    const nested =
      asTrimmed(spRec.jobTitle) ??
      asTrimmed(spRec.title) ??
      asTrimmed(spRec.department);
    if (nested) return nested;
  }
  return null;
}

function pickDepartmentFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null =>
    nonEmptyTrimmedString(v) ?? null;
  const department = asTrimmed(o.department);
  if (department) return department;
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickDepartmentFromProfile(nested.data);
  }
  const sp = o.shortProfile;
  if (sp) {
    const nested = ShortProfileRowSchema.safeParse(sp);
    if (nested.success) return pickDepartmentFromProfile(nested.data);
  }
  return null;
}

function pickCompanyNameFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null =>
    nonEmptyTrimmedString(v) ?? null;
  const companyName = asTrimmed(o.companyName);
  if (companyName) return companyName;
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickCompanyNameFromProfile(nested.data);
  }
  const sp = o.shortProfile;
  if (sp) {
    const nested = ShortProfileRowSchema.safeParse(sp);
    if (nested.success) return pickCompanyNameFromProfile(nested.data);
  }
  return null;
}

function pickTenantNameFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null =>
    nonEmptyTrimmedString(v) ?? null;
  const tenantName = asTrimmed(o.tenantName);
  if (tenantName) return tenantName;
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickTenantNameFromProfile(nested.data);
  }
  const sp = o.shortProfile;
  if (sp) {
    const nested = ShortProfileRowSchema.safeParse(sp);
    if (nested.success) return pickTenantNameFromProfile(nested.data);
  }
  return null;
}

function pickLocationFromProfile(o: ShortProfileRow): string | null {
  const asTrimmed = (v: unknown): string | null =>
    nonEmptyTrimmedString(v) ?? null;
  const location = asTrimmed(o.userLocation);
  if (location) return location;
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickLocationFromProfile(nested.data);
  }
  const sp = o.shortProfile;
  if (sp) {
    const nested = ShortProfileRowSchema.safeParse(sp);
    if (nested.success) return pickLocationFromProfile(nested.data);
  }
  return null;
}

export function emailFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.email ?? null;
}

export function jobTitleFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.jobTitle ?? null;
}

export function displayNameFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.displayName ?? null;
}

export function departmentFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.department ?? null;
}

export function companyNameFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.companyName ?? null;
}

export function tenantNameFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.tenantName ?? null;
}

export function locationFromShortProfileRow(row: unknown): string | null {
  return normalizeShortProfileRecord(row)?.location ?? null;
}

export function applyProfileDisplayNameToRowMrIs(
  row: unknown,
  displayName: string,
  setForMri: (mri: string, label: string) => void,
): void {
  const trimmed = displayName.trim().replace(/\s+/g, " ");
  if (!trimmed) return;
  const fromJson = collectSkypeMriLikeStringsFromJson(row);
  const primary = mriFromShortProfileRow(row);
  const keys = new Set<string>(fromJson);
  if (primary) keys.add(primary);
  if (keys.size === 0) return;
  for (const mri of keys) {
    setForMri(mri, trimmed);
  }
}

export function collectSkypeMriLikeStringsFromJson(root: unknown): string[] {
  const found = new Set<string>();
  const walk = (x: unknown, depth: number) => {
    if (depth > 18 || x == null) return;
    if (typeof x === "string") {
      const t = extractAvatarMriLike(x);
      if (t && t.length > 4) {
        found.add(t);
      }
      return;
    }
    if (typeof x !== "object") return;
    if (Array.isArray(x)) {
      for (const y of x) walk(y, depth + 1);
      return;
    }
    for (const v of Object.values(x)) walk(v, depth + 1);
  };
  walk(root, 0);
  return [...found];
}

export function applyProfilePhotoDataUrlToRowMrIs(
  row: unknown,
  dataUrl: string,
  setForMri: (mri: string, data: string) => void,
): void {
  const fromJson = collectSkypeMriLikeStringsFromJson(row);
  const primary = mriFromShortProfileRow(row);
  const keys = new Set<string>(fromJson);
  if (primary) keys.add(primary);
  if (keys.size === 0) return;
  for (const mri of keys) {
    setForMri(mri, dataUrl);
  }
}

function pickMri(o: ShortProfileRow): string | null {
  const a = o.mri ?? o.userMri;
  if (typeof a === "string") {
    const parsed = extractAvatarMriLike(a);
    if (parsed) return parsed;
  }
  const st = o.skypeTeamsInfo;
  if (st) {
    const stO = ShortProfileRowSchema.safeParse(st);
    const b = stO.success ? (stO.data.userMri ?? stO.data.mri) : null;
    if (typeof b === "string") {
      const parsed = extractAvatarMriLike(b);
      if (parsed) return parsed;
    }
  }
  return null;
}

function pickImageUrl(o: ShortProfileRow): string | null {
  const keys = [
    "imageUri",
    "imageURL",
    "profileImageUri",
    "profilePictureUri",
    "profileImageUrl",
    "avatarUrl",
    "avatarURL",
    "pictureUrl",
    "highResolutionImageUrl",
    "linkedInProfilePictureUrl",
  ];
  for (const k of keys) {
    const v = o[k];
    if (typeof v === "string" && v.length > 0) {
      if (
        v.startsWith("http://") ||
        v.startsWith("https://") ||
        v.startsWith("/")
      ) {
        return v;
      }
    }
  }
  const st = o.skypeTeamsInfo;
  if (st) {
    const nested = ShortProfileRowSchema.safeParse(st);
    if (nested.success) return pickImageUrl(nested.data);
  }
  return null;
}
