import { removeCachedImageFiles } from "@/lib/electron-bridge";
import Database from "@/lib/electron-sqlite";
import { canonAvatarMri } from "@/lib/teams-profile-avatars";
import type { ThreadQueryData } from "@/lib/teams-thread-query";
import type {
  Conversation,
  TeamsAccountOption,
  TeamsProfilePresentation,
  TeamsSessionInfo,
  TeamsWorkspaceShellSnapshot,
  TeamsWorkspaceShellTenantSnapshot,
} from "@/services/teams/types";

let dbPromise: Promise<Database> | null = null;
let writeQueue = Promise.resolve();

function isLockedSqliteError(error: unknown): boolean {
  return (
    error instanceof Error &&
    /database is locked|code:\s*5/i.test(error.message)
  );
}

async function sleep(ms: number): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, ms));
}

async function withLockedRetry<T>(operation: () => Promise<T>): Promise<T> {
  let attempt = 0;
  while (true) {
    try {
      return await operation();
    } catch (error) {
      if (!isLockedSqliteError(error) || attempt >= 5) {
        throw error;
      }
      attempt += 1;
      await sleep(20 * 2 ** (attempt - 1));
    }
  }
}

async function enqueueWrite<T>(operation: () => Promise<T>): Promise<T> {
  const run = writeQueue.then(() => withLockedRetry(operation));
  writeQueue = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

function enqueueWriteSafely(operation: () => Promise<unknown>): void {
  void enqueueWrite(operation).catch(() => undefined);
}

async function getDb(): Promise<Database> {
  if (!dbPromise) {
    dbPromise = Database.load("sqlite:better-teams-cache.db").then(
      async (db) => {
        await initSchema(db);
        return db;
      },
    );
  }
  return dbPromise;
}

async function initSchema(db: Database): Promise<void> {
  await db.execute("PRAGMA journal_mode = WAL");
  await db.execute("PRAGMA busy_timeout = 5000");
  await db.execute(`
    CREATE TABLE IF NOT EXISTS profile_cache (
      mri TEXT PRIMARY KEY,
      avatar TEXT,
      avatar_thumb TEXT,
      avatar_full TEXT,
      display_name TEXT,
      email TEXT,
      job_title TEXT,
      department TEXT,
      company_name TEXT,
      tenant_name TEXT,
      location TEXT,
      avatar_fetched_at INTEGER,
      profile_fetched_at INTEGER,
      last_requested_at INTEGER NOT NULL
    )
  `);
  const columns = await db.select<Array<{ name: string }>>(
    "PRAGMA table_info(profile_cache)",
  );
  const columnNames = new Set(columns.map((column) => column.name));
  if (!columnNames.has("avatar_thumb")) {
    await db.execute("ALTER TABLE profile_cache ADD COLUMN avatar_thumb TEXT");
  }
  if (!columnNames.has("avatar_full")) {
    await db.execute("ALTER TABLE profile_cache ADD COLUMN avatar_full TEXT");
  }
  if (!columnNames.has("department")) {
    await db.execute("ALTER TABLE profile_cache ADD COLUMN department TEXT");
  }
  if (!columnNames.has("company_name")) {
    await db.execute("ALTER TABLE profile_cache ADD COLUMN company_name TEXT");
  }
  if (!columnNames.has("tenant_name")) {
    await db.execute("ALTER TABLE profile_cache ADD COLUMN tenant_name TEXT");
  }
  if (!columnNames.has("location")) {
    await db.execute("ALTER TABLE profile_cache ADD COLUMN location TEXT");
  }
  await db.execute(`
    UPDATE profile_cache
    SET avatar = NULL,
        avatar_thumb = NULL,
        avatar_full = NULL,
        avatar_fetched_at = NULL
    WHERE avatar LIKE 'data:%'
       OR avatar_thumb LIKE 'data:%'
       OR avatar_full LIKE 'data:%'
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS image_cache (
      url TEXT PRIMARY KEY,
      data_url TEXT NOT NULL,
      updated_at INTEGER NOT NULL,
      last_used_at INTEGER NOT NULL
    )
  `);
  const imageColumns = await db.select<Array<{ name: string }>>(
    "PRAGMA table_info(image_cache)",
  );
  const imageColumnNames = new Set(imageColumns.map((column) => column.name));
  if (!imageColumnNames.has("file_path")) {
    await db.execute("ALTER TABLE image_cache ADD COLUMN file_path TEXT");
  }
  await db.execute(
    "DELETE FROM image_cache WHERE file_path IS NULL OR data_url LIKE 'data:%'",
  );
  await db.execute(`
    CREATE TABLE IF NOT EXISTS workspace_shell (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL,
      updated_at INTEGER NOT NULL
    )
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS query_cache (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL,
      updated_at INTEGER NOT NULL
    )
  `);
  await db.execute(`
    CREATE TABLE IF NOT EXISTS thread_cache (
      tenant_id TEXT NOT NULL,
      conversation_id TEXT NOT NULL,
      value TEXT NOT NULL,
      updated_at INTEGER NOT NULL,
      last_requested_at INTEGER NOT NULL,
      PRIMARY KEY (tenant_id, conversation_id)
    )
  `);
}

const MAX_PROFILE_ENTRIES = 5_000;
const PROFILE_TTL_MS = 7 * 24 * 60 * 60 * 1000;
const NEGATIVE_AVATAR_TTL_MS = 12 * 60 * 60 * 1000;

function emptyPresentation(): TeamsProfilePresentation {
  return {
    avatarThumbs: {},
    avatarFull: {},
    displayNames: {},
    emails: {},
    jobTitles: {},
    departments: {},
    companyNames: {},
    tenantNames: {},
    locations: {},
  };
}

export type CachedProfileLookup = {
  presentation: TeamsProfilePresentation;
  missingMris: string[];
};

type ProfileRow = {
  mri: string;
  avatar: string | null;
  avatar_thumb: string | null;
  avatar_full: string | null;
  display_name: string | null;
  email: string | null;
  job_title: string | null;
  department: string | null;
  company_name: string | null;
  tenant_name: string | null;
  location: string | null;
  avatar_fetched_at: number | null;
  profile_fetched_at: number | null;
  last_requested_at: number;
};

export const SqliteProfileCache = {
  async lookupProfiles(mris: string[]): Promise<CachedProfileLookup> {
    const db = await getDb();
    const presentation = emptyPresentation();
    const missingMris: string[] = [];
    const now = Date.now();

    // Canonicalize all MRIs up front
    const canonMris = mris.map((raw) => canonAvatarMri(raw));

    if (canonMris.length === 0) {
      return { presentation, missingMris };
    }

    // Batch SELECT: fetch all profiles in a single query
    const placeholders = canonMris.map((_, i) => `$${i + 1}`).join(", ");
    const rows = await db.select<ProfileRow[]>(
      `SELECT * FROM profile_cache WHERE mri IN (${placeholders})`,
      canonMris,
    );

    // Index results by MRI for quick lookup
    const rowsByMri = new Map<string, ProfileRow>();
    for (const row of rows) {
      rowsByMri.set(row.mri, row);
    }

    const foundMris: string[] = [];

    for (const mri of canonMris) {
      const entry = rowsByMri.get(mri);
      if (!entry) {
        missingMris.push(mri);
        continue;
      }

      foundMris.push(mri);

      if (entry.avatar_thumb ?? entry.avatar) {
        presentation.avatarThumbs[mri] =
          entry.avatar_thumb ?? entry.avatar ?? "";
      }
      if (entry.avatar_full ?? entry.avatar) {
        presentation.avatarFull[mri] = entry.avatar_full ?? entry.avatar ?? "";
      }
      if (entry.display_name)
        presentation.displayNames[mri] = entry.display_name;
      if (entry.email) presentation.emails[mri] = entry.email;
      if (entry.job_title) presentation.jobTitles[mri] = entry.job_title;
      if (entry.department) presentation.departments[mri] = entry.department;
      if (entry.company_name)
        presentation.companyNames[mri] = entry.company_name;
      if (entry.tenant_name) presentation.tenantNames[mri] = entry.tenant_name;
      if (entry.location) presentation.locations[mri] = entry.location;

      const avatarFresh =
        typeof entry.avatar === "string" ||
        (typeof entry.avatar_fetched_at === "number" &&
          now - entry.avatar_fetched_at < NEGATIVE_AVATAR_TTL_MS);
      const profileFresh =
        typeof entry.profile_fetched_at === "number" &&
        now - entry.profile_fetched_at < PROFILE_TTL_MS;

      if (!avatarFresh || !profileFresh) {
        missingMris.push(mri);
      }
    }

    if (foundMris.length > 0) {
      const updatePlaceholders = foundMris
        .map((_, i) => `$${i + 2}`)
        .join(", ");
      enqueueWriteSafely(() =>
        db.execute(
          `UPDATE profile_cache SET last_requested_at = $1 WHERE mri IN (${updatePlaceholders})`,
          [now, ...foundMris],
        ),
      );
    }

    return { presentation, missingMris };
  },

  async storeProfiles(
    requestedMris: string[],
    incoming: TeamsProfilePresentation,
  ): Promise<void> {
    const db = await getDb();
    const now = Date.now();

    await enqueueWrite(async () => {
      for (const rawMri of requestedMris) {
        const mri = canonAvatarMri(rawMri);
        const avatarThumb = incoming.avatarThumbs[mri] ?? null;
        const avatarFull = incoming.avatarFull[mri] ?? null;
        const displayName = incoming.displayNames[mri] ?? null;
        const email = incoming.emails[mri] ?? null;
        const jobTitle = incoming.jobTitles[mri] ?? null;
        const department = incoming.departments[mri] ?? null;
        const companyName = incoming.companyNames[mri] ?? null;
        const tenantName = incoming.tenantNames[mri] ?? null;
        const location = incoming.locations[mri] ?? null;

        await db.execute(
          `INSERT INTO profile_cache (mri, avatar, avatar_thumb, avatar_full, display_name, email, job_title, department, company_name, tenant_name, location, avatar_fetched_at, profile_fetched_at, last_requested_at)
             VALUES ($1, $2, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $11, $11)
             ON CONFLICT(mri) DO UPDATE SET
               avatar = COALESCE($2, avatar),
               avatar_thumb = COALESCE($2, avatar_thumb),
               avatar_full = COALESCE($3, avatar_full),
               display_name = COALESCE($4, display_name),
               email = COALESCE($5, email),
               job_title = COALESCE($6, job_title),
               department = COALESCE($7, department),
               company_name = COALESCE($8, company_name),
               tenant_name = COALESCE($9, tenant_name),
               location = COALESCE($10, location),
               avatar_fetched_at = $11,
               profile_fetched_at = $11,
               last_requested_at = $11`,
          [
            mri,
            avatarThumb,
            avatarFull,
            displayName,
            email,
            jobTitle,
            department,
            companyName,
            tenantName,
            location,
            now,
          ],
        );
      }

      await db.execute(
        `DELETE FROM profile_cache WHERE mri NOT IN (
          SELECT mri FROM profile_cache ORDER BY last_requested_at DESC LIMIT $1
        )`,
        [MAX_PROFILE_ENTRIES],
      );
    });
  },

  merge(
    cached: TeamsProfilePresentation,
    fresh: TeamsProfilePresentation,
  ): TeamsProfilePresentation {
    return {
      avatarThumbs: { ...cached.avatarThumbs, ...fresh.avatarThumbs },
      avatarFull: { ...cached.avatarFull, ...fresh.avatarFull },
      displayNames: { ...cached.displayNames, ...fresh.displayNames },
      emails: { ...cached.emails, ...fresh.emails },
      jobTitles: { ...cached.jobTitles, ...fresh.jobTitles },
      departments: { ...cached.departments, ...fresh.departments },
      companyNames: { ...cached.companyNames, ...fresh.companyNames },
      tenantNames: { ...cached.tenantNames, ...fresh.tenantNames },
      locations: { ...cached.locations, ...fresh.locations },
    };
  },
};

const MAX_IMAGE_ENTRIES = 2_000;

const _pendingImageLastUsedUrls = new Set<string>();
let _imageLastUsedTimer: ReturnType<typeof setTimeout> | null = null;

function scheduleImageLastUsedFlush(): void {
  if (_imageLastUsedTimer !== null) return;
  _imageLastUsedTimer = setTimeout(async () => {
    _imageLastUsedTimer = null;
    const urls = [..._pendingImageLastUsedUrls];
    _pendingImageLastUsedUrls.clear();
    if (urls.length === 0) return;
    try {
      const db = await getDb();
      const now = Date.now();
      const placeholders = urls.map((_, i) => `$${i + 2}`).join(", ");
      await enqueueWrite(() =>
        db.execute(
          `UPDATE image_cache SET last_used_at = $1 WHERE url IN (${placeholders})`,
          [now, ...urls],
        ),
      );
    } catch {
      _pendingImageLastUsedUrls.clear();
    }
  }, 5_000);
}

let _imageCacheInsertCount = 0;

type ImageRow = {
  url: string;
  file_path: string | null;
  updated_at: number;
  last_used_at: number;
};

export const SqliteImageCache = {
  async get(url: string): Promise<string | null> {
    const db = await getDb();
    const rows = await db.select<ImageRow[]>(
      "SELECT file_path FROM image_cache WHERE url = $1",
      [url],
    );
    if (!rows[0]) return null;
    if (!rows[0].file_path) {
      enqueueWriteSafely(() =>
        db.execute("DELETE FROM image_cache WHERE url = $1", [url]),
      );
      return null;
    }
    _pendingImageLastUsedUrls.add(url);
    scheduleImageLastUsedFlush();
    return rows[0].file_path;
  },

  async set(url: string, filePath: string): Promise<void> {
    const db = await getDb();
    const now = Date.now();
    const existingRows = await db.select<Array<{ file_path: string | null }>>(
      "SELECT file_path FROM image_cache WHERE url = $1",
      [url],
    );
    const pathsToDelete = new Set<string>();
    const previousPath = existingRows[0]?.file_path;
    if (previousPath && previousPath !== filePath) {
      pathsToDelete.add(previousPath);
    }

    await enqueueWrite(async () => {
      await db.execute(
        `INSERT INTO image_cache (url, data_url, file_path, updated_at, last_used_at)
         VALUES ($1, '', $2, $3, $3)
         ON CONFLICT(url) DO UPDATE SET data_url = '', file_path = $2, updated_at = $3, last_used_at = $3`,
        [url, filePath, now],
      );

      _imageCacheInsertCount++;
      if (_imageCacheInsertCount >= 50) {
        _imageCacheInsertCount = 0;
        const prunedRows = await db.select<Array<{ file_path: string | null }>>(
          `SELECT file_path FROM image_cache WHERE url NOT IN (
            SELECT url FROM image_cache ORDER BY last_used_at DESC LIMIT $1
          )`,
          [MAX_IMAGE_ENTRIES],
        );
        await db.execute(
          `DELETE FROM image_cache WHERE url NOT IN (
            SELECT url FROM image_cache ORDER BY last_used_at DESC LIMIT $1
          )`,
          [MAX_IMAGE_ENTRIES],
        );
        for (const row of prunedRows) {
          if (row.file_path) pathsToDelete.add(row.file_path);
        }
      }
    });

    if (pathsToDelete.size > 0) {
      await removeCachedImageFiles([...pathsToDelete]);
    }
  },
};

const MAX_CONVERSATIONS = 50;
const MAX_THREAD_ENTRIES = 150;
const THREAD_CACHE_TTL_MS = 2 * 60 * 1000;

function normalizeTenantKey(tenantId?: string | null): string {
  return tenantId ?? "__default__";
}

type ThreadCacheRow = {
  value: string;
  updated_at: number;
};

type ThreadCacheSnapshot = {
  data: ThreadQueryData;
  updatedAt: number;
};

export const SqliteWorkspaceShellStore = {
  async getSnapshot(): Promise<TeamsWorkspaceShellSnapshot | null> {
    const db = await getDb();

    const accountRows = await db.select<{ key: string; value: string }[]>(
      "SELECT value FROM workspace_shell WHERE key = 'accounts'",
    );
    const accounts: TeamsAccountOption[] = accountRows[0]
      ? (JSON.parse(accountRows[0].value) as TeamsAccountOption[])
      : [];

    const tenantRows = await db.select<
      { key: string; value: string; updated_at: number }[]
    >(
      "SELECT key, value, updated_at FROM workspace_shell WHERE key LIKE 'tenant:%'",
    );

    const tenants: Record<string, TeamsWorkspaceShellTenantSnapshot> = {};
    const sessionPrefix = "tenant:";
    const sessionSuffix = ":session";
    const convSuffix = ":conversations";

    const tenantData = new Map<
      string,
      {
        session?: TeamsSessionInfo;
        conversations?: Conversation[];
        updatedAt?: number;
      }
    >();

    for (const row of tenantRows) {
      if (row.key.endsWith(sessionSuffix)) {
        const tenantId = row.key
          .slice(sessionPrefix.length)
          .slice(0, -sessionSuffix.length);
        const existing = tenantData.get(tenantId) ?? {};
        existing.session = JSON.parse(row.value) as TeamsSessionInfo;
        existing.updatedAt = row.updated_at;
        tenantData.set(tenantId, existing);
      } else if (row.key.endsWith(convSuffix)) {
        const tenantId = row.key
          .slice(sessionPrefix.length)
          .slice(0, -convSuffix.length);
        const existing = tenantData.get(tenantId) ?? {};
        existing.conversations = JSON.parse(row.value) as Conversation[];
        tenantData.set(tenantId, existing);
      }
    }

    for (const [tenantId, data] of tenantData) {
      if (data.session) {
        tenants[tenantId] = {
          updatedAt: data.updatedAt ?? Date.now(),
          session: data.session,
          conversations: data.conversations ?? [],
        };
      }
    }

    if (accounts.length === 0 && Object.keys(tenants).length === 0) {
      return null;
    }
    return { accounts, tenants };
  },

  async updateAccounts(accounts: TeamsAccountOption[]): Promise<void> {
    const db = await getDb();
    await enqueueWrite(() =>
      db.execute(
        `INSERT INTO workspace_shell (key, value, updated_at)
         VALUES ('accounts', $1, $2)
         ON CONFLICT(key) DO UPDATE SET value = $1, updated_at = $2`,
        [JSON.stringify(accounts), Date.now()],
      ),
    );
  },

  async updateSession(
    tenantId: string | undefined,
    session: TeamsSessionInfo,
  ): Promise<void> {
    const db = await getDb();
    const key = `tenant:${normalizeTenantKey(tenantId)}:session`;
    await enqueueWrite(() =>
      db.execute(
        `INSERT INTO workspace_shell (key, value, updated_at)
         VALUES ($1, $2, $3)
         ON CONFLICT(key) DO UPDATE SET value = $2, updated_at = $3`,
        [key, JSON.stringify(session), Date.now()],
      ),
    );
  },

  async updateConversations(
    tenantId: string | undefined,
    conversations: Conversation[],
  ): Promise<void> {
    const db = await getDb();
    const key = `tenant:${normalizeTenantKey(tenantId)}:conversations`;
    const trimmed = conversations.slice(0, MAX_CONVERSATIONS);
    await enqueueWrite(() =>
      db.execute(
        `INSERT INTO workspace_shell (key, value, updated_at)
         VALUES ($1, $2, $3)
         ON CONFLICT(key) DO UPDATE SET value = $2, updated_at = $3`,
        [key, JSON.stringify(trimmed), Date.now()],
      ),
    );
  },
};

let _threadCacheStoreCount = 0;

export const SqliteThreadCache = {
  async getSnapshot(
    tenantId: string | undefined,
    conversationId: string,
  ): Promise<ThreadCacheSnapshot | null> {
    const db = await getDb();
    const normalizedTenantId = normalizeTenantKey(tenantId);
    const rows = await db.select<ThreadCacheRow[]>(
      `SELECT value, updated_at
       FROM thread_cache
       WHERE tenant_id = $1 AND conversation_id = $2`,
      [normalizedTenantId, conversationId],
    );
    const row = rows[0];
    if (!row) return null;
    enqueueWriteSafely(() =>
      db.execute(
        `UPDATE thread_cache
         SET last_requested_at = $1
         WHERE tenant_id = $2 AND conversation_id = $3`,
        [Date.now(), normalizedTenantId, conversationId],
      ),
    );
    return {
      data: JSON.parse(row.value) as ThreadQueryData,
      updatedAt: row.updated_at,
    };
  },

  async getFreshSnapshot(
    tenantId: string | undefined,
    conversationId: string,
    maxAgeMs = THREAD_CACHE_TTL_MS,
  ): Promise<ThreadCacheSnapshot | null> {
    const snapshot = await this.getSnapshot(tenantId, conversationId);
    if (!snapshot) return null;
    return Date.now() - snapshot.updatedAt <= maxAgeMs ? snapshot : null;
  },

  async storeThread(
    tenantId: string | undefined,
    conversationId: string,
    data: ThreadQueryData,
  ): Promise<void> {
    const db = await getDb();
    const now = Date.now();
    await enqueueWrite(async () => {
      await db.execute(
        `INSERT INTO thread_cache (tenant_id, conversation_id, value, updated_at, last_requested_at)
         VALUES ($1, $2, $3, $4, $4)
         ON CONFLICT(tenant_id, conversation_id) DO UPDATE SET
           value = $3,
           updated_at = $4,
           last_requested_at = $4`,
        [
          normalizeTenantKey(tenantId),
          conversationId,
          JSON.stringify(data),
          now,
        ],
      );

      _threadCacheStoreCount++;
      if (_threadCacheStoreCount >= 10) {
        _threadCacheStoreCount = 0;
        await db.execute(
          `DELETE FROM thread_cache
           WHERE (tenant_id, conversation_id) NOT IN (
             SELECT tenant_id, conversation_id
             FROM thread_cache
             ORDER BY last_requested_at DESC
             LIMIT $1
           )`,
          [MAX_THREAD_ENTRIES],
        );
      }
    });
  },
};

export const SqliteQueryPersister = {
  getStorage() {
    return {
      getItem: async (key: string): Promise<string | null> => {
        const db = await getDb();
        const rows = await db.select<{ value: string }[]>(
          "SELECT value FROM query_cache WHERE key = $1",
          [key],
        );
        return rows[0]?.value ?? null;
      },
      setItem: async (key: string, value: string): Promise<void> => {
        const db = await getDb();
        await enqueueWrite(() =>
          db.execute(
            `INSERT INTO query_cache (key, value, updated_at)
             VALUES ($1, $2, $3)
             ON CONFLICT(key) DO UPDATE SET value = $2, updated_at = $3`,
            [key, value, Date.now()],
          ),
        );
      },
      removeItem: async (key: string): Promise<void> => {
        const db = await getDb();
        await enqueueWrite(() =>
          db.execute("DELETE FROM query_cache WHERE key = $1", [key]),
        );
      },
    };
  },
};
