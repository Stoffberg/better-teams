import fs from "node:fs";
import path from "node:path";
import Database from "better-sqlite3";
import { app } from "electron";

type BindValue = string | number | boolean | null;

let db: Database.Database | null = null;

export function execute(sql: string, bindValues: BindValue[] = []): void {
  const query = normalizeSql(sql, bindValues);
  getDb()
    .prepare(query.sql)
    .run(...query.bindValues);
}

export function select<T>(sql: string, bindValues: BindValue[] = []): T {
  const query = normalizeSql(sql, bindValues);
  return getDb()
    .prepare(query.sql)
    .all(...query.bindValues) as T;
}

function getDb(): Database.Database {
  if (db) return db;
  const dbPath = cacheDbPath();
  fs.mkdirSync(path.dirname(dbPath), { recursive: true });
  db = new Database(dbPath);
  return db;
}

function cacheDbPath(): string {
  const base =
    process.platform === "darwin"
      ? path.join(app.getPath("appData"), "com.betterteams.app")
      : app.getPath("userData");
  return path.join(base, "better-teams-cache.db");
}

function normalizeSql(
  sql: string,
  bindValues: BindValue[],
): { sql: string; bindValues: BindValue[] } {
  const normalizedValues: BindValue[] = [];
  const normalizedSql = sql.replace(/\$(\d+)/g, (_match, indexText: string) => {
    const index = Number(indexText) - 1;
    normalizedValues.push(bindValues[index] ?? null);
    return "?";
  });
  return {
    sql: normalizedSql,
    bindValues: normalizedValues.length > 0 ? normalizedValues : bindValues,
  };
}
