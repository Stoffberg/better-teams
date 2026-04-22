type BindValue = string | number | boolean | null;

export default class ElectronDatabase {
  static async load(_path: string): Promise<ElectronDatabase> {
    return new ElectronDatabase();
  }

  async execute(sql: string, bindValues: BindValue[] = []): Promise<void> {
    await api().sqlite.execute(sql, bindValues);
  }

  async select<T>(sql: string, bindValues: BindValue[] = []): Promise<T> {
    return api().sqlite.select<T>(sql, bindValues);
  }
}

function api(): BetterTeamsDesktopApi {
  if (!window.betterTeams) {
    throw new Error("Better Teams desktop API is not available");
  }
  return window.betterTeams;
}
