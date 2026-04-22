export async function openExternal(url: string): Promise<void> {
  try {
    if (!window.betterTeams) throw new Error("Desktop API is not available");
    await window.betterTeams.shell.openExternal(url);
  } catch {
    window.open(url, "_blank", "noopener,noreferrer");
  }
}
