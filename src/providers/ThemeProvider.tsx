import {
  createContext,
  type ReactNode,
  useCallback,
  useContext,
  useLayoutEffect,
  useMemo,
  useState,
} from "react";

const STORAGE_KEY = "better-teams-theme";

export type ThemePreference = "light" | "dark" | "system";

type ThemeContextValue = {
  theme: ThemePreference;
  setTheme: (value: ThemePreference) => void;
  resolved: "light" | "dark";
};

const ThemeContext = createContext<ThemeContextValue | null>(null);

function readStoredTheme(): ThemePreference {
  try {
    const s = localStorage.getItem(STORAGE_KEY);
    if (s === "dark" || s === "light" || s === "system") return s;
  } catch {
    /* ignore */
  }
  return "light";
}

function systemPrefersDark(): boolean {
  return window.matchMedia("(prefers-color-scheme: dark)").matches;
}

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [theme, setThemeState] = useState<ThemePreference>(readStoredTheme);
  const [resolved, setResolved] = useState<"light" | "dark">("light");

  const setTheme = useCallback((value: ThemePreference) => {
    try {
      localStorage.setItem(STORAGE_KEY, value);
    } catch {
      /* ignore */
    }
    setThemeState(value);
  }, []);

  useLayoutEffect(() => {
    const root = document.documentElement;
    const apply = () => {
      const dark =
        theme === "dark" || (theme === "system" && systemPrefersDark());
      root.classList.toggle("dark", dark);
      setResolved(dark ? "dark" : "light");
    };
    apply();
    if (theme !== "system") return;
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    mq.addEventListener("change", apply);
    return () => mq.removeEventListener("change", apply);
  }, [theme]);

  const value = useMemo(
    () => ({ theme, setTheme, resolved }),
    [theme, setTheme, resolved],
  );

  return (
    <ThemeContext.Provider value={value}>{children}</ThemeContext.Provider>
  );
}

export function useTheme(): ThemeContextValue {
  const ctx = useContext(ThemeContext);
  if (!ctx) {
    throw new Error("useTheme must be used within ThemeProvider");
  }
  return ctx;
}
