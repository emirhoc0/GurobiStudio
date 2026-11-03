import { useEffect, useState, type ReactNode } from "react";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { excel } from "../excel";

const FONT = '-apple-system, BlinkMacSystemFont, "SF Pro Text", "Helvetica Neue", sans-serif';

export function useDark() {
  const read = () => excel().officeDark() ?? matchMedia("(prefers-color-scheme: dark)").matches;
  const [dark, setDark] = useState(read);
  useEffect(() => {
    const query = matchMedia("(prefers-color-scheme: dark)");
    const update = () => setDark(read());
    query.addEventListener("change", update);
    addEventListener("focus", update); // Excel has no theme-changed event on Mac
    return () => {
      query.removeEventListener("change", update);
      removeEventListener("focus", update);
    };
  }, []);
  return dark;
}

export function Themed({ children }: { children: ReactNode }) {
  const base = useDark() ? webDarkTheme : webLightTheme;
  return (
    <FluentProvider
      theme={{ ...base, fontFamilyBase: FONT }}
      style={{ height: "100%", display: "flex", flexDirection: "column" }}
    >
      {children}
    </FluentProvider>
  );
}

/** ⌘↵ and ⌘. Capture phase, because the editor would otherwise take ⌘↵ for itself. */
export function useRunShortcuts(onSolve: () => void, onStop: () => void) {
  useEffect(() => {
    const listener = (e: KeyboardEvent) => {
      if (!e.metaKey || (e.key !== "Enter" && e.key !== ".")) return;
      e.preventDefault();
      e.stopPropagation();
      if (e.key === "Enter") onSolve();
      else onStop();
    };
    addEventListener("keydown", listener, true);
    return () => removeEventListener("keydown", listener, true);
  }, []);
}
