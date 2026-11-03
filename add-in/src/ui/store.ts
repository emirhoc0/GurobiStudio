import { useSyncExternalStore } from "react";

export type TabName = "model" | "data" | "output" | "settings";

let tab: TabName = "model";
const listeners = new Set<() => void>();

export function setTab(next: TabName) {
  tab = next;
  listeners.forEach((l) => l());
}

export function useTab(): TabName {
  return useSyncExternalStore(
    (l) => {
      listeners.add(l);
      return () => listeners.delete(l);
    },
    () => tab,
  );
}
