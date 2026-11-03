import { useSyncExternalStore } from "react";
import type { DataItem } from "../core/shapes";
import { DEFAULT_SETTINGS, type Settings } from "../core/state";

export type TabName = "model" | "data" | "output" | "settings";

export interface RunError {
  type: string;
  message: string;
  line: number | null;
  traceback?: string;
  hint?: string | null;
}

export interface RunState {
  phase: "idle" | "running" | "finished" | "stopped" | "failed";
  runId?: string;
  log: string[];
  status?: string;
  objective?: number | null;
  elapsed?: number;
  changed?: string[];
  warnings?: string[];
  error?: RunError;
}

function createStore<T>(initial: T) {
  let value = initial;
  const listeners = new Set<() => void>();
  const subscribe = (listener: () => void) => {
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  };
  return {
    get: () => value,
    set: (next: T) => {
      value = next;
      listeners.forEach((l) => l());
    },
    subscribe,
    use: () => useSyncExternalStore(subscribe, () => value),
  };
}

const tabStore = createStore<TabName>("model");
export const setTab = tabStore.set;
export const useTab = tabStore.use;

export const runStore = createStore<RunState>({ phase: "idle", log: [] });
export const useRun = runStore.use;

export const jumpStore = createStore<number | null>(null); // a model line the editor should select

export interface ConfirmRequest {
  title: string;
  body: string;
  code?: string;
  confirm: string;
  cancel: string;
  resolve: (answer: boolean) => void;
}
export const confirmStore = createStore<ConfirmRequest | null>(null);
export const askConfirm = (request: Omit<ConfirmRequest, "resolve">) =>
  new Promise<boolean>((resolve) => confirmStore.set({ ...request, resolve }));

export interface Notice {
  text: string;
  intent: "success" | "error" | "info";
}
export const noticeStore = createStore<Notice | null>(null);
export const notify = (text: string, intent: Notice["intent"] = "info") =>
  noticeStore.set({ text, intent });

// Saved in the workbook by core/persist.ts. The first two are keyed by worksheet id.
export const itemsStore = createStore<Record<string, DataItem[]>>({});
export const modelStore = createStore<Record<string, string>>({});

export const settingsStore = createStore<Settings>(DEFAULT_SETTINGS);
export const workbookIdStore = createStore("");
export const saveBlockedStore = createStore<string | null>(null);

export const MODEL_TEMPLATE = `from SolverStudio import *
from gurobipy import *

`;
