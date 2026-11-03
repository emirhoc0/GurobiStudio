import { useSyncExternalStore } from "react";
import type { DataItem } from "../core/shapes";

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
    use: () => useSyncExternalStore(subscribe, () => value),
  };
}

const tabStore = createStore<TabName>("model");
export const setTab = tabStore.set;
export const useTab = tabStore.use;

export const runStore = createStore<RunState>({ phase: "idle", log: [] });
export const useRun = runStore.use;

// ponytail: in memory until M5 saves these in the workbook. Both are keyed by worksheet id.
export const itemsStore = createStore<Record<string, DataItem[]>>({});
export const modelStore = createStore<Record<string, string>>({});

export const MODEL_TEMPLATE = `from SolverStudio import *
from gurobipy import *

`;
