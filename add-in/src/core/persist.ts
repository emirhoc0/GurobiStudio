// Keeps the stores and the workbook in step. Nothing here ever starts a solve.
import type { ExcelAdapter } from "../excel/adapter";
import {
  itemsStore,
  modelStore,
  notify,
  saveBlockedStore,
  settingsStore,
  workbookIdStore,
} from "../ui/store";
import { parseState, remapSheets, SCHEMA, type SheetState, type WorkbookState } from "./state";

const DEBOUNCE_MS = 300;
const RETRY_MS = 2000;

let markLoaded = () => {};
const loaded = new Promise<void>((resolve) => {
  markLoaded = resolve;
});
/** A ribbon Solve can arrive before the workbook's state has been read. */
export const whenLoaded = () => loaded;

let lastNames: Record<string, string> = {};

async function snapshot(excel: ExcelAdapter): Promise<WorkbookState> {
  for (const s of await excel.listSheets()) lastNames[s.id] = s.name;
  const models = modelStore.get();
  const items = itemsStore.get();
  const ids = new Set([...Object.keys(models), ...Object.keys(items)]);
  const sheets: Record<string, SheetState> = {};
  for (const id of ids) {
    sheets[id] = {
      name: lastNames[id] ?? "",
      language: "gurobi",
      ...(id in models ? { model: models[id] } : {}),
      items: items[id] ?? [],
    };
  }
  return {
    schema: SCHEMA,
    workbookId: workbookIdStore.get(),
    settings: settingsStore.get(),
    sheets,
  };
}

function startSaving(excel: ExcelAdapter) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let saving = false;
  let dirty = false;
  let warned = false;

  const save = async () => {
    if (saving) {
      dirty = true; // writes never overlap, so an older one can't land last
      return;
    }
    saving = true;
    dirty = false;
    try {
      await excel.saveState(JSON.stringify(await snapshot(excel)));
      warned = false;
    } catch (e) {
      if (!warned) notify(`Could not save to the workbook. (${(e as Error).message})`, "error");
      warned = true;
      timer = setTimeout(() => void save(), RETRY_MS);
    }
    saving = false;
    if (dirty) void save();
  };
  const schedule = () => {
    clearTimeout(timer);
    timer = setTimeout(() => void save(), DEBOUNCE_MS);
  };
  [modelStore, itemsStore, settingsStore].forEach((store) => store.subscribe(schedule));
}

/** Fills the stores from the workbook. If its state can't be read, saving stays off. */
export async function loadWorkbook(excel: ExcelAdapter) {
  try {
    const text = await excel.loadState();
    if (text === null) {
      workbookIdStore.set(crypto.randomUUID());
    } else {
      const state = parseState(text);
      const sheets = remapSheets(state.sheets, await excel.listSheets());
      const models: Record<string, string> = {};
      const items: Record<string, SheetState["items"]> = {};
      lastNames = {};
      for (const [id, sheet] of Object.entries(sheets)) {
        lastNames[id] = sheet.name;
        if (sheet.model !== undefined) models[id] = sheet.model;
        items[id] = sheet.items ?? [];
      }
      workbookIdStore.set(state.workbookId);
      settingsStore.set(state.settings);
      modelStore.set(models);
      itemsStore.set(items);
    }
    startSaving(excel);
  } catch (e) {
    workbookIdStore.set(crypto.randomUUID());
    saveBlockedStore.set(
      `This workbook's GurobiStudio data could not be read, so changes won't be saved and the data is untouched. (${(e as Error).message})`,
    );
  }
  markLoaded();
}
