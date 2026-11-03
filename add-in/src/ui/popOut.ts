// The pane's side of the pop-out editor. The dialog is a separate page with no access to our
// stores, so the model text travels as messages, in both directions.
import { solve, stop } from "../core/solve";
import { excel } from "../excel";
import type { EditorDialog } from "../excel/adapter";
import { MODEL_TEMPLATE, modelStore, notify } from "./store";
import type { Sheet } from "./useActiveSheet";

export type ToPane = { type: "ready" | "solve" | "stop" } | { type: "text"; text: string };
export type ToDialog = { type: "text"; text: string; sheet: string };

let open: EditorDialog | null = null;

export async function popOut(sheet: Sheet) {
  if (open) return;
  const textNow = () => modelStore.get()[sheet.id] ?? MODEL_TEMPLATE;
  let shown = "";
  const push = () => {
    if (textNow() === shown) return;
    shown = textNow();
    const message: ToDialog = { type: "text", text: shown, sheet: sheet.name };
    open?.send(JSON.stringify(message));
  };
  let unsubscribe = () => {};
  try {
    open = await excel().openEditorDialog(
      (raw) => {
        const message = JSON.parse(raw) as ToPane;
        if (message.type === "ready") {
          shown = "";
          push();
        } else if (message.type === "text") {
          shown = message.text;
          modelStore.set({ ...modelStore.get(), [sheet.id]: message.text });
        } else if (message.type === "solve") {
          void solve(excel());
        } else {
          void stop();
        }
      },
      () => {
        open = null;
        unsubscribe();
      },
    );
    unsubscribe = modelStore.subscribe(push);
  } catch (e) {
    notify(`Could not open the editor window. (${(e as Error).message})`, "error");
  }
}
