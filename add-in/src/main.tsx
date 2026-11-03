import { createRoot } from "react-dom/client";
import { registerCommands } from "./commands";
import { runDialogSide } from "./core/diagnostics";
import { loadWorkbook } from "./core/persist";
import { excel, setExcel } from "./excel";
import { createMockAdapter } from "./excel/mockAdapter";
import { officeAdapter } from "./excel/officeAdapter";
import { App } from "./ui/App";
import { EditorWindow } from "./ui/EditorWindow";
import { itemsStore, modelStore, settingsStore } from "./ui/store";

const view = new URLSearchParams(location.search).get("view");

// At the top level, as Microsoft's samples do, so a click that starts a fresh runtime finds its action.
if (typeof Office !== "undefined" && !view) registerCommands();

async function start(inExcel: boolean) {
  if (view === "editor") {
    setExcel(inExcel ? officeAdapter : createMockAdapter().adapter);
    createRoot(document.getElementById("root")!).render(<EditorWindow />);
    return;
  }
  if (inExcel) {
    setExcel(officeAdapter);
  } else {
    const mock = createMockAdapter();
    setExcel(mock.adapter);
    // For browser tests.
    Object.assign(window, {
      mockCells: mock.cells,
      mockExcel: mock,
      gsModels: modelStore,
      gsItems: itemsStore,
      gsSettings: settingsStore,
    });
  }
  const root = document.getElementById("root")!;
  root.textContent = "Loading…";
  await loadWorkbook(excel());
  root.textContent = "";
  createRoot(root).render(<App inExcel={inExcel} />);
}

if (typeof Office === "undefined") {
  void start(false); // office.js could not load; still usable in a browser with the mock adapter
} else {
  Office.onReady((info) => {
    if (view === "diag-dialog") {
      document.getElementById("root")!.textContent = "Testing dialog messaging…";
      runDialogSide();
      return;
    }
    void start(Boolean(info.host));
  });
}
