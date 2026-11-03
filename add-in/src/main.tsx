import { createRoot } from "react-dom/client";
import { registerCommands } from "./commands";
import { runDialogSide } from "./core/diagnostics";
import { setExcel } from "./excel";
import { createMockAdapter } from "./excel/mockAdapter";
import { officeAdapter } from "./excel/officeAdapter";
import { App } from "./ui/App";
import { EditorWindow } from "./ui/EditorWindow";
import { itemsStore, modelStore } from "./ui/store";

const view = new URLSearchParams(location.search).get("view");

function start(inExcel: boolean) {
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
    });
  }
  createRoot(document.getElementById("root")!).render(<App inExcel={inExcel} />);
}

if (typeof Office === "undefined") {
  start(false); // office.js could not load; still usable in a browser with the mock adapter
} else {
  Office.onReady((info) => {
    if (view === "diag-dialog") {
      document.getElementById("root")!.textContent = "Testing dialog messaging…";
      runDialogSide();
      return;
    }
    if (info.host && view !== "editor") registerCommands();
    start(Boolean(info.host));
  });
}
