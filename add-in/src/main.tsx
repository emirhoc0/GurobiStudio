import { createRoot } from "react-dom/client";
import { registerCommands } from "./commands";
import { runDialogSide } from "./core/diagnostics";
import { setExcel } from "./excel";
import { createMockAdapter } from "./excel/mockAdapter";
import { officeAdapter } from "./excel/officeAdapter";
import { App } from "./ui/App";

function start(inExcel: boolean) {
  if (inExcel) {
    setExcel(officeAdapter);
  } else {
    const mock = createMockAdapter();
    setExcel(mock.adapter);
    Object.assign(window, { mockCells: mock.cells }); // lets browser tests inspect results
  }
  createRoot(document.getElementById("root")!).render(<App inExcel={inExcel} />);
}

const view = new URLSearchParams(location.search).get("view");

if (typeof Office === "undefined") {
  start(false); // office.js could not load; still usable in a browser with the mock adapter
} else {
  Office.onReady((info) => {
    if (view === "diag-dialog") {
      document.getElementById("root")!.textContent = "Testing dialog messaging…";
      runDialogSide();
      return;
    }
    if (info.host) registerCommands();
    start(Boolean(info.host));
  });
}
