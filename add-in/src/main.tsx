import { createRoot } from "react-dom/client";
import { registerCommands } from "./commands";
import { runDialogSide } from "./core/diagnostics";
import { App } from "./ui/App";

function start(inExcel: boolean) {
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
