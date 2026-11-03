import { solve, stop } from "./core/solve";
import { excel } from "./excel";
import { viewWorkingFiles } from "./ui/OutputTab";
import { setTab, type TabName } from "./ui/store";

type Event = Office.AddinCommands.Event;

async function showPane(tab: TabName, event: Event) {
  try {
    setTab(tab);
    await Office.addin.showAsTaskpane();
  } finally {
    event.completed();
  }
}

export function registerCommands() {
  const open = (tab: TabName) => (event: Event) => showPane(tab, event);
  Office.actions.associate("editData", open("data"));
  Office.actions.associate("showModel", open("model"));
  Office.actions.associate("solveModel", async (event: Event) => {
    await showPane("output", event); // completes first, since a solve can run for hours
    void solve(excel());
  });
  Office.actions.associate("stopSolve", async (event: Event) => {
    try {
      await stop();
    } finally {
      event.completed();
    }
  });
  Office.actions.associate("openSettings", open("settings"));
  Office.actions.associate("openExamples", open("model"));
  Office.actions.associate("openWorkingFiles", async (event: Event) => {
    try {
      await viewWorkingFiles();
    } finally {
      event.completed();
    }
  });
  Office.actions.associate("openHelp", open("settings"));
}
