import { setTab, type TabName } from "./ui/store";

async function showPane(tab: TabName, event: Office.AddinCommands.Event) {
  try {
    setTab(tab);
    await Office.addin.showAsTaskpane();
  } finally {
    event.completed();
  }
}

export function registerCommands() {
  const open = (tab: TabName) => (event: Office.AddinCommands.Event) => showPane(tab, event);
  Office.actions.associate("editData", open("data"));
  Office.actions.associate("showModel", open("model"));
  Office.actions.associate("solveModel", open("output"));
  Office.actions.associate("stopSolve", (event: Office.AddinCommands.Event) => event.completed());
  Office.actions.associate("openSettings", open("settings"));
  Office.actions.associate("openExamples", open("model"));
  Office.actions.associate("openWorkingFiles", open("output"));
  Office.actions.associate("openHelp", open("settings"));
}
