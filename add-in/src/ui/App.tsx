import { useEffect, useState } from "react";
import { Badge, Tab, TabList, Text } from "@fluentui/react-components";
import { getHealth, type Health } from "../core/api";
import { solve, stop } from "../core/solve";
import { excel } from "../excel";
import { ConfirmDialog, Notices } from "./ConfirmDialog";
import { DataItemsTab } from "./DataItemsTab";
import { ExamplesDialog } from "./ExamplesDialog";
import { ModelTab } from "./ModelTab";
import { OutputTab } from "./OutputTab";
import { SettingsTab } from "./SettingsTab";
import { saveBlockedStore, setTab, useTab, type TabName } from "./store";
import { Themed, useRunShortcuts } from "./Themed";

function HelperStatus() {
  const [health, setHealth] = useState<Health | null>(null);
  const [error, setError] = useState(false);
  useEffect(() => {
    getHealth().then(setHealth, () => setError(true));
  }, []);
  if (error) return <Badge color="danger">Helper unreachable</Badge>;
  if (!health) return <Badge appearance="outline">Checking helper</Badge>;
  return (
    <Badge color="success" title={`Helper ${health.version}, Python ${health.python}`}>
      Gurobi {health.gurobipy ?? "not installed"}
    </Badge>
  );
}

const TABS: [TabName, string][] = [
  ["model", "Model"],
  ["data", "Data Items"],
  ["output", "Output"],
  ["settings", "Settings"],
];

export function App({ inExcel }: { inExcel: boolean }) {
  const tab = useTab();
  const saveBlocked = saveBlockedStore.use();
  useRunShortcuts(
    () => void solve(excel()),
    () => void stop(),
  );
  return (
    <Themed>
      <div style={{ padding: "8px 12px 0", display: "flex", justifyContent: "flex-start" }}>
        <HelperStatus />
      </div>
      {saveBlocked && (
        <Text role="alert" size={200} style={{ padding: "4px 12px", color: "#c50f1f" }}>
          {saveBlocked}
        </Text>
      )}
      <TabList
        size="small"
        selectedValue={tab}
        onTabSelect={(_, d) => setTab(d.value as TabName)}
        style={{ padding: "0 4px" }}
      >
        {TABS.map(([value, label]) => (
          <Tab key={value} value={value}>
            {label}
          </Tab>
        ))}
      </TabList>
      <div style={{ padding: 12, overflow: "auto", flex: 1 }}>
        {/* These stay mounted: a half-filled form and the editor's undo history survive a tab switch. */}
        <div hidden={tab !== "data"}>
          <DataItemsTab />
        </div>
        <div style={{ height: "100%", display: tab === "model" ? "block" : "none" }}>
          <ModelTab />
        </div>
        {tab === "settings" && <SettingsTab inExcel={inExcel} />}
        {tab === "output" && <OutputTab />}
      </div>
      <ConfirmDialog />
      <ExamplesDialog />
      <Notices />
    </Themed>
  );
}
