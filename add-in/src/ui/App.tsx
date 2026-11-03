import { useEffect, useState } from "react";
import { Badge, Button, Tab, TabList, Text } from "@fluentui/react-components";
import { getHealth, type Health } from "../core/api";
import { runDiagnostics, type ProbeResult } from "../core/diagnostics";
import { solve, stop } from "../core/solve";
import { excel } from "../excel";
import { ConfirmDialog, Notices } from "./ConfirmDialog";
import { DataItemsTab } from "./DataItemsTab";
import { ModelTab } from "./ModelTab";
import { OutputTab } from "./OutputTab";
import { setTab, useTab, type TabName } from "./store";
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

function Diagnostics({ inExcel }: { inExcel: boolean }) {
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState<ProbeResult[]>([]);
  const [saved, setSaved] = useState<string>();
  const run = async () => {
    setBusy(true);
    const report = await runDiagnostics();
    setResults(report.results);
    setSaved(report.saved);
    setBusy(false);
  };
  return (
    <div style={{ display: "grid", gap: 8 }}>
      <Text>Checks what this version of Excel supports. Use a blank workbook.</Text>
      <Button appearance="primary" disabled={!inExcel || busy} onClick={run}>
        {busy ? "Running…" : "Run diagnostics"}
      </Button>
      {!inExcel && <Text size={200}>Only available inside Excel.</Text>}
      {results.map((r) => (
        <div key={r.name} style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
          <Badge color={r.ok ? "success" : "danger"}>{r.ok ? "Pass" : "Fail"}</Badge>
          <Text size={200}>{r.name}</Text>
        </div>
      ))}
      {saved && <Text size={200}>Report saved for review.</Text>}
    </div>
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
  useRunShortcuts(
    () => void solve(excel()),
    () => void stop(),
  );
  return (
    <Themed>
      <div style={{ padding: "8px 12px 0", display: "flex", justifyContent: "flex-start" }}>
        <HelperStatus />
      </div>
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
        {tab === "settings" && <Diagnostics inExcel={inExcel} />}
        {tab === "output" && <OutputTab />}
      </div>
      <ConfirmDialog />
      <Notices />
    </Themed>
  );
}
