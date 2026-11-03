import { useEffect, useState, type ReactNode } from "react";
import { Badge, Button, Field, Input, Switch, Text } from "@fluentui/react-components";
import {
  checkLicense,
  getGlobalSettings,
  saveGlobalSettings,
  type GlobalSettings,
  type LicenseReport,
} from "../core/api";
import { runDiagnostics, type ProbeResult } from "../core/diagnostics";
import type { Settings } from "../core/state";
import { notify, settingsStore } from "./store";

function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section style={{ display: "grid", gap: 8 }}>
      <Text weight="semibold">{title}</Text>
      {children}
    </section>
  );
}

function WorkbookSettings() {
  const settings = settingsStore.use();
  const change = (patch: Partial<Settings>) => settingsStore.set({ ...settings, ...patch });
  const setRow = (i: number, patch: Partial<Settings["custom"][number]>) =>
    change({ custom: settings.custom.map((row, j) => (j === i ? { ...row, ...patch } : row)) });
  const numeric: [keyof Settings & ("timeLimit" | "mipGap" | "threads"), string, string][] = [
    ["timeLimit", "TimeLimit", "Seconds"],
    ["mipGap", "MIPGap", "For example 0.01"],
    ["threads", "Threads", "Blank lets Gurobi choose"],
  ];
  return (
    <Section title="Solver settings for this workbook">
      {numeric.map(([key, label, hint]) => (
        <Field key={key} label={label} hint={hint}>
          <Input
            size="small"
            type="number"
            min="0"
            value={settings[key]}
            onChange={(_, d) => change({ [key]: d.value })}
          />
        </Field>
      ))}
      <Switch
        label="OutputFlag (show Gurobi's log)"
        checked={settings.outputFlag}
        onChange={(_, d) => change({ outputFlag: d.checked })}
      />
      <Text size={200}>Other Gurobi parameters. These override the fields above.</Text>
      {settings.custom.map((row, i) => (
        <div key={i} style={{ display: "flex", gap: 4 }}>
          <Input
            size="small"
            aria-label={`Parameter ${i + 1} name`}
            placeholder="Name"
            value={row.key}
            onChange={(_, d) => setRow(i, { key: d.value })}
            style={{ flex: 1, minWidth: 0 }}
          />
          <Input
            size="small"
            aria-label={`Parameter ${i + 1} value`}
            placeholder="Value"
            value={row.value}
            onChange={(_, d) => setRow(i, { value: d.value })}
            style={{ flex: 1, minWidth: 0 }}
          />
          <Button
            size="small"
            aria-label={`Remove parameter ${i + 1}`}
            onClick={() => change({ custom: settings.custom.filter((_, j) => j !== i) })}
          >
            Remove
          </Button>
        </div>
      ))}
      <div>
        <Button
          size="small"
          onClick={() => change({ custom: [...settings.custom, { key: "", value: "" }] })}
        >
          Add parameter
        </Button>
      </div>
    </Section>
  );
}

function License() {
  const [busy, setBusy] = useState(false);
  const [report, setReport] = useState<LicenseReport>();
  const [error, setError] = useState<string>();
  const run = async () => {
    setBusy(true);
    setReport(undefined);
    setError(undefined);
    try {
      setReport(await checkLicense());
    } catch (e) {
      setError((e as Error).message);
    }
    setBusy(false);
  };
  const color = !report?.ok ? "danger" : report.restricted ? "warning" : "success";
  return (
    <Section title="License">
      <div>
        <Button size="small" disabled={busy} onClick={() => void run()}>
          {busy ? "Checking…" : "Check License"}
        </Button>
      </div>
      {error && <Text role="alert">{error}</Text>}
      {report && (
        <div data-testid="license-report" style={{ display: "grid", gap: 4 }}>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <Badge color={color}>{report.ok ? report.type : "Problem"}</Badge>
            {report.expiry && <Text size={200}>Expires {report.expiry}</Text>}
          </div>
          <Text size={200}>{report.message}</Text>
        </div>
      )}
    </Section>
  );
}

function HelperSettings() {
  const [values, setValues] = useState<GlobalSettings>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    getGlobalSettings().then(setValues, (e) => setError((e as Error).message));
  }, []);
  if (!values) return error ? <Text role="alert">{error}</Text> : null;
  const save = async () => {
    setError(undefined);
    try {
      setValues(await saveGlobalSettings(values));
      notify("Helper settings saved. They apply to the next run.", "success");
    } catch (e) {
      setError((e as Error).message);
    }
  };
  return (
    <Section title="Helper settings for this Mac">
      <Field label="Python path" hint={`Blank uses ${values.default_python}`}>
        <Input
          size="small"
          value={values.python_path}
          onChange={(_, d) => setValues({ ...values, python_path: d.value })}
        />
      </Field>
      <Field label="License file path" hint="Blank lets Gurobi find it, usually ~/gurobi.lic">
        <Input
          size="small"
          value={values.license_path}
          onChange={(_, d) => setValues({ ...values, license_path: d.value })}
        />
      </Field>
      <Field label="Keep runs" hint="How many run folders to keep">
        <Input
          size="small"
          type="number"
          min="1"
          value={String(values.keep_runs)}
          onChange={(_, d) => setValues({ ...values, keep_runs: Number(d.value) })}
        />
      </Field>
      <Text size={200}>
        Helper port: {values.port}. To change it, run GS_PORT=&lt;port&gt; scripts/install.sh,
        restart the helper, then restart Excel.
      </Text>
      {error && <Text role="alert">{error}</Text>}
      <div>
        <Button size="small" onClick={() => void save()}>
          Save helper settings
        </Button>
      </div>
    </Section>
  );
}

function Help() {
  return (
    <Section title="Help">
      <Text size={200}>
        Bind sheet cells to Python names on Data Items, write a gurobipy model on Model, then Solve
        Model. Assign to a data item to write its cells back; changed items show a * in the log.
      </Text>
      <Text size={200}>
        ⌘↵ solves and ⌘. stops. Models and items are saved inside the workbook, so save the workbook
        to keep them. The helper must be running: scripts/run-helper.sh.
      </Text>
    </Section>
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
    <Section title="Diagnostics">
      <Text size={200}>Checks what this version of Excel supports. Use a blank workbook.</Text>
      <div>
        <Button size="small" disabled={!inExcel || busy} onClick={run}>
          {busy ? "Running…" : "Run diagnostics"}
        </Button>
      </div>
      {!inExcel && <Text size={200}>Only available inside Excel.</Text>}
      {results.map((r) => (
        <div key={r.name} style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
          <Badge color={r.ok ? "success" : "danger"}>{r.ok ? "Pass" : "Fail"}</Badge>
          <Text size={200}>{r.name}</Text>
        </div>
      ))}
      {saved && <Text size={200}>Report saved for review.</Text>}
    </Section>
  );
}

export function SettingsTab({ inExcel }: { inExcel: boolean }) {
  return (
    <div style={{ display: "grid", gap: 20 }}>
      <WorkbookSettings />
      <License />
      <HelperSettings />
      <Help />
      <Diagnostics inExcel={inExcel} />
    </div>
  );
}
