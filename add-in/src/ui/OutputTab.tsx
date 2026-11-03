import { useEffect, useRef } from "react";
import { Badge, Button, Link, Text } from "@fluentui/react-components";
import { revealLatestRun } from "../core/api";
import { solve, stop } from "../core/solve";
import { excel } from "../excel";
import { jumpStore, notify, runStore, setTab, useRun, type RunState } from "./store";

const MONO = '"SF Mono", Menlo, monospace';

function chip(run: RunState) {
  if (run.phase === "running") return <Badge color="informative">Running</Badge>;
  if (run.phase === "stopped") return <Badge color="warning">Stopped</Badge>;
  if (run.phase === "failed") return <Badge color="danger">Error</Badge>;
  if (run.phase !== "finished") return null;
  const status = run.status ?? "Done";
  const color = status === "Optimal" ? "success" : status === "Done" ? "informative" : "warning";
  return <Badge color={color}>{status}</Badge>;
}

function jumpTo(line: number) {
  jumpStore.set(line);
  setTab("model");
}

const LINE_REF = /(File "model\.py", line \d+)/;

function Traceback({ text }: { text: string }) {
  return (
    <pre style={{ fontFamily: MONO, fontSize: 11, whiteSpace: "pre-wrap" }}>
      {text.split(LINE_REF).map((part, i) =>
        LINE_REF.test(part) ? (
          <Link key={i} as="button" onClick={() => jumpTo(Number(part.match(/\d+$/)![0]))}>
            {part}
          </Link>
        ) : (
          part
        ),
      )}
    </pre>
  );
}

async function copyText(text: string) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // Excel's web view can refuse the clipboard API.
    const area = document.createElement("textarea");
    area.value = text;
    document.body.append(area);
    area.select();
    const copied = document.execCommand("copy");
    area.remove();
    if (!copied) throw new Error("The clipboard is not available here.");
  }
}

export async function viewWorkingFiles() {
  try {
    if (!(await revealLatestRun())) notify("No runs yet.");
  } catch (e) {
    notify(`Could not open the run folder. (${(e as Error).message})`, "error");
  }
}

export function OutputTab() {
  const run = useRun();
  const logEnd = useRef<HTMLDivElement>(null);
  useEffect(() => {
    logEnd.current?.scrollIntoView({ block: "nearest" });
  }, [run.log.length]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, height: "100%" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <Button
          appearance="primary"
          size="small"
          disabled={run.phase === "running"}
          onClick={() => void solve(excel())}
        >
          Solve Model
        </Button>
        <Button size="small" disabled={run.phase !== "running"} onClick={() => void stop()}>
          Stop
        </Button>
        {chip(run)}
      </div>
      <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
        <Button
          size="small"
          appearance="subtle"
          disabled={!run.log.length}
          onClick={() =>
            copyText(run.log.join("\n")).then(
              () => notify("Log copied.", "success"),
              (e) => notify((e as Error).message, "error"),
            )
          }
        >
          Copy log
        </Button>
        <Button
          size="small"
          appearance="subtle"
          disabled={run.phase === "running" || run.phase === "idle"}
          onClick={() => runStore.set({ phase: "idle", log: [] })}
        >
          Clear
        </Button>
        <Button size="small" appearance="subtle" onClick={() => void viewWorkingFiles()}>
          View Working Files
        </Button>
      </div>

      {run.phase === "finished" && (
        <Text size={200}>
          {run.objective != null && `Objective ${run.objective}. `}
          {run.elapsed != null && `${run.elapsed}s. `}
          {run.changed?.length
            ? `Changed: ${run.changed.map((n) => `${n}*`).join(", ")}`
            : "No items changed."}
        </Text>
      )}
      {run.warnings?.map((w) => (
        <Text key={w} size={200}>
          Warning: {w}
        </Text>
      ))}
      {run.error && (
        <div role="alert">
          <Text weight="semibold" style={{ whiteSpace: "pre-wrap" }}>
            {run.error.type}: {run.error.message}
          </Text>
          {run.error.line != null && (
            <Text block>
              <Link as="button" onClick={() => jumpTo(run.error!.line!)}>
                Go to line {run.error.line}
              </Link>
            </Text>
          )}
          {run.error.hint && <Text block>{run.error.hint}</Text>}
          {run.error.traceback && <Traceback text={run.error.traceback} />}
        </div>
      )}

      {run.phase === "idle" ? (
        <Text>Click Solve Model to run. The log appears here.</Text>
      ) : (
        <pre
          data-testid="log"
          style={{
            flex: 1,
            margin: 0,
            overflow: "auto",
            fontFamily: MONO,
            fontSize: 11,
            lineHeight: 1.4,
          }}
        >
          {run.log.join("\n")}
          <div ref={logEnd} />
        </pre>
      )}
    </div>
  );
}
