import { useEffect, useRef } from "react";
import { Badge, Button, Text } from "@fluentui/react-components";
import { solve, stop } from "../core/solve";
import { excel } from "../excel";
import { useRun, type RunState } from "./store";

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
          <Text weight="semibold">
            {run.error.type}: {run.error.message}
            {run.error.line != null && ` (line ${run.error.line})`}
          </Text>
          {run.error.hint && <Text block>{run.error.hint}</Text>}
          {run.error.traceback && (
            <pre style={{ fontFamily: MONO, fontSize: 11, whiteSpace: "pre-wrap" }}>
              {run.error.traceback}
            </pre>
          )}
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
