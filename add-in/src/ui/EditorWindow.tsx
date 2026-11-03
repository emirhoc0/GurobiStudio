import { useEffect, useRef, useState } from "react";
import { Text } from "@fluentui/react-components";
import { excel } from "../excel";
import { CodeEditor } from "./CodeEditor";
import type { ToDialog, ToPane } from "./popOut";
import { Themed, useDark, useRunShortcuts } from "./Themed";

function Editor() {
  const dark = useDark();
  const [state, setState] = useState<{ text: string; sheet: string } | null>(null);
  const send = useRef<(message: ToPane) => void>(() => {});

  useEffect(() => {
    const channel = excel().editorChannel((raw) => {
      const { text, sheet } = JSON.parse(raw) as ToDialog;
      setState({ text, sheet });
    });
    send.current = (message) => channel(JSON.stringify(message));
    send.current({ type: "ready" });
  }, []);
  useRunShortcuts(
    () => send.current({ type: "solve" }),
    () => send.current({ type: "stop" }),
  );

  if (!state) return <Text>Loading the model…</Text>;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, height: "100%", padding: 12 }}>
      <Text size={200}>Model for {state.sheet}. Results and the log appear in the task pane.</Text>
      <CodeEditor
        value={state.text}
        dark={dark}
        onChange={(text) => {
          setState((s) => s && { ...s, text });
          send.current({ type: "text", text });
        }}
      />
    </div>
  );
}

export function EditorWindow() {
  return (
    <Themed>
      <Editor />
    </Themed>
  );
}
