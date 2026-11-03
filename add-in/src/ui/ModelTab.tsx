// ponytail: plain text box so M3's items can run a real model; M4 brings CodeMirror.
import { Text } from "@fluentui/react-components";
import { MODEL_TEMPLATE, modelStore } from "./store";
import { useActiveSheet } from "./useActiveSheet";

export function ModelTab() {
  const sheet = useActiveSheet();
  const models = modelStore.use();
  if (!sheet) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, height: "100%" }}>
      <Text size={200}>Model for {sheet.name}</Text>
      <textarea
        aria-label="Model code"
        spellCheck={false}
        value={models[sheet.id] ?? MODEL_TEMPLATE}
        onChange={(e) => modelStore.set({ ...models, [sheet.id]: e.target.value })}
        style={{
          flex: 1,
          minHeight: 240,
          resize: "none",
          padding: 8,
          fontFamily: '"SF Mono", Menlo, monospace',
          fontSize: 12,
          lineHeight: 1.4,
          color: "inherit",
          background: "transparent",
          border: "1px solid currentColor",
          borderRadius: 4,
        }}
      />
    </div>
  );
}
