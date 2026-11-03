import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Text,
} from "@fluentui/react-components";
import { EXAMPLES, insertExample, type Example } from "../core/examples";
import { excel } from "../excel";
import { examplesOpenStore, notify, setTab } from "./store";

export function ExamplesDialog() {
  const open = examplesOpenStore.use();
  const insert = async (example: Example) => {
    examplesOpenStore.set(false);
    setTab("model");
    try {
      const sheet = await insertExample(excel(), example);
      notify(`Added the ${sheet.name} sheet. Solve Model to run it.`, "success");
    } catch (e) {
      notify(`Could not add the example. (${(e as Error).message})`, "error");
    }
  };
  return (
    <Dialog open={open} onOpenChange={(_, d) => examplesOpenStore.set(d.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Examples</DialogTitle>
          <DialogContent style={{ display: "grid", gap: 12 }}>
            <Text block>Each one adds a new sheet with its data, data items and model.</Text>
            {EXAMPLES.map((example) => (
              <div key={example.name} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <Text block weight="semibold">
                    {example.name}
                  </Text>
                  <Text block size={200}>
                    {example.description}
                  </Text>
                </div>
                <Button
                  size="small"
                  aria-label={`Insert ${example.name}`}
                  onClick={() => void insert(example)}
                >
                  Insert
                </Button>
              </div>
            ))}
          </DialogContent>
          <DialogActions>
            <Button onClick={() => examplesOpenStore.set(false)}>Close</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
