import {
  Badge,
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Select,
  Text,
} from "@fluentui/react-components";
import { CodeEditor } from "./CodeEditor";
import { viewWorkingFiles } from "./OutputTab";
import { popOut } from "./popOut";
import { examplesOpenStore, itemsStore, jumpStore, MODEL_TEMPLATE, modelStore } from "./store";
import { useDark } from "./Themed";
import { useActiveSheet } from "./useActiveSheet";

export function ModelTab() {
  const sheet = useActiveSheet();
  const models = modelStore.use();
  const items = itemsStore.use();
  const jump = jumpStore.use();
  const dark = useDark();
  if (!sheet) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, height: "100%" }}>
      <div style={{ display: "flex", gap: 4, alignItems: "center", flexWrap: "wrap" }}>
        <Select size="small" disabled aria-label="Language" value="gurobi">
          <option value="gurobi">Gurobi</option>
        </Select>
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button size="small">File</Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem onClick={() => void viewWorkingFiles()}>View Working Files</MenuItem>
              <MenuItem onClick={() => examplesOpenStore.set(true)}>View Examples</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
        <Button size="small" onClick={() => void popOut(sheet)}>
          Pop out
        </Button>
      </div>
      <Text size={200}>Model for {sheet.name}</Text>
      {(items[sheet.id] ?? []).length > 0 && (
        <div
          data-testid="item-chips"
          aria-label="Data items"
          style={{ display: "flex", gap: 4, flexWrap: "wrap" }}
        >
          {items[sheet.id].map((item) => (
            <Badge key={item.name} appearance="tint" color="informative" title={item.type}>
              {item.name}
            </Badge>
          ))}
        </div>
      )}
      {/* One editor per sheet, so undo can't bring back another sheet's model. */}
      <CodeEditor
        key={sheet.id}
        value={models[sheet.id] ?? MODEL_TEMPLATE}
        dark={dark}
        onChange={(text) => modelStore.set({ ...modelStore.get(), [sheet.id]: text })}
        jump={jump}
        onJumped={() => jumpStore.set(null)}
      />
    </div>
  );
}
