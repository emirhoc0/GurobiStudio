import { useEffect, useMemo, useState } from "react";
import { Badge, Button, Checkbox, Field, Input, Select, Text } from "@fluentui/react-components";
import { getHealth } from "../core/api";
import { localize, parseRef } from "../core/refs";
import {
  entries,
  overlapWarnings,
  preview,
  readItems,
  toItemIndexes,
  type BuiltItem,
  type DataItem,
  type ItemResult,
  type ItemType,
} from "../core/shapes";
import { validateItem } from "../core/validate";
import { excel } from "../excel";
import { itemsStore, useTab } from "./store";
import { useActiveSheet } from "./useActiveSheet";

const MONO = '"SF Mono", Menlo, monospace';
const NAMES_LIST = "gs-range-names";
const TYPES: [ItemType, string][] = [
  ["single", "Single value"],
  ["list", "List"],
  ["indexed", "Indexed"],
  ["table", "Table"],
];

interface Draft {
  name: string;
  type: ItemType;
  cell: string;
  index: [string, string];
  readOnly: boolean;
}

const EMPTY: Draft = { name: "", type: "list", cell: "", index: ["", ""], readOnly: false };

const toItem = (d: Draft): DataItem => ({
  name: d.name.trim(),
  type: d.type,
  cellRange: d.cell.trim(),
  indexRanges: toItemIndexes(d.type, d.index),
  readOnly: d.readOnly,
});

const toDraft = (i: DataItem): Draft => ({
  name: i.name,
  type: i.type,
  cell: i.cellRange,
  index: i.type === "table" ? [i.indexRanges[0], i.indexRanges[1]] : [i.indexRanges.join(", "), ""],
  readOnly: i.readOnly,
});

const isBuilt = (r: ItemResult): r is BuiltItem => !("error" in r);
const isA1 = (text: string) => {
  try {
    return parseRef(text).kind === "a1";
  } catch {
    return false;
  }
};

function RangeField(props: {
  label: string;
  hint: string;
  value: string;
  error?: string;
  showError: boolean;
  onChange: (value: string) => void;
  onPick: () => void;
}) {
  return (
    <Field
      label={props.label}
      hint={props.hint}
      validationMessage={props.showError ? props.error : undefined}
    >
      {(fieldProps) => (
        <div style={{ display: "flex", gap: 4 }}>
          <Input
            {...fieldProps}
            size="small"
            list={NAMES_LIST}
            value={props.value}
            onChange={(_, d) => props.onChange(d.value)}
            style={{ flex: 1, minWidth: 0, fontFamily: MONO }}
          />
          <Button
            size="small"
            onClick={props.onPick}
            aria-label={`Pick ${props.label} from selection`}
          >
            Pick from selection
          </Button>
        </div>
      )}
    </Field>
  );
}

export function DataItemsTab() {
  const tab = useTab();
  const [draft, setDraft] = useState(EMPTY);
  const [editing, setEditing] = useState<string | null>(null); // name of the item being edited
  const [baseline, setBaseline] = useState(EMPTY);
  const dirty = JSON.stringify(draft) !== JSON.stringify(baseline);
  const sheet = useActiveSheet(dirty);

  const allItems = itemsStore.use();
  const items = useMemo(() => (sheet ? (allItems[sheet.id] ?? []) : []), [allItems, sheet]);
  const [results, setResults] = useState<ItemResult[]>([]);
  const [reserved, setReserved] = useState<string[]>([]);
  const [names, setNames] = useState<string[]>([]);
  const [shown, setShown] = useState<{ built?: BuiltItem; error?: string }>({});
  const [message, setMessage] = useState<string>();
  const [attempted, setAttempted] = useState(false);
  const [refreshes, setRefreshes] = useState(0);

  const item = toItem(draft);
  const others = items.filter((i) => i.name !== editing);
  const { errors, warnings } = validateItem(item, others, reserved);
  const rangesOk = !errors.cell && !errors.index0 && !errors.index1;

  const load = (next: Draft, name: string | null) => {
    setDraft(next);
    setBaseline(next);
    setEditing(name);
    setAttempted(false);
    setMessage(undefined);
  };

  // A different sheet has its own items, so an open form no longer applies.
  const sheetId = sheet?.id;
  useEffect(() => load(EMPTY, null), [sheetId]);

  useEffect(() => {
    getHealth().then(
      (h) => setReserved(h.reserved_names),
      () => undefined,
    );
  }, []);

  useEffect(() => {
    if (!sheetId || tab !== "data") return;
    let cancelled = false;
    void readItems(excel(), items, sheetId).then((r) => {
      if (!cancelled) setResults(r);
    });
    void excel()
      .listNames()
      .then((n) => {
        if (!cancelled) setNames(n);
      });
    return () => {
      cancelled = true;
    };
  }, [items, sheetId, tab, refreshes]);

  // Live preview of the form, a moment after typing stops.
  const previewKey = rangesOk && sheetId ? JSON.stringify([item, sheetId, refreshes]) : null;
  useEffect(() => {
    if (!previewKey) {
      setShown({});
      return;
    }
    const [next, id] = JSON.parse(previewKey) as [DataItem, string];
    let cancelled = false;
    const timer = setTimeout(() => {
      const named = { ...next, name: next.name || "item" };
      void readItems(excel(), [named], id).then(([r]) => {
        if (!cancelled) setShown(isBuilt(r) ? { built: r } : { error: r.error });
      });
    }, 250);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [previewKey]);

  const save = (saved: DataItem) => {
    const next = editing ? items.map((i) => (i.name === editing ? saved : i)) : [...items, saved];
    itemsStore.set({ ...allItems, [sheet!.id]: next });
    const asDraft = toDraft(saved);
    load(asDraft, saved.name);
  };

  const submit = () => {
    setAttempted(true);
    if (Object.keys(errors).length === 0) save(item);
  };

  const remove = () => {
    itemsStore.set({ ...allItems, [sheet!.id]: items.filter((i) => i.name !== editing) });
    load(EMPTY, null);
  };

  const pick = async (apply: (value: string) => void) => {
    try {
      apply(localize(await excel().selection(), sheet!.name));
      setMessage(undefined);
    } catch (e) {
      setMessage(`Select cells in the sheet first. (${(e as Error).message})`);
    }
  };

  const createName = async () => {
    setAttempted(true);
    if (Object.keys(errors).length > 0) return;
    try {
      const [read] = await excel().readRefs([item.cellRange], sheet!.id);
      if ("error" in read) throw new Error(read.error);
      await excel().createName(item.name, read.ref!);
      save({ ...item, cellRange: item.name });
      setRefreshes((n) => n + 1);
    } catch (e) {
      setMessage(`Could not create the name: ${(e as Error).message}`);
    }
  };

  const set = (patch: Partial<Draft>) => setDraft({ ...draft, ...patch });
  const setIndex = (i: 0 | 1, value: string) => {
    const index: [string, string] = [...draft.index];
    index[i] = value;
    set({ index });
  };

  if (!sheet) return null;

  const overlap = overlapWarnings(results.filter(isBuilt));
  const shownPreview = shown.built && preview(shown.built);

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <datalist id={NAMES_LIST}>
        {names.map((n) => (
          <option key={n} value={n} />
        ))}
      </datalist>

      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <Text weight="semibold">Data items on {sheet.name}</Text>
        <Button size="small" appearance="subtle" onClick={() => setRefreshes((n) => n + 1)}>
          Refresh
        </Button>
      </div>

      {items.length === 0 ? (
        <Text size={200}>
          No data items yet. A data item binds a Python name to cells. Add the first one below.
        </Text>
      ) : (
        <ul style={{ listStyle: "none", margin: 0, padding: 0, display: "grid", gap: 2 }}>
          {items.map((i, at) => {
            const result = results.find((r) => r.item.name === i.name) ?? results[at];
            return (
              <li key={i.name}>
                <Button
                  size="small"
                  appearance={i.name === editing ? "secondary" : "subtle"}
                  onClick={() => load(toDraft(i), i.name)}
                  style={{ width: "100%", justifyContent: "space-between", fontWeight: 400 }}
                >
                  <span style={{ fontFamily: MONO, overflow: "hidden", textOverflow: "ellipsis" }}>
                    {i.name}
                    {i.readOnly && " (read-only)"}
                  </span>
                  {result && !isBuilt(result) ? (
                    <Badge color="danger" size="small">
                      Problem
                    </Badge>
                  ) : (
                    <Text size={100}>
                      {TYPES.find(([t]) => t === i.type)![1]}
                      {result && i.type !== "single" && `, ${entries(result.count)}`}
                    </Text>
                  )}
                </Button>
              </li>
            );
          })}
        </ul>
      )}
      {results
        .filter((r) => !isBuilt(r))
        .map((r) => (
          <Text key={r.item.name} size={200} role="alert">
            {"error" in r && r.error}
          </Text>
        ))}
      {overlap.map((w) => (
        <Text key={w} size={200}>
          Warning: {w}
        </Text>
      ))}

      <form
        onSubmit={(e) => {
          e.preventDefault();
          submit();
        }}
        style={{ display: "grid", gap: 8, borderTop: "1px solid #8884", paddingTop: 12 }}
      >
        <Field label="Name" validationMessage={attempted || draft.name ? errors.name : undefined}>
          <Input
            size="small"
            value={draft.name}
            onChange={(_, d) => set({ name: d.value })}
            style={{ fontFamily: MONO }}
          />
        </Field>

        <Field label="Type">
          <Select
            size="small"
            value={draft.type}
            onChange={(_, d) => set({ type: d.value as ItemType })}
          >
            {TYPES.map(([value, label]) => (
              <option key={value} value={value}>
                {label}
              </option>
            ))}
          </Select>
        </Field>

        {draft.type === "indexed" && (
          <RangeField
            label="Index Range(s)"
            hint="One range, or several separated by commas for tuple keys."
            value={draft.index[0]}
            error={errors.index0}
            showError={attempted || draft.index[0] !== ""}
            onChange={(v) => setIndex(0, v)}
            onPick={() => void pick((v) => setIndex(0, v))}
          />
        )}
        {draft.type === "table" && (
          <>
            <RangeField
              label="Row Index"
              hint="The keys down the side of the table."
              value={draft.index[0]}
              error={errors.index0}
              showError={attempted || draft.index[0] !== ""}
              onChange={(v) => setIndex(0, v)}
              onPick={() => void pick((v) => setIndex(0, v))}
            />
            <RangeField
              label="Column Index"
              hint="The keys across the top of the table."
              value={draft.index[1]}
              error={errors.index1}
              showError={attempted || draft.index[1] !== ""}
              onChange={(v) => setIndex(1, v)}
              onPick={() => void pick((v) => setIndex(1, v))}
            />
          </>
        )}
        <RangeField
          label="Cell Range"
          hint="An address, named range, or Table1[Col]."
          value={draft.cell}
          error={errors.cell}
          showError={attempted || draft.cell !== ""}
          onChange={(cell) => set({ cell })}
          onPick={() => void pick((cell) => set({ cell }))}
        />

        <Checkbox
          label="Read-only: never write results to these cells"
          checked={draft.readOnly}
          onChange={(_, d) => set({ readOnly: Boolean(d.checked) })}
        />

        {warnings.map((w) => (
          <Text key={w} size={200}>
            Warning: {w}
          </Text>
        ))}

        <div
          data-testid="preview"
          style={{ border: "1px solid #8884", borderRadius: 4, padding: 8, minHeight: 48 }}
        >
          <Text size={100} block>
            {shownPreview ? `Python sees: ${shownPreview.summary}` : "Python preview"}
          </Text>
          {shownPreview && (
            <pre style={{ margin: "4px 0 0", fontFamily: MONO, fontSize: 11, overflow: "auto" }}>
              {(item.name || "item") + " = " + shownPreview.code}
            </pre>
          )}
          {shown.error && (
            <Text size={200} block role="alert">
              {shown.error}
            </Text>
          )}
        </div>

        <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
          {editing ? (
            <>
              <Button size="small" appearance="primary" type="submit">
                Update Data Item
              </Button>
              <Button size="small" onClick={remove}>
                Delete
              </Button>
              <Button size="small" onClick={() => load(EMPTY, null)}>
                New
              </Button>
            </>
          ) : (
            <Button size="small" appearance="primary" type="submit">
              Add
            </Button>
          )}
          {isA1(draft.cell) && (
            <Button size="small" onClick={() => void createName()}>
              Create named range
            </Button>
          )}
        </div>
        {message && (
          <Text size={200} role="alert">
            {message}
          </Text>
        )}
        {dirty && sheet && (
          <Text size={100}>
            Editing for {sheet.name}. You can select cells on another sheet and pick them here.
          </Text>
        )}
      </form>
    </div>
  );
}
