export type Cell = string | number | boolean;

export interface RangeRef {
  sheetId: string;
  address: string;
}

export interface RangeWrite {
  ref: RangeRef;
  /** null leaves a cell untouched; "" clears it. */
  values: (Cell | null)[][];
}

export interface RefRead {
  /** null when the reference has values but no address, as a formula name can. Not writable. */
  ref: RangeRef | null;
  values: Cell[][];
  /** Positions of cells holding Excel errors such as #N/A. Their text is in values. */
  errorCells: [number, number][];
}

export type RefResult = RefRead | { error: string };

/** The pane's side of the pop-out editor. */
export interface EditorDialog {
  send(message: string): void;
}

export interface ExcelAdapter {
  activeSheet(): Promise<{ id: string; name: string }>;
  onSheetActivated(listener: () => void): void;
  /** Resolves and reads every reference in one batched call. A bare A1 address means sheetId. */
  readRefs(refs: string[], sheetId: string): Promise<RefResult[]>;
  /** Writes every range in one batched call. Values only, never formatting. */
  writeRanges(writes: RangeWrite[]): Promise<void>;
  /** The current selection's address, including its sheet. Comma separated for several areas. */
  selection(): Promise<string>;
  /** Named ranges and tables, for suggestions in range fields. */
  listNames(): Promise<string[]>;
  /** Creates a workbook-level name. Rejects when Excel refuses the name or it already exists. */
  createName(name: string, ref: RangeRef): Promise<void>;
  /** For each write, which of its non-null cells currently hold a formula. */
  formulaCells(writes: RangeWrite[]): Promise<boolean[][][]>;
  /** Opens this page with view=editor in a dialog. */
  openEditorDialog(
    onMessage: (message: string) => void,
    onClosed: () => void,
  ): Promise<EditorDialog>;
  /** Called inside that dialog: listens to the pane and returns how to send to it. */
  editorChannel(onMessage: (message: string) => void): (message: string) => void;
  /** Excel's own theme, or null when it doesn't say. */
  officeDark(): boolean | null;
}
