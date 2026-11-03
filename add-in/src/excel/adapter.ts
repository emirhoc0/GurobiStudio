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

export interface ExcelAdapter {
  activeSheet(): Promise<{ id: string; name: string }>;
  /** Reads every range in one batched call. */
  readRanges(refs: RangeRef[]): Promise<Cell[][][]>;
  /** Writes every range in one batched call. Values only, never formatting. */
  writeRanges(writes: RangeWrite[]): Promise<void>;
}
