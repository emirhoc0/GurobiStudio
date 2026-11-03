import { useEffect, useRef, useState } from "react";
import { excel } from "../excel";

export interface Sheet {
  id: string;
  name: string;
}

/**
 * Follows the active worksheet. While pinned it stays put, so a range can be picked elsewhere, and
 * after that it waits for the next sheet change, so a saved item doesn't seem to vanish.
 */
export function useActiveSheet(pinned = false): Sheet | null {
  const [sheet, setSheet] = useState<Sheet | null>(null);
  const pinnedRef = useRef(pinned);
  pinnedRef.current = pinned;

  useEffect(() => {
    void excel().activeSheet().then(setSheet);
    excel().onSheetActivated(() => {
      if (!pinnedRef.current) void excel().activeSheet().then(setSheet);
    });
  }, []);

  return sheet;
}
