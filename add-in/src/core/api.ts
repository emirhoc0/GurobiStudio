const token = document.querySelector<HTMLMetaElement>('meta[name="gs-token"]')?.content ?? "";

export function apiFetch(path: string, init: RequestInit = {}): Promise<Response> {
  return fetch(path, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...init.headers,
    },
  });
}

export interface Health {
  version: string;
  python: string;
  gurobipy: string | null;
  reserved_names: string[];
}

export async function getHealth(): Promise<Health> {
  const res = await apiFetch("/api/health");
  if (!res.ok) throw new Error(`Helper returned ${res.status}`);
  return res.json();
}

/** Reveals the newest run's folder in Finder. False when there are no runs yet. */
export async function revealLatestRun(): Promise<boolean> {
  const res = await apiFetch("/api/runs/latest/reveal", { method: "POST" });
  if (res.status === 404) return false;
  if (!res.ok) throw new Error(`Helper returned ${res.status}`);
  return true;
}
