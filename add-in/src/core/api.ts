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

export interface GlobalSettings {
  python_path: string;
  license_path: string;
  keep_runs: number;
  port: number;
  default_python: string;
}

export async function getGlobalSettings(): Promise<GlobalSettings> {
  const res = await apiFetch("/api/settings");
  if (!res.ok) throw new Error(`Helper returned ${res.status}`);
  return res.json();
}

export async function saveGlobalSettings(values: Partial<GlobalSettings>): Promise<GlobalSettings> {
  const res = await apiFetch("/api/settings", { method: "PUT", body: JSON.stringify(values) });
  const body = await res.json();
  if (!res.ok) throw new Error(body.error ?? `Helper returned ${res.status}`);
  return body;
}

export interface LicenseReport {
  ok: boolean;
  type: string;
  expiry: string | null;
  restricted: boolean;
  offline: boolean;
  message: string;
}

export async function checkLicense(): Promise<LicenseReport> {
  const res = await apiFetch("/api/license/check", { method: "POST" });
  if (res.status === 409) throw new Error("A solve is running. Check again when it finishes.");
  if (!res.ok) throw new Error(`Helper returned ${res.status}`);
  return res.json();
}
