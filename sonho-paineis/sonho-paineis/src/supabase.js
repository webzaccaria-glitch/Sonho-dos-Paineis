// ─────────────────────────────────────────────
// SUPABASE — armazenamento em nuvem
// ─────────────────────────────────────────────
const SUPA_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPA_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

const headers = {
  'Content-Type': 'application/json',
  'apikey': SUPA_KEY,
  'Authorization': `Bearer ${SUPA_KEY}`,
  'Prefer': 'resolution=merge-duplicates',
};

export async function dbGet(key, fallback) {
  try {
    const url = `${SUPA_URL}/rest/v1/app_data?key=eq.${encodeURIComponent(key)}&select=value`;
    const res  = await fetch(url, { headers });
    if (!res.ok) return fallback;
    const rows = await res.json();
    return rows.length ? rows[0].value : fallback;
  } catch {
    return fallback;
  }
}

export async function dbSet(key, value) {
  try {
    const url = `${SUPA_URL}/rest/v1/app_data`;
    await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify({ key, value, updated_at: new Date().toISOString() }),
    });
  } catch {}
}
