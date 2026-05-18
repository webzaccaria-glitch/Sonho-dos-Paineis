// ─────────────────────────────────────────────
// SUPABASE — banco de dados em nuvem
// ─────────────────────────────────────────────
const URL  = 'https://pvmdhfevxoxrzxkbjxzr.supabase.co';
const KEY  = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB2bWRoZmV2eG94cnp4a2JqeHpyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkwNjY5NTcsImV4cCI6MjA5NDY0Mjk1N30.U-wNxHNb8enavVCEVcuMz5u61vOzsqNc0A7Wp1uAZeM';

const H = {
  'Content-Type': 'application/json',
  'apikey': KEY,
  'Authorization': `Bearer ${KEY}`,
  'Prefer': 'resolution=merge-duplicates',
};

export async function dbGet(key, fallback) {
  try {
    const res  = await fetch(`${URL}/rest/v1/app_data?key=eq.${encodeURIComponent(key)}&select=value`, { headers: H });
    if (!res.ok) return fallback;
    const rows = await res.json();
    return rows.length > 0 ? rows[0].value : fallback;
  } catch {
    return fallback;
  }
}

export async function dbSet(key, value) {
  try {
    await fetch(`${URL}/rest/v1/app_data`, {
      method:  'POST',
      headers: H,
      body:    JSON.stringify({ key, value, updated_at: new Date().toISOString() }),
    });
  } catch {}
}
