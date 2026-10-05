function normalizeModules(value) {
  if (!Array.isArray(value)) return [];
  const seen = new Set();
  const modules = [];
  for (const item of value) {
    const moduleName = String(item ?? '').trim();
    if (!moduleName) continue;
    const key = moduleName.toLowerCase().replace(/\s+/g, '');
    if (seen.has(key)) continue;
    seen.add(key);
    modules.push(moduleName);
  }
  return modules;
}

function getSupabaseConfig() {
  const url = String(process.env.SUPABASE_URL || process.env.NEXT_PUBLIC_SUPABASE_URL || '').replace(/\/$/, '');
  const key = String(process.env.SUPABASE_SECRET_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY || '').trim();
  return { url, key };
}

async function supabaseRequest(path, options = {}) {
  const { url, key } = getSupabaseConfig();
  if (!url || !key) throw new Error('Missing SUPABASE_URL/NEXT_PUBLIC_SUPABASE_URL or SUPABASE_SECRET_KEY/SUPABASE_SERVICE_ROLE_KEY in Vercel.');

  const response = await fetch(`${url}/rest/v1${path}`, {
    ...options,
    headers: {
      apikey: key,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });

  const body = await response.text();
  if (!response.ok) {
    throw new Error(`Supabase HTTP ${response.status}: ${body || response.statusText}`);
  }
  try { return body ? JSON.parse(body) : null; } catch { return body; }
}

async function getCentralModules() {
  const rows = await supabaseRequest('/modules?select=module_name,sort_order&active=eq.true&order=sort_order.asc');
  return normalizeModules((rows || []).map(row => row?.module_name));
}

async function replaceCentralModules(modules) {
  // Use the REST table API directly instead of an RPC. This avoids RPC signature/grant
  // mismatches and works with both the current sb_secret_* key and legacy service_role key.
  await supabaseRequest('/modules?id=not.is.null', {
    method: 'DELETE',
    headers: { Prefer: 'return=minimal' },
  });

  const rows = modules.map((moduleName, index) => ({
    module_name: moduleName,
    sort_order: index + 1,
    active: true,
  }));

  const saved = await supabaseRequest('/modules?select=module_name,sort_order&order=sort_order.asc', {
    method: 'POST',
    headers: { Prefer: 'return=representation' },
    body: JSON.stringify(rows),
  });

  return normalizeModules((saved || []).map(row => row?.module_name));
}

export default async function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const modules = await getCentralModules();
      return res.status(200).json({ modules, source: 'supabase' });
    } catch (error) {
      console.error('Module GET failed:', error);
      return res.status(503).json({ error: error.message });
    }
  }

  if (req.method === 'POST') {
    const expectedToken = String(process.env.ADMIN_SYNC_TOKEN || '');
    const providedToken = Array.isArray(req.headers['x-admin-token'])
      ? req.headers['x-admin-token'][0]
      : String(req.headers['x-admin-token'] || '');

    if (!expectedToken || providedToken !== expectedToken) {
      return res.status(401).json({ error: 'Unauthorized admin sync request.' });
    }

    try {
      const modules = normalizeModules(req.body?.modules);
      if (!modules.length) return res.status(400).json({ error: 'No module names were supplied.' });

      const savedModules = await replaceCentralModules(modules);
      if (savedModules.length !== modules.length) {
        throw new Error(`Supabase saved ${savedModules.length} modules, but ${modules.length} were supplied.`);
      }
      return res.status(200).json({ modules: savedModules, source: 'supabase' });
    } catch (error) {
      console.error('Module POST failed:', error);
      return res.status(503).json({ error: error.message });
    }
  }

  res.setHeader('Allow', 'GET, POST');
  return res.status(405).json({ error: 'Method not allowed.' });
}
