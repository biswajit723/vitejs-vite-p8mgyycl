const KEY = 'rgt:master-module-list:v1';

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
    // Preserve the order from the Master Excel.
    modules.push(moduleName);
  }
  return modules;
}

function configError() {
  return !process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY;
}

function supabaseHeaders() {
  return {
    apikey: process.env.SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${process.env.SUPABASE_SERVICE_ROLE_KEY}`,
    'Content-Type': 'application/json',
  };
}

async function supabaseRequest(path, options = {}) {
  if (configError()) throw new Error('Supabase environment variables are missing.');
  const response = await fetch(`${process.env.SUPABASE_URL}/rest/v1${path}`, {
    ...options,
    headers: {
      ...supabaseHeaders(),
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    const body = await response.text().catch(() => '');
    throw new Error(`Supabase returned HTTP ${response.status}${body ? `: ${body}` : ''}`);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

async function getCentralModules() {
  const rows = await supabaseRequest(
    '/modules?select=module_name,sort_order&active=eq.true&order=sort_order.asc'
  );
  return normalizeModules((rows || []).map((row) => row?.module_name));
}

async function replaceCentralModules(modules) {
  // Atomic replacement is implemented as a Postgres RPC.
  const rows = await supabaseRequest('/rpc/replace_modules', {
    method: 'POST',
    body: JSON.stringify({ p_modules: modules }),
  });
  return normalizeModules((rows || []).map((row) => row?.module_name));
}

export default async function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const modules = await getCentralModules();
      return res.status(200).json({ modules, source: 'supabase' });
    } catch (error) {
      console.error('Module GET failed:', error);
      return res.status(503).json({ error: 'Central module service is not configured or unavailable.' });
    }
  }

  if (req.method === 'POST') {
    const expectedToken = process.env.ADMIN_SYNC_TOKEN;
    const providedToken = Array.isArray(req.headers['x-admin-token'])
      ? req.headers['x-admin-token'][0]
      : req.headers['x-admin-token'];

    if (!expectedToken || providedToken !== expectedToken) {
      return res.status(401).json({ error: 'Unauthorized admin sync request.' });
    }

    try {
      const modules = normalizeModules(req.body?.modules);
      if (!modules.length) return res.status(400).json({ error: 'No module names were supplied.' });

      const savedModules = await replaceCentralModules(modules);
      return res.status(200).json({ modules: savedModules, source: 'supabase' });
    } catch (error) {
      console.error('Module POST failed:', error);
      return res.status(503).json({ error: 'Central module service is unavailable.' });
    }
  }

  res.setHeader('Allow', 'GET, POST');
  return res.status(405).json({ error: 'Method not allowed.' });
}
