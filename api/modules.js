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
    modules.push(moduleName);
  }
  return modules.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
}

function configError() {
  return !process.env.UPSTASH_REDIS_REST_URL || !process.env.UPSTASH_REDIS_REST_TOKEN;
}

async function redisCommand(command, args = []) {
  if (configError()) throw new Error('Upstash Redis environment variables are missing.');
  const response = await fetch(process.env.UPSTASH_REDIS_REST_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.UPSTASH_REDIS_REST_TOKEN}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify([command, ...args]),
  });
  if (!response.ok) throw new Error(`Upstash returned HTTP ${response.status}`);
  const data = await response.json();
  if (data?.error) throw new Error(data.error);
  return data?.result;
}

export default async function handler(req, res) {
  if (req.method === 'GET') {
    try {
      const raw = await redisCommand('GET', [KEY]);
      const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
      const modules = normalizeModules(parsed?.modules ?? parsed ?? []);
      return res.status(200).json({ modules, source: 'central' });
    } catch (error) {
      console.error('Module GET failed:', error);
      return res.status(503).json({ error: 'Central module service is not configured or unavailable.' });
    }
  }

  if (req.method === 'POST') {
    const expectedToken = process.env.ADMIN_SYNC_TOKEN;
    const providedToken = Array.isArray(req.headers['x-admin-token']) ? req.headers['x-admin-token'][0] : req.headers['x-admin-token'];
    if (!expectedToken || providedToken !== expectedToken) {
      return res.status(401).json({ error: 'Unauthorized admin sync request.' });
    }

    try {
      const modules = normalizeModules(req.body?.modules);
      if (!modules.length) return res.status(400).json({ error: 'No module names were supplied.' });
      await redisCommand('SET', [KEY, JSON.stringify({ modules, updatedAt: new Date().toISOString() })]);
      return res.status(200).json({ modules, source: 'central' });
    } catch (error) {
      console.error('Module POST failed:', error);
      return res.status(503).json({ error: 'Central module service is unavailable.' });
    }
  }

  res.setHeader('Allow', 'GET, POST');
  return res.status(405).json({ error: 'Method not allowed.' });
}
