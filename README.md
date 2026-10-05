# Report Generate Tool — Supabase Central Module Fix

Central module list behavior:
- Admin syncs the Master Excel module names once.
- Names are stored centrally in Supabase.
- All users read the same active module list.
- Browser close/reopen keeps the centrally published list.
- Open pages refresh the module list every 30 seconds.
- Public reads use the Vercel API first and direct Supabase REST as a fallback.

Required Vercel environment variables:
- VITE_SUPABASE_URL
- VITE_SUPABASE_PUBLISHABLE_KEY (or VITE_SUPABASE_ANON_KEY)
- SUPABASE_URL
- SUPABASE_SERVICE_ROLE_KEY
- ADMIN_SYNC_TOKEN

Run `supabase.sql` once in Supabase SQL Editor.
