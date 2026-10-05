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


## Vercel environment variables

Required for the admin publish API:
- `ADMIN_SYNC_TOKEN`
- `SUPABASE_SECRET_KEY` (recommended) OR `SUPABASE_SERVICE_ROLE_KEY` (legacy)
- `SUPABASE_URL` OR `NEXT_PUBLIC_SUPABASE_URL`

For public browser reads, either `VITE_SUPABASE_URL` + `VITE_SUPABASE_PUBLISHABLE_KEY`, or the existing `NEXT_PUBLIC_SUPABASE_URL` + `NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY` variables can be used. The Vite config maps the latter into the client build.

After changing Vercel environment variables, redeploy.
