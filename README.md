# Report Generate Tool — Supabase Central Module List

## Vercel Environment Variables
Set these for **Preview and Production**:

- `SUPABASE_URL` = Supabase Project URL
- `SUPABASE_SECRET_KEY` = Supabase Secret key (`sb_secret_...`) — server only
- `ADMIN_SYNC_TOKEN` = your private admin sync token

The public browser read path may also use:
- `NEXT_PUBLIC_SUPABASE_URL`
- `NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY`

## Supabase setup
Run `supabase.sql` once in Supabase SQL Editor. It creates `public.modules` and the read policy.

## Important
Do not put `SUPABASE_SECRET_KEY` in source code or any `NEXT_PUBLIC_` variable.

## Central sync flow
Master Excel -> `/api/modules` POST -> Supabase `public.modules` -> all users read the same module list.
