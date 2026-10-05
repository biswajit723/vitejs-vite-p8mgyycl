# Reports Generate Tools

## Central Master Module List — Supabase

The Master Excel module list is stored centrally in Supabase so all website users can see the same modules.

### Behavior

- Admin connects the Master Excel folder once for a sync.
- Module names are extracted from the Excel `Module` / `Module Name` column.
- The Excel order is preserved.
- The server replaces the central module list atomically in Supabase.
- All users load the same central module list when opening Reports and Status.
- An already-open page refreshes the central list every 30 seconds.
- Browser close/reopen keeps the central list because Supabase is persistent.
- LocalStorage remains only as a fallback cache when the central service is unavailable.
- Existing 4-column structure, count boxes, Size Status behavior, Excel downloads, report logic, colors and layout are preserved.

### Supabase setup

1. Create a Supabase project.
2. Open **SQL Editor → New query**.
3. Paste and run the contents of `supabase.sql`.
4. Go to **Project Settings → API** and copy:
   - Project URL
   - service_role key (server-side secret)
5. Add these environment variables to Vercel:

```text
SUPABASE_URL=https://your-project-ref.supabase.co
SUPABASE_SERVICE_ROLE_KEY=...
ADMIN_SYNC_TOKEN=...
```

6. Redeploy the Vercel project after saving the variables.

### Master Excel sync

1. Open the deployed site.
2. Go to Reports and Status.
3. Click **Connect Excel Folder**.
4. Select the Master Excel folder.
5. Enter the same `ADMIN_SYNC_TOKEN` configured in Vercel when asked.
6. The module names are published to Supabase.
7. All users then receive the same central module list.
8. Other open pages refresh the module list within about 30 seconds; reopening the website also loads the central list.

### Security

- Do not expose `SUPABASE_SERVICE_ROLE_KEY` in frontend code or any `VITE_*` variable.
- Only `/api/modules` uses the service-role key.
- Publishing is protected by `ADMIN_SYNC_TOKEN`.
- The module table has RLS enabled.

### Important limitation

The central service stores the module list, not browser-local Excel file handles. Browser security prevents one user's local folder permission from being reused by other users. Existing report counts and Excel downloads still use Excel files available in the current browser session.
