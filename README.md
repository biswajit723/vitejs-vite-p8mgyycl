# Reports Generate Tools

## Central Master Module List

The app now supports a persistent central module list for deployed environments.

### What changed

- The first successful Master Excel connection still reads the module names exactly as before.
- The extracted module names are cached locally for offline/browser-reopen fallback.
- The app also publishes the module names to `/api/modules` when an admin sync token is provided.
- When any user opens the Reports page, the app first tries to load the centrally stored module list.
- The same saved module names are therefore shown to users on different computers/browsers after the central service is configured.
- The current 4-column structure, count boxes, Size Status table, Excel downloads, report logic, colors, and layout are otherwise preserved.

### Required deployment setup

This project uses an Upstash Redis REST endpoint through the Vercel serverless function in `api/modules.js`. No Redis package is required in the browser bundle.

Create a Redis database and add these environment variables to the Vercel project:

```text
UPSTASH_REDIS_REST_URL=...
UPSTASH_REDIS_REST_TOKEN=...
ADMIN_SYNC_TOKEN=...
```

See `.env.example` for the variable names.

### First-time Master Excel publish

1. Deploy the project with the three environment variables above.
2. Open the deployed website and go to Reports and Status.
3. Click **Connect Excel Folder** and select the Master Excel folder.
4. When prompted for `ADMIN SYNC TOKEN`, enter the same value configured as `ADMIN_SYNC_TOKEN` in Vercel.
5. The module names are published centrally.
6. Other users do not need to connect the Excel folder just to see the module list.
7. Closing/reopening the browser does not remove the central module list.

### Important limitation

The central service stores the **module list**, not the user's local Excel file handles. Browser security prevents one user's local folder permission from being reused by other users. Existing report counts and Excel downloads still use the Excel files available in the current browser session.
