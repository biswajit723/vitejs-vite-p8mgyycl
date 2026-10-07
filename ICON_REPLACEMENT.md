# App logo replacement

Replace these files inside `public`:
- `app-logo-192.png` (exactly 192x192)
- `app-logo-512.png` (exactly 512x512)
- `favicon.ico` (Windows URL shortcut icon)

For compatibility, also replace:
- `icon-192.png`
- `icon-512.png`

After changing a logo, increment every `v=3` value in `index.html` and `manifest.webmanifest`, and change `rgt-icon-v3` in both service-worker files. Then rebuild and redeploy.

Windows keeps old shortcut icons in its icon cache. Delete the old desktop shortcut. If it is an installed Edge app, uninstall it from `edge://apps`, close Edge, reopen the deployed site, and install the app again. Do not use the browser's plain "Create shortcut" command if an installed PWA icon is required.
