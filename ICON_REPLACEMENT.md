# Replace the app logo

To use a different logo without changing code:

1. Prepare a square PNG logo.
2. Create a 192x192 version named `app-logo-192.png`.
3. Create a 512x512 version named `app-logo-512.png`.
4. Replace both files in the `public` folder.
5. Change the cache name in `public/service-worker.js` and `public/sw.js`, for example `rgt-icon-v3`.
6. Rebuild and redeploy.
7. Uninstall the old desktop PWA shortcut and install it again.

The manifest, favicon, Apple icon, desktop shortcut icon, Start menu icon, and taskbar icon all use these two files.
