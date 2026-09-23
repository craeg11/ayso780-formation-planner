# AYSO Lineup Planner PWA

This version is local-first: it works with locally cached data offline and synchronizes confirmed changes to Google Sheets when it is online.

## Install

1. In the Google Sheet’s Apps Script project, replace Code.gs and run setupPlanner once. This adds the sync-receipt sheet and upgrades the Teams headers.
2. Deploy a new Apps Script web-app version. Keep the deployment set to execute as you.
3. Put the deployment URL ending in /exec in SYNC_SERVER_URL near the top of index.html.
4. Publish index.html, sw.js, manifest.json, and icon.svg to the GitHub Pages repository root.

## Important behavior

- The first online visit loads the shared team list. Opening a team with its PIN loads its roster and saved plans onto that device.
- Changes are saved locally immediately and remain queued until the server confirms them.
- A green Online badge means there are no known queued changes. Waiting to sync or Sync error means changes should remain on the device until resolved.
- New team names are generated automatically, for example U14G-Craig_Fall_2026.

PINs are hashed when new teams are saved. A four-digit PIN remains a convenience lock for trusted coaches, not strong account-level security.
