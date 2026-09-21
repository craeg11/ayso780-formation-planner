# AYSO Lineup Planner

This is a mobile-friendly Google Apps Script web app. It stores only team names, optional coach names, player display names, jersey numbers, and lineup plans.

Each coach creates or selects a team. Every team has its own roster, saved plans, unavailable-player list, and quarter-by-quarter visual field. The player dropdowns are placed on a soccer field so younger players can see the position as well as its name.

## Set up

1. Create a blank Google Sheet.
2. In the Sheet, choose **Extensions → Apps Script**.
3. Replace the default script with `Code.gs`, add an HTML file named `Index`, and paste in `Index.html`.
4. Replace the project manifest with `appsscript.json` if prompted.
5. Run `setupPlanner` once and approve the requested permission.
6. Deploy as a web app: **Deploy → New deployment → Web app**. To have coaches sign in, select an access setting that requires a Google account. Share the resulting URL with the coaching group.

The app creates `Teams`, `Team Rosters`, and `Lineup Plans` tabs automatically. Keep the spreadsheet private; coaches use the shared web-app link.

## Important access note

The sign-in setting asks people to authenticate with Google, but this starter app does not yet enforce that a particular Google account may access only one team. Coaches choose their own team from the list. For a small, trusted AYSO coaching group, that is usually the simplest setup. If you need strict team-by-team privacy, use separate planner copies per team or add a team-membership approval feature before sharing broadly.
