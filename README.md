# RCI Serve

One home for every RCI team someone serves on, published at **https://astmacary.github.io/**.

The four team apps live "inside" this address, so they open within the same installed app and share one sign-in:

| Team | App |
|---|---|
| Productions | https://astmacary.github.io/rci-productions/ |
| Security | https://astmacary.github.io/rci-security/ |
| Ushers & Altar Workers | https://astmacary.github.io/rci-ushers/ |
| Worship & Arts | https://astmacary.github.io/rci-worship-arts/ |

- `index.html` — the hub: Home (your teams, requests from every team with Accept, next 7 days), My Calendar, Teams, Profile.
- `sw.js` — offline helper for the hub's own files only; never touches the team apps.
- `manifest.webmanifest` — install settings; scope `/` covers all four team apps.
- The Switch Team button inside each app is `team-switch.js` in that app's own repository (identical copies).

When the hub changes, bump both `APP_VERSION` in `index.html` and `CACHE` in `sw.js`.
