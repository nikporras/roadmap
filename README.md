# Nico's Roadmap

PlayRun-branded side project planner, installable as a PWA. Shows tonight's tasks, a stripboard timeline, deadlines and blocked shift days.

All data lives in a private Google Sheet, read and saved through an Apps Script web app. This repo holds only the app code. The Sheet URL and token are entered in the app's settings and stored on each device, never in the code.

- `index.html`: the whole app
- `sw.js`: offline support
- `manifest.webmanifest` + icons: install as an app
