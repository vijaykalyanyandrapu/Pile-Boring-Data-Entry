# Pile Log Web Database

A local web application based on the supplied 16B pile bore-log structure.

## Requirements
- Windows/macOS/Linux
- Node.js 18+
- A modern browser

## Install
1. Extract the project.
2. Open a terminal in the project folder.
3. Run:
   npm install
4. Run:
   npm start
5. Open:
   http://localhost:3000

The SQLite database is created automatically at `data/pilelog.db`.

## Features
- New pile form covering every field on the 16B reference sheet: company name, client name, serial no., report date, project location, rig no., drawing no., dimensions, boring-completed date, cage/trimmer lowering, flushing, reinforcement, and concreting details
- Edit existing pile
- Delete pile
- Search piles
- Boring Details table supports **adding as many rows as you need** ("+ Add boring row" button) — the exported sheet grows to fit
- Excel export, one worksheet per pile, worksheet name = Pile ID
- Company name, client name, project location, serial no. and report date are all editable per pile (nothing hardcoded)
- Formula cells for penetration, bore depth, reinforcement totals/weights and theoretical concrete volume, matching the original 16B formulas
- Reinforcement rows are automatically totalled by bar diameter (25/16/12/8 mm), same as the reference sheet

## Important
The web application deliberately separates the database from Excel. The JSON data file is the source of truth; Excel is a generated report.

The exporter rebuilds the 16B layout (merged cells, headings, and formulas) to match the reference workbook. Because the Boring Details table can now have any number of rows, the sheet automatically shifts the Cage/Trimmer/Flushing/Reinforcement/Concreting sections down if a pile has more than the original ~21 boring rows, so nothing overlaps.

## Database
This version uses data/pilelog-data.json and does not require SQLite, better-sqlite3, node-gyp, Python, or Visual Studio Build Tools.

## Username & password login
The app is gated behind a single shared username and password (not per-person accounts).
- Locally, set them via environment variables before starting, e.g. on macOS/Linux:
  `AUTH_USERNAME=youruser AUTH_PASSWORD=yourpass SESSION_SECRET=someRandomLongString npm start`
  On Windows PowerShell:
  `$env:AUTH_USERNAME="youruser"; $env:AUTH_PASSWORD="yourpass"; $env:SESSION_SECRET="someRandomLongString"; npm start`
- If you don't set these, they default to username `admin` / password `changeme123` — change this before sharing the link.
- Anyone with the credentials can log in and stays logged in for 12 hours (a cookie/session), then needs to sign in again.
- There's a "Log out" button in the header that ends the session immediately.

## Deploying to Render (free) so others can access it via a link
1. Create a free GitHub account if you don't have one, and create a new repository.
2. Upload all files in this project folder to that repository (drag-and-drop upload works, or use `git push`).
3. Go to https://render.com, sign up (free), and click "New +" → "Web Service".
4. Connect your GitHub repository.
5. Render should detect `render.yaml` automatically. If it asks for settings manually, use:
   - Build command: `npm install`
   - Start command: `npm start`
6. Under "Environment", set:
   - `AUTH_USERNAME` = the username people will log in with
   - `AUTH_PASSWORD` = the password people will log in with
   - `SESSION_SECRET` = any long random string (Render can auto-generate this if you used `render.yaml`)
7. Click "Create Web Service". Render will build and deploy the app, then give you a public URL like `https://pile-log-web.onrender.com`.
8. Share that URL plus the username/password with anyone you want to have access.

### Important limitation — data persistence on Render's free tier
This app stores data in a JSON file on disk (`data/pilelog-data.json`). Render's **free** web service plan does not include a persistent disk, so the data file can be wiped whenever the service restarts or redeploys (it will also "sleep" after 15 minutes of inactivity and take ~30-60 seconds to wake up on the next visit).

If you need the entered pile data to survive restarts long-term, you have two options:
1. Upgrade to a paid Render instance (Starter plan, ~$7/month) and attach a **Persistent Disk**, then set the `DATA_DIR` environment variable to the disk's mount path (e.g. `/var/data`). This is the simplest fix.
2. Regularly use the "Download Excel" button to back up your data, and re-enter/import as needed if the service restarts.

For anything beyond casual/demo use with a handful of people, option 1 (a few dollars a month) is strongly recommended.
