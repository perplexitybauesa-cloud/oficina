# Oficina Pobresa i Eficiència Energètica — Sant Adrià de Besòs

## Overview
A CRM-style web application for managing energy poverty consultations for the Office of Poverty and Energy Efficiency in Sant Adrià de Besòs. Originally a Google Apps Script project, adapted to run as a Node.js/Express app on Replit with Google Sheets as the database.

## Architecture
- **Backend**: Node.js + Express (`server.js`) on port 5000
- **Frontend**: Vanilla HTML/CSS/JS (`index.html`) served by Express as static file
- **Database**: Google Sheets (Spreadsheet ID: `13-HCGphpkB8zi4m6299YdMXEADWtXLc0Hy1CJot_Cj0`)
- **Auth**: Google Service Account credentials via `GOOGLE_CREDENTIALS_JSON` secret (or fallback `google-credentials.json` file)

## Key Files
- `server.js` — Express server with all API endpoints replicating the original `Code.gs` logic
- `index.html` — Full frontend UI (adapted from `Index.html`, replaces `google.script.run` with `fetch`)
- `Index.html` / `Code.gs` — Original Google Apps Script source files (kept for reference)
- `google-credentials.json` — Service account credentials (gitignored, not committed)

## API Endpoints
- `GET /api/health` — Health check
- `GET /api/dashboard` — Dashboard KPIs and stats
- `GET /api/users` — All users (for global search cache)
- `GET /api/user/:dni` — User profile + history + notes
- `POST /api/atencio` — Save new consultation (user + attention + assessments + notes)
- `POST /api/nota` — Save internal note
- `GET /api/gestions` — Pending management actions
- `POST /api/gestions/:id/estat` — Update management status
- `GET /api/registres` — Recent attendance history (with date filters)
- `GET /api/export` — Returns Google Sheets URL for export

## Google Sheets Structure
Sheets used: `USUARIS`, `ATENCIONS`, `ASSESSORAMENTS`, `GESTIONS_LLUM`, `GESTIONS_GAS`, `GESTIONS_AIGUA`, `ALTRES_GESTIONS`, `NOTES`

## Setup Notes
- The `GOOGLE_CREDENTIALS_JSON` environment secret should contain the full service account JSON
- The service account email must be shared on the Google Sheet with Editor permissions
- Run with `node server.js` or `npm start`
