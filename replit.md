# Hospital San José - Computer Equipment Inventory System

## Project Overview
A Google Apps Script (GAS) web application for managing IT equipment inventory at Hospital San José. It tracks PCs, monitors, peripherals, personnel assignments, and maintenance records using Google Sheets as a database.

## Architecture
- **Backend:** Google Apps Script (`codigogs.gs`) — runs in Google's cloud, communicates with Google Sheets
- **Frontend:** Single HTML file (`index.html`) — vanilla HTML/CSS/JavaScript, served statically in Replit preview
- **Database:** Google Sheets (via SpreadsheetApp API in production GAS environment)
- **Static Server:** `server.js` — simple Node.js HTTP server to serve `index.html` for preview on port 5000

## Key Files
- `codigogs.gs` — Server-side Google Apps Script logic (CRUD, routing, sheet initialization)
- `index.html` — Full frontend (HTML + CSS + JS, ~4600 lines)
- `formulario.html` — Public ticket report form; looks up staff by DNI from `Personal` and loads assigned equipment from `Equipos`
- `server.js` — Minimal Node.js static server for Replit preview with explicit route handling and basic security headers

## Running in Replit
The workflow "Start application" runs `node server.js`, serving `index.html` on `/` and `formulario.html` on `/formulario`. The server binds to `0.0.0.0` and uses `process.env.PORT` when Replit provides one, falling back to port 5000 for local preview.

Note: The full app (with live data) must be deployed as a Google Apps Script Web App connected to a Google Sheets spreadsheet. In Replit, only the frontend UI is previewed statically — backend `google.script.run` calls won't function without the GAS environment.

## Formulario de Reporte
The public form searches staff by DNI in the `Personal` sheet. After a DNI is selected, `codigogs.gs` reads the `Equipos` sheet and returns the equipment assigned to that DNI or matching user name. If one equipment item is found it is selected automatically; if multiple are found the user chooses the affected equipment.

## External Libraries (CDN)
- `xlsx.full.min.js` — Excel export
- `html2pdf.bundle.min.js` — PDF generation
