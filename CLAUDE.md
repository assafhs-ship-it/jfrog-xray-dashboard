# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a JFrog interview home assignment for the role of **Senior PM, Security Adoption**. It consists of two deliverables:

- **Part 1** (`Part 1.docx`) — Adoption strategy document covering Personas, Journey Map, and a 30-day plan. Prepared for a live 20-minute brainstorm in the interview. Written in English and Hebrew.
- **Part 2** (`index.html`) — A live, self-contained browser dashboard tracking Xray customer adoption. Built with vanilla HTML/CSS/JS and Chart.js (CDN). No build step, no dependencies to install.

## Running the Dashboard

Open directly in a browser — no server required:
```
open index.html
```

Or serve locally if needed:
```
python3 -m http.server 8080
```

## Architecture of index.html

Single-file app (~1100 lines). All CSS, HTML, and JS are inline. Structure:

- **CSS variables** at the top define the JFrog design tokens (`--navy`, `--green`, stage colors `--s0`–`--s3`)
- **Mock data** — `const customers` array (18 accounts) is the sole data source; no backend
- **Filter state** — `filtered` array is derived from `customers` on every filter/sort change
- **Rendering** — `renderTable()`, `renderKPIs()`, `updateDonut()`, `renderRiskList()` are called together via `applyFilters()` whenever any input changes
- **Charts** — Chart.js donut (stage distribution, re-renders on filter) and a static line chart (6-month trend, initialized once)
- **Column tooltips** — each `<th>` has an `.info-btn` that toggles a `.th-tooltip` div; a shared `closeAllTooltips()` resets all; viewport-overflow is corrected via `getBoundingClientRect()` after open
- **Bio modal** — triggered by clicking the candidate chip in the header; closed by ✕ button, overlay click, or Escape key

## Deployment

Hosted on GitHub Pages at:
```
https://assafhs-ship-it.github.io/jfrog-xray-dashboard/
```
Repo: `https://github.com/assafhs-ship-it/jfrog-xray-dashboard`

To push updates:
```
git add index.html
git commit -m "your message"
git push
```

## Generating Part 1.docx

`generate_part1.py` uses `python-docx` to regenerate the Word document:
```
python3 generate_part1.py
```
