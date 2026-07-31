<!--
Copyright 2026 Google LLC

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
-->
# 🏒 NHL Analytics & Executive Presentation Engine
### Enterprise-Grade Google Apps Script, TypeScript, & Google Slides Automation Workflow

[![TypeScript](https://img.shields.io/badge/TypeScript-5.9-blue.svg)](https://www.typescriptlang.org/)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8%20Runtime-4285F4.svg)](https://developers.google.com/apps-script)
[![Bundler](https://img.shields.io/badge/Bundled%20With-Rollup-ff3e00.svg)](https://rollupjs.org/)
[![Testing](https://img.shields.io/badge/Tested%20With-Jest-C21325.svg)](https://jestjs.io/)
[![License](https://img.shields.io/badge/License-Apache%202.0-green.svg)](LICENSE)

---

## 📖 The Story: Modernizing Sports Analytics & Executive Reporting

Every modern sports organization, media analyst, and sports performance group relies heavily on data spreadsheets to record game metrics, track team performance, evaluate rest intervals, and analyze home versus away trends. However, converting thousands of raw transactional game rows into executive-ready presentation decks is historically a manual, error-prone, and time-consuming bottleneck. Furthermore, standard web scripts edited directly in browser IDEs often produce dull data dumps, lack visual polish, and turn into unmaintainable legacy script debt over time.

**This project elevates Google Workspace automation into an enterprise-grade software engineering discipline.**

Instead of relying on fragile, uncompiled scripts, this portfolio application treats Google Apps Script as a compiled, fully-typed deployment target. Built with TypeScript, bundled using Rollup, and tested with Jest, it transforms multi-dimensional NHL game datasets into high-impact Google Slides presentation decks—complete with executive KPI callout cards, visual progress bars, and dynamic embedded Google Sheets charts wrapped in an automated multi-environment deployment pipeline (`clasp`).

---

## ✨ Key Features & Business Impact

- ⚡ **Automated Slide Deck Generation**: Creates board-ready Google Slides decks on demand directly from Google Sheets data with a single click from a native spreadsheet menu (`Generate Report`).
- 📈 **Dynamic Embedded Google Sheets Charts**: Automatically builds visual column and bar charts from team analytics and embeds them seamlessly onto slide pages as high-resolution visual charts.
- 🎨 **Visual KPI Callout Cards & Progress Bars**: Replaces raw tabular text with modern rounded metric cards (`ROUND_RECTANGLE`), left accent status strips, and visual fill progress bars for win rates, special teams, and goal differentials.
- 🏒 **Multi-Dimensional Hockey Analytics**: Computes team standings, goals for/against rates, shot differentials, Power Play % (PP), Penalty Kill % (PK), Save %, Faceoff %, and rest-day impact breakdowns.
- 🛡️ **Fail-Safe & Type-Safe Architecture**: Includes defensive fallbacks ensuring empty text frames or missing values never trigger Google Slides API exceptions.
- 🔄 **Multi-Stage Deployment Pipeline**: Seamless deployment across `Development` and `Production` Apps Script projects with auto-swapping `clasp` target configurations.
- 🧪 **Zero-Breakage CI Workflow**: Comprehensive unit testing via Jest and strict linting via ESLint/Prettier to guarantee production stability.

---

## 🏗️ Technical Stack

- **Core Language**: TypeScript 5.x compiled targeting Google Apps Script V8 Engine.
- **Workspace Services**: Google Apps Script `SlidesApp` & `SpreadsheetApp`.
- **Bundling & Optimization**: Rollup with custom global scope preservation plugins.
- **Deployment & Orchestration**: `@google/clasp` for multi-stage deployments.
- **Testing & Quality Assurance**: Jest + `ts-jest`, ESLint, Prettier, and Apache 2.0 license checks.

---

## 📁 Project Architecture

```text
nhl-auto-report/
├── src/
│   ├── index.ts           # Entry point & custom Google Sheets menu initialization ('Generate Report')
│   ├── slides-service.ts  # Visual presentation deck generator, KPI cards, & embedded chart orchestrator
│   ├── data-service.ts    # Game data parsing, stat aggregation, Home/Away splits, & rest day analytics
│   ├── example-module.ts  # Modular helper routines
│   └── types.ts           # Strictly-typed schemas for NHL game records & computed team stats
├── test/                  # Unit test suite powered by Jest
│   └── example-module.test.ts
├── .clasp-dev.json        # Staging / Development deployment target
├── .clasp-prod.json       # Live Production deployment target
├── appsscript.json        # Apps Script manifest settings
└── rollup.config.mjs      # Production bundle settings (ESM -> Apps Script standard)
```

---

## 📊 Sheet Data Schema & Analyzed Metrics

The pipeline parses raw game rows from Google Sheets formatted with key performance columns:

| Metric Category | Source Columns | Output Presentation Analytics |
| :--- | :--- | :--- |
| **Match Identity** | `date`, `venue`, `team_name`, `is_home` | Title slide, team overview header, Home/Away split |
| **Scoring & Shots** | `score`, `opp_score`, `shots`, `opp_shots` | Goals For/gm, Goals Against/gm, Goal Differential, Shot Differential |
| **Special Teams** | `power_play_goals`, `power_play_opportunities` | Power Play % (PP%), Penalty Kill % (PK%) |
| **Goaltending & Control**| `save_pct`, `faceoff_win_pct` | Average Save %, Faceoff Win % |
| **Physicality** | `hits`, `blocked_shots`, `pim` | Hits/game, Blocked Shots/game, PIM/game |
| **Schedule Impact** | `rest_days` | Performance by Rest Days (0, 1, 2, 3+ days) |

---

## 🛠️ Developer Guide & Operations

### Installation
```bash
npm install
```

### Build & Bundle
Bundles TypeScript source files into a clean `dist/` build output optimized for Apps Script V8 execution:
```bash
npm run build
```

### Run Unit Tests
```bash
npm test
```

### Static Analysis & Linting
```bash
npm run lint
```

### Multi-Stage Deployment
Deploy seamlessly to **Development** or **Production** Apps Script instances:
```bash
# Deploys to Staging / Development Apps Script Project
npm run deploy

# Deploys directly to Production Apps Script Project
npm run deploy:prod
```
