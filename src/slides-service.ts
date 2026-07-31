/**
 * Copyright 2025 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import {TeamStats, HomeAwayStats, RestDayStats} from './types';

const COLORS = {
  darkBlue: '#1A264D',
  accentBlue: '#3366CC',
  lightBlue: '#E8F0FE',
  white: '#FFFFFF',
  lightGray: '#F5F6F8',
  cardGray: '#EAECEF',
  green: '#26A64D',
  red: '#CC2626',
  orange: '#F28C1A',
  black: '#1F1F1F',
};

function fmt(input: number, decimals = 1): string {
  return input.toFixed(decimals);
}

function fmtPct(input: number): string {
  return (input * 100).toFixed(1) + '%';
}

/** Creates a blank slide and returns the presentation, slide, and slide ID. */
function addSlide(presentation: GoogleAppsScript.Slides.Presentation) {
  const slide = presentation.appendSlide();
  return {presentation, slide, slideId: slide.getObjectId()};
}

/** Helper to add a text box. */
function addTextBox(
  slide: GoogleAppsScript.Slides.Slide,
  text: string,
  top: number,
  left: number,
  width: number,
  height: number,
  fontSize: number,
  bold: boolean,
  color = COLORS.black,
  alignment: GoogleAppsScript.Slides.ContentAlignment = SlidesApp
    .ContentAlignment.MIDDLE,
): GoogleAppsScript.Slides.Shape {
  const content = text || ' ';
  const shape = slide.insertTextBox(content, left, top, width, height);
  const textRange = shape.getText();
  const style = textRange.getTextStyle();
  style.setFontSize(fontSize);
  style.setBold(bold);
  style.setForegroundColor(color);
  const paragraphAlignment =
    alignment === SlidesApp.ContentAlignment.TOP
      ? SlidesApp.ParagraphAlignment.START
      : SlidesApp.ParagraphAlignment.CENTER;
  textRange.getParagraphStyle().setParagraphAlignment(paragraphAlignment);
  shape.setContentAlignment(alignment);
  return shape;
}

/** Fills a shape with a solid color. */
function setFill(
  element: GoogleAppsScript.Slides.Shape | GoogleAppsScript.Slides.TableCell,
  color: string,
): void {
  element.getFill().setSolidFill(color);
}

/** Helper to add a visual progress bar. */
function addProgressBar(
  slide: GoogleAppsScript.Slides.Slide,
  top: number,
  left: number,
  width: number,
  height: number,
  fillPct: number,
  color: string,
): void {
  const track = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    left,
    top,
    width,
    height,
  );
  track.getFill().setSolidFill('#E0E4EC');
  track.getBorder().setTransparent();

  const safePct = Math.min(Math.max(fillPct, 0), 1);
  if (safePct > 0) {
    const fillWidth = Math.max(width * safePct, 2);
    const fill = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      left,
      top,
      fillWidth,
      height,
    );
    fill.getFill().setSolidFill(color);
    fill.getBorder().setTransparent();
  }
}

/** Helper to add a styled visual metric card. */
function addMetricCard(
  slide: GoogleAppsScript.Slides.Slide,
  top: number,
  left: number,
  width: number,
  height: number,
  title: string,
  value: string,
  subtext: string,
  accentColor: string,
  progressPct?: number,
): void {
  // Card background
  const bg = slide.insertShape(
    SlidesApp.ShapeType.ROUND_RECTANGLE,
    left,
    top,
    width,
    height,
  );
  bg.getFill().setSolidFill(COLORS.lightGray);
  bg.getBorder().setTransparent();

  // Accent strip on left
  const bar = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    left,
    top,
    5,
    height,
  );
  bar.getFill().setSolidFill(accentColor);
  bar.getBorder().setTransparent();

  // Title
  addTextBox(
    slide,
    title,
    top + 4,
    left + 12,
    width - 18,
    16,
    10,
    true,
    COLORS.accentBlue,
    SlidesApp.ContentAlignment.TOP,
  );

  // Value
  addTextBox(
    slide,
    value,
    top + 20,
    left + 12,
    width - 18,
    22,
    15,
    true,
    COLORS.black,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  // Progress Bar if applicable
  if (progressPct !== undefined) {
    addProgressBar(
      slide,
      top + 44,
      left + 12,
      width - 24,
      5,
      progressPct,
      accentColor,
    );
  }

  // Subtext
  if (subtext) {
    addTextBox(
      slide,
      subtext,
      top + (progressPct !== undefined ? 52 : 42),
      left + 12,
      width - 18,
      16,
      9,
      false,
      COLORS.darkBlue,
      SlidesApp.ContentAlignment.MIDDLE,
    );
  }
}

/** Helper to manage temporary spreadsheet chart creation. */
function getOrCreateChartSheet(): GoogleAppsScript.Spreadsheet.Sheet | null {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    let sheet = ss.getSheetByName('_temp_charts');
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet('_temp_charts');
    return sheet;
  } catch (e) {
    console.warn('Chart sheet setup skipped:', e);
    return null;
  }
}

function cleanupChartSheet(
  chartSheet: GoogleAppsScript.Spreadsheet.Sheet | null,
): void {
  if (!chartSheet) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getSheetByName('_temp_charts')) {
      ss.deleteSheet(chartSheet);
    }
  } catch {
    // Ignore cleanup error
  }
}

/** Creates the title slide. */
function createTitleSlide(presentation: GoogleAppsScript.Slides.Presentation) {
  const {slide} = addSlide(presentation);

  // Main Background
  setFill(
    slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 405),
    COLORS.darkBlue,
  );

  // Visual Decorative Banner
  const decBar = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    50,
    95,
    620,
    4,
  );
  decBar.getFill().setSolidFill(COLORS.orange);
  decBar.getBorder().setTransparent();

  addTextBox(
    slide,
    'NHL TEAM PERFORMANCE REPORT',
    115,
    50,
    620,
    70,
    34,
    true,
    COLORS.white,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  addTextBox(
    slide,
    'Comprehensive Visual Analytics & Team Insights',
    195,
    50,
    620,
    35,
    16,
    false,
    COLORS.lightGray,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  const today = new Date();
  addTextBox(
    slide,
    `Generated on ${today.toLocaleDateString('en-US', {month: 'long', day: 'numeric', year: 'numeric'})}`,
    240,
    50,
    620,
    30,
    13,
    false,
    COLORS.lightGray,
    SlidesApp.ContentAlignment.MIDDLE,
  );
}

/** Creates a single team overview slide with visual KPI cards and split charts. */
function createTeamSlide(
  presentation: GoogleAppsScript.Slides.Presentation,
  stats: TeamStats,
  homeAway: HomeAwayStats | null,
  chartSheet: GoogleAppsScript.Spreadsheet.Sheet | null,
  rowOffset: number,
): number {
  const {slide} = addSlide(presentation);

  // Header bar
  const headerShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    0,
    0,
    720,
    55,
  );
  setFill(headerShape, COLORS.darkBlue);
  headerShape.getBorder().setTransparent();

  addTextBox(
    slide,
    `${stats.teamName} — Performance Dashboard`,
    8,
    15,
    690,
    40,
    22,
    true,
    COLORS.white,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  // Left Column: Visual KPI Cards
  const leftX = 20;
  const cardW = 330;
  const cardH = 72;
  const gapY = 82;
  let startY = 68;

  // KPI Card 1: Record & Win %
  const winColor =
    stats.winPct >= 0.6
      ? COLORS.green
      : stats.winPct >= 0.4
        ? COLORS.orange
        : COLORS.red;
  addMetricCard(
    slide,
    startY,
    leftX,
    cardW,
    cardH,
    'SEASON RECORD & WIN RATE',
    `${stats.wins}W - ${stats.losses}L  (${fmtPct(stats.winPct)})`,
    `Total Points: ${stats.points} pts | Games Played: ${stats.gamesPlayed}`,
    winColor,
    stats.winPct,
  );

  // KPI Card 2: Offense & Defense
  startY += gapY;
  const gfPerGame = stats.goalsFor / stats.gamesPlayed;
  const gaPerGame = stats.goalsAgainst / stats.gamesPlayed;
  const diffSign = stats.goalDiff >= 0 ? '+' : '';
  addMetricCard(
    slide,
    startY,
    leftX,
    cardW,
    cardH,
    'SCORING EFFICIENCY',
    `${fmt(gfPerGame)} GF/gm  vs  ${fmt(gaPerGame)} GA/gm`,
    `Goal Differential: ${diffSign}${stats.goalDiff} goals`,
    stats.goalDiff >= 0 ? COLORS.green : COLORS.red,
  );

  // KPI Card 3: Special Teams
  startY += gapY;
  addMetricCard(
    slide,
    startY,
    leftX,
    cardW,
    cardH,
    'SPECIAL TEAMS & GOALTENDING',
    `PP: ${fmtPct(stats.powerPlayPct / 100)}  |  PK: ${fmtPct(stats.penaltyKillPct / 100)}`,
    `Save Pct: ${fmtPct(stats.avgSavePct)} | Faceoff: ${fmt(stats.faceoffWinPct, 1)}%`,
    COLORS.accentBlue,
    stats.powerPlayPct / 100,
  );

  // KPI Card 4: Physicality & Discipline
  startY += gapY;
  addMetricCard(
    slide,
    startY,
    leftX,
    cardW,
    cardH,
    'SHOTS & PHYSICALITY / GAME',
    `Shots: ${fmt(stats.avgShotsFor)} For / ${fmt(stats.avgShotsAgainst)} Against`,
    `Hits: ${fmt(stats.avgHits)} | Blocks: ${fmt(stats.avgBlockedShots)} | PIM: ${fmt(stats.avgPim)}`,
    COLORS.orange,
  );

  // Right Column: Home/Away Chart or Fallback Table
  let nextRow = rowOffset;
  if (homeAway && chartSheet) {
    try {
      // Write chart data to temp sheet
      chartSheet
        .getRange(nextRow, 1, 1, 4)
        .setValues([['Location', 'Win %', 'GF/gm', 'GA/gm']]);
      chartSheet.getRange(nextRow + 1, 1, 2, 4).setValues([
        [
          'Home',
          homeAway.homeWinPct,
          homeAway.homeGoalsFor / (homeAway.homeGames || 1),
          homeAway.homeGoalsAgainst / (homeAway.homeGames || 1),
        ],
        [
          'Away',
          homeAway.awayWinPct,
          homeAway.awayGoalsFor / (homeAway.awayGames || 1),
          homeAway.awayGoalsAgainst / (homeAway.awayGames || 1),
        ],
      ]);

      const dataRange = chartSheet.getRange(nextRow, 1, 3, 4);
      const chart = chartSheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setOption('title', `${stats.teamName} — Home vs Away Split`)
        .setOption('colors', [COLORS.accentBlue, COLORS.green, COLORS.red])
        .setOption('legend', {position: 'bottom'})
        .build();

      chartSheet.insertChart(chart);
      slide.insertSheetsChartAsImage(chart, 365, 68, 335, 320);
      nextRow += 5;
    } catch (e) {
      console.warn('Home/Away chart render skipped:', e);
      renderHomeAwayTableFallback(slide, homeAway);
    }
  } else if (homeAway) {
    renderHomeAwayTableFallback(slide, homeAway);
  }

  return nextRow;
}

/** Fallback table for Home/Away split if charts are unavailable. */
function renderHomeAwayTableFallback(
  slide: GoogleAppsScript.Slides.Slide,
  homeAway: HomeAwayStats,
): void {
  addTextBox(
    slide,
    'Home / Away Split',
    75,
    380,
    320,
    24,
    14,
    true,
    COLORS.accentBlue,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  const haData = [
    [
      'Home',
      `${homeAway.homeWins}-${homeAway.homeGames - homeAway.homeWins}`,
      fmtPct(homeAway.homeWinPct),
      `${fmt(homeAway.homeGoalsFor / (homeAway.homeGames || 1))} GF/gm`,
      `${fmt(homeAway.homeGoalsAgainst / (homeAway.homeGames || 1))} GA/gm`,
    ],
    [
      'Away',
      `${homeAway.awayWins}-${homeAway.awayGames - homeAway.awayWins}`,
      fmtPct(homeAway.awayWinPct),
      `${fmt(homeAway.awayGoalsFor / (homeAway.awayGames || 1))} GF/gm`,
      `${fmt(homeAway.awayGoalsAgainst / (homeAway.awayGames || 1))} GA/gm`,
    ],
  ];

  const headers = ['Split', 'Record', 'Win%', 'GF/gm', 'GA/gm'];
  const colWidths = [70, 90, 80, 90, 90];
  let cx = 370;

  for (let c = 0; c < headers.length; c++) {
    addTextBox(
      slide,
      headers[c],
      110,
      cx,
      colWidths[c],
      20,
      9,
      true,
      COLORS.black,
      SlidesApp.ContentAlignment.MIDDLE,
    );
    cx += colWidths[c];
  }

  for (let r = 0; r < haData.length; r++) {
    cx = 370;
    for (let c = 0; c < haData[r].length; c++) {
      addTextBox(
        slide,
        haData[r][c],
        132 + r * 22,
        cx,
        colWidths[c],
        20,
        9,
        false,
        COLORS.black,
        SlidesApp.ContentAlignment.MIDDLE,
      );
      cx += colWidths[c];
    }
  }
}

/** Creates a rest-day analysis slide with visual chart. */
function createRestDaySlide(
  presentation: GoogleAppsScript.Slides.Presentation,
  teamName: string,
  restDayStats: RestDayStats[],
  chartSheet: GoogleAppsScript.Spreadsheet.Sheet | null,
  rowOffset: number,
): number {
  const {slide} = addSlide(presentation);

  // Header bar
  const headerShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    0,
    0,
    720,
    55,
  );
  setFill(headerShape, COLORS.darkBlue);
  headerShape.getBorder().setTransparent();

  addTextBox(
    slide,
    `${teamName} — Performance by Rest Days`,
    8,
    15,
    690,
    40,
    22,
    true,
    COLORS.white,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  if (restDayStats.length === 0) {
    addTextBox(
      slide,
      'No rest day data available.',
      100,
      40,
      600,
      40,
      14,
      false,
      COLORS.black,
      SlidesApp.ContentAlignment.MIDDLE,
    );
    return rowOffset;
  }

  // Left side: Rest Day Data Table
  const tableTop = 75;
  const tableLeft = 20;
  const cols = 6;
  const rows = restDayStats.length + 1;
  const colWidth = 56;
  const rowHeight = 24;

  const table = slide.insertTable(
    rows,
    cols,
    tableLeft,
    tableTop,
    cols * colWidth,
    rows * rowHeight,
  );

  const headers = ['Rest', 'GP', 'W', 'Win%', 'GF/gm', 'GA/gm'];
  for (let c = 0; c < cols; c++) {
    const cell = table.getCell(0, c);
    cell.getText().setText(headers[c] || ' ');
    cell.getText().getTextStyle().setBold(true).setFontSize(9);
    setFill(cell, COLORS.darkBlue);
    cell.getText().getTextStyle().setForegroundColor(COLORS.white);
  }

  for (let r = 0; r < restDayStats.length; r++) {
    const s = restDayStats[r];
    const values = [
      `${s.restDays}d`,
      s.games.toString(),
      s.wins.toString(),
      fmtPct(s.winPct),
      fmt(s.avgGoalsFor),
      fmt(s.avgGoalsAgainst),
    ];
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r + 1, c);
      cell.getText().setText(values[c] || ' ');
      cell.getText().getTextStyle().setFontSize(9);
      if (r % 2 === 0) {
        setFill(cell, COLORS.lightGray);
      }
    }
  }

  // Right side: Rest Day Visual Chart
  let nextRow = rowOffset;
  if (chartSheet) {
    try {
      const chartRows: (string | number)[][] = [
        ['Rest Days', 'Win Rate', 'Goals For/gm'],
      ];
      for (const s of restDayStats) {
        chartRows.push([`${s.restDays} Rest Day(s)`, s.winPct, s.avgGoalsFor]);
      }

      chartSheet.getRange(nextRow, 1, chartRows.length, 3).setValues(chartRows);

      const dataRange = chartSheet.getRange(nextRow, 1, chartRows.length, 3);
      const chart = chartSheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setOption('title', `${teamName} — Win Rate by Rest Days`)
        .setOption('colors', [COLORS.green, COLORS.accentBlue])
        .setOption('legend', {position: 'bottom'})
        .build();

      chartSheet.insertChart(chart);
      slide.insertSheetsChartAsImage(chart, 365, 75, 335, 310);
      nextRow += chartRows.length + 2;
    } catch (e) {
      console.warn('Rest day chart render skipped:', e);
    }
  }

  return nextRow;
}

/** Creates a summary / rankings slide comparing all teams with a visual chart. */
function createSummarySlide(
  presentation: GoogleAppsScript.Slides.Presentation,
  allStats: TeamStats[],
  chartSheet: GoogleAppsScript.Spreadsheet.Sheet | null,
  rowOffset: number,
): number {
  const {slide} = addSlide(presentation);

  const headerShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    0,
    0,
    720,
    55,
  );
  setFill(headerShape, COLORS.darkBlue);
  headerShape.getBorder().setTransparent();

  addTextBox(
    slide,
    'League Summary — Team Rankings & Comparison',
    8,
    15,
    690,
    40,
    22,
    true,
    COLORS.white,
    SlidesApp.ContentAlignment.MIDDLE,
  );

  const sorted = [...allStats].sort((a, b) => b.winPct - a.winPct);

  // Left side: Team Rankings Table
  const tableTop = 72;
  const tableLeft = 20;
  const cols = 7;
  const rows = sorted.length + 1;
  const colWidths = [38, 100, 35, 32, 32, 42, 55];
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  const rowHeight = 22;

  const table = slide.insertTable(
    rows,
    cols,
    tableLeft,
    tableTop,
    totalWidth,
    rows * rowHeight,
  );

  const headers = ['Rank', 'Team', 'GP', 'W', 'L', 'Pts', 'Win%'];
  for (let c = 0; c < cols; c++) {
    const cell = table.getCell(0, c);
    cell.getText().setText(headers[c] || ' ');
    cell.getText().getTextStyle().setBold(true).setFontSize(9);
    setFill(cell, COLORS.darkBlue);
    cell.getText().getTextStyle().setForegroundColor(COLORS.white);
  }

  for (let r = 0; r < sorted.length; r++) {
    const s = sorted[r];
    const values = [
      (r + 1).toString(),
      s.teamName,
      s.gamesPlayed.toString(),
      s.wins.toString(),
      s.losses.toString(),
      s.points.toString(),
      fmtPct(s.winPct),
    ];
    const winPctColor =
      s.winPct >= 0.6
        ? COLORS.green
        : s.winPct >= 0.4
          ? COLORS.orange
          : COLORS.red;

    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r + 1, c);
      cell.getText().setText(values[c] || ' ');
      cell.getText().getTextStyle().setFontSize(9);
      if (r % 2 === 0) {
        setFill(cell, COLORS.lightGray);
      }
      if (c === 6) {
        cell.getText().getTextStyle().setForegroundColor(winPctColor);
      }
    }
  }

  // Right side: Team Win % Visual Column Chart
  let nextRow = rowOffset;
  if (chartSheet) {
    try {
      const chartRows: (string | number)[][] = [['Team', 'Win Rate']];
      for (const s of sorted) {
        chartRows.push([s.teamName, s.winPct]);
      }

      chartSheet.getRange(nextRow, 1, chartRows.length, 2).setValues(chartRows);

      const dataRange = chartSheet.getRange(nextRow, 1, chartRows.length, 2);
      const chart = chartSheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setOption('title', 'Team Win % Comparison')
        .setOption('colors', [COLORS.accentBlue])
        .setOption('legend', {position: 'none'})
        .setOption('vAxis', {format: '0%'})
        .build();

      chartSheet.insertChart(chart);
      slide.insertSheetsChartAsImage(chart, 365, 72, 335, 315);
      nextRow += chartRows.length + 2;
    } catch (e) {
      console.warn('Summary chart render skipped:', e);
    }
  }

  return nextRow;
}

/** Main function: generates the full slide deck from the computed data. */
export function generateSlides(
  teamNames: string[],
  allStats: Map<string, TeamStats>,
  homeAwayMap: Map<string, HomeAwayStats>,
  restDayMap: Map<string, RestDayStats[]>,
): void {
  const presentationName =
    'NHL Report — ' +
    new Date().toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
    });

  const presentation = SlidesApp.create(presentationName);
  const chartSheet = getOrCreateChartSheet();
  let rowOffset = 1;

  try {
    // 1. Title slide
    createTitleSlide(presentation);

    // 2. League summary slide with visual comparison chart
    const statsArray: TeamStats[] = [];
    for (const name of teamNames) {
      const s = allStats.get(name);
      if (s) statsArray.push(s);
    }
    rowOffset = createSummarySlide(
      presentation,
      statsArray,
      chartSheet,
      rowOffset,
    );

    // 3. Per-team slides with KPI cards and embedded charts
    for (const teamName of teamNames) {
      const stats = allStats.get(teamName);
      if (!stats) continue;

      const ha = homeAwayMap.get(teamName) ?? null;
      rowOffset = createTeamSlide(
        presentation,
        stats,
        ha,
        chartSheet,
        rowOffset,
      );

      const rd = restDayMap.get(teamName) ?? [];
      if (rd.length > 0) {
        rowOffset = createRestDaySlide(
          presentation,
          teamName,
          rd,
          chartSheet,
          rowOffset,
        );
      }
    }
  } finally {
    cleanupChartSheet(chartSheet);
  }

  // Open the presentation
  const url = presentation.getUrl();
  const html = HtmlService.createHtmlOutput(
    `<p>Report generated! <a href="${url}" target="_blank">Open Slides</a></p>`,
  )
    .setWidth(300)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Report Generated');
}
