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

import {
  readGameData,
  computeTeamStats,
  computeHomeAwayStats,
  computeRestDayStats,
  getTeamNames,
} from './data-service';
import {TeamStats, HomeAwayStats, RestDayStats} from './types';
import {generateSlides} from './slides-service';

/** Creates the "Generate Report" custom menu when the sheet opens. */
export function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Generate Report')
    .addItem('Generate Slides Report', 'generateReport')
    .addToUi();
}

/** Entry point: reads the active sheet, computes stats, and generates slides. */
export function generateReport(): void {
  const ui = SpreadsheetApp.getUi();

  try {
    const games = readGameData();
    const teamNames = getTeamNames(games);

    if (teamNames.length === 0) {
      ui.alert(
        'No Data',
        'No team data found in the active sheet. Make sure the CSV data is imported.',
        ui.ButtonSet.OK,
      );
      return;
    }

    const allStats = new Map<string, TeamStats>();
    const homeAwayMap = new Map<string, HomeAwayStats>();
    const restDayMap = new Map<string, RestDayStats[]>();

    for (const teamName of teamNames) {
      const stats = computeTeamStats(games, teamName);
      if (stats) allStats.set(teamName, stats);

      const ha = computeHomeAwayStats(games, teamName);
      if (ha) homeAwayMap.set(teamName, ha);

      const rd = computeRestDayStats(games, teamName);
      restDayMap.set(teamName, rd);
    }

    generateSlides(teamNames, allStats, homeAwayMap, restDayMap);
  } catch (e) {
    ui.alert(
      'Error',
      `Failed to generate report: ${(e as Error).message}`,
      ui.ButtonSet.OK,
    );
  }
}
