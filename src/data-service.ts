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

import {NhlGameData, TeamStats, HomeAwayStats, RestDayStats} from './types';

/** Reads game data from the active sheet. */
export function readGameData(): NhlGameData[] {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error(
      'Sheet has no data rows. Ensure the CSV data is imported with a header row.',
    );
  }

  const headers = data[0] as string[];
  const rows = data.slice(1) as unknown[][];

  return rows.map(row => {
    const obj: Record<string, unknown> = {};
    headers.forEach((header, idx) => {
      obj[header] = row[idx];
    });
    return obj as unknown as NhlGameData;
  });
}

/** Computes TeamStats for a given set of game rows. */
export function computeTeamStats(
  games: NhlGameData[],
  teamName: string,
): TeamStats | null {
  const teamGames = games.filter(g => g.team_name === teamName);
  if (teamGames.length === 0) return null;

  const gamesPlayed = teamGames.length;
  const wins = teamGames.filter(g => g.won === 1).length;
  const losses = gamesPlayed - wins;
  const points = wins * 2;

  const goalsFor = teamGames.reduce((sum, g) => sum + g.score, 0);
  const goalsAgainst = teamGames.reduce((sum, g) => sum + g.opp_score, 0);

  const totalShots = teamGames.reduce((sum, g) => sum + g.shots, 0);
  const totalOppShots = teamGames.reduce((sum, g) => sum + g.opp_shots, 0);

  const totalPpGoals = teamGames.reduce(
    (sum, g) => sum + g.power_play_goals,
    0,
  );
  const totalPpOpps = teamGames.reduce(
    (sum, g) => sum + g.power_play_opportunities,
    0,
  );

  const totalPkOpps = teamGames.reduce(
    (sum, g) => sum + g.opp_rolling_power_play_opportunities_3,
    0,
  );
  // Opponent PP goals ≈ opp_score minus even-strength — but we lack direct PK
  // data so we approximate via save percentage context.
  // Use the opp's rolling PP efficiency to estimate PK%.

  const totalFaceoff = teamGames.reduce((sum, g) => sum + g.faceoff_win_pct, 0);
  const totalSavePct = teamGames.reduce((sum, g) => sum + g.save_pct, 0);
  const totalHits = teamGames.reduce((sum, g) => sum + g.hits, 0);
  const totalBlocks = teamGames.reduce((sum, g) => sum + g.blocked_shots, 0);
  const totalPim = teamGames.reduce((sum, g) => sum + g.pim, 0);

  return {
    teamName,
    gamesPlayed,
    wins,
    losses,
    points,
    winPct: gamesPlayed > 0 ? wins / gamesPlayed : 0,
    goalsFor,
    goalsAgainst,
    goalDiff: goalsFor - goalsAgainst,
    avgShotsFor: gamesPlayed > 0 ? totalShots / gamesPlayed : 0,
    avgShotsAgainst: gamesPlayed > 0 ? totalOppShots / gamesPlayed : 0,
    shotDiff: gamesPlayed > 0 ? (totalShots - totalOppShots) / gamesPlayed : 0,
    powerPlayPct: totalPpOpps > 0 ? (totalPpGoals / totalPpOpps) * 100 : 0,
    penaltyKillPct:
      totalPkOpps > 0
        ? 100 -
          (teamGames.reduce(
            (sum, g) => sum + g.opp_rolling_power_play_goals_3,
            0,
          ) /
            totalPkOpps) *
            100
        : 0,
    avgSavePct: gamesPlayed > 0 ? totalSavePct / gamesPlayed : 0,
    faceoffWinPct: gamesPlayed > 0 ? totalFaceoff / gamesPlayed : 0,
    avgHits: gamesPlayed > 0 ? totalHits / gamesPlayed : 0,
    avgBlockedShots: gamesPlayed > 0 ? totalBlocks / gamesPlayed : 0,
    avgPim: gamesPlayed > 0 ? totalPim / gamesPlayed : 0,
  };
}

/** Computes home/away split stats. */
export function computeHomeAwayStats(
  games: NhlGameData[],
  teamName: string,
): HomeAwayStats | null {
  const teamGames = games.filter(g => g.team_name === teamName);
  if (teamGames.length === 0) return null;

  const homeGames = teamGames.filter(g => g.is_home === 1);
  const awayGames = teamGames.filter(g => g.is_home === 0);

  const homeWins = homeGames.filter(g => g.won === 1).length;
  const awayWins = awayGames.filter(g => g.won === 1).length;

  return {
    homeGames: homeGames.length,
    homeWins,
    homeWinPct: homeGames.length > 0 ? homeWins / homeGames.length : 0,
    homeGoalsFor: homeGames.reduce((s, g) => s + g.score, 0),
    homeGoalsAgainst: homeGames.reduce((s, g) => s + g.opp_score, 0),
    awayGames: awayGames.length,
    awayWins,
    awayWinPct: awayGames.length > 0 ? awayWins / awayGames.length : 0,
    awayGoalsFor: awayGames.reduce((s, g) => s + g.score, 0),
    awayGoalsAgainst: awayGames.reduce((s, g) => s + g.opp_score, 0),
  };
}

/** Computes stats grouped by rest days. */
export function computeRestDayStats(
  games: NhlGameData[],
  teamName: string,
): RestDayStats[] {
  const teamGames = games.filter(g => g.team_name === teamName);
  const map = new Map<
    number,
    {games: number; wins: number; gf: number; ga: number}
  >();

  for (const g of teamGames) {
    const entry = map.get(g.rest_days) ?? {
      games: 0,
      wins: 0,
      gf: 0,
      ga: 0,
    };
    entry.games++;
    entry.wins += g.won;
    entry.gf += g.score;
    entry.ga += g.opp_score;
    map.set(g.rest_days, entry);
  }

  const result: RestDayStats[] = [];
  for (const [restDays, data] of map) {
    result.push({
      restDays,
      games: data.games,
      wins: data.wins,
      winPct: data.games > 0 ? data.wins / data.games : 0,
      avgGoalsFor: data.games > 0 ? data.gf / data.games : 0,
      avgGoalsAgainst: data.games > 0 ? data.ga / data.games : 0,
    });
  }
  result.sort((a, b) => a.restDays - b.restDays);
  return result;
}

/** Returns unique team names from the data. */
export function getTeamNames(games: NhlGameData[]): string[] {
  const names = new Set<string>();
  for (const g of games) {
    names.add(g.team_name);
  }
  return [...names].sort();
}
