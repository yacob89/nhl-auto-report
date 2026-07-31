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
  computeTeamStats,
  computeHomeAwayStats,
  computeRestDayStats,
  getTeamNames,
} from '../src/data-service';
import {NhlGameData} from '../src/types';

function makeGame(overrides: Partial<NhlGameData>): NhlGameData {
  return {
    won: 1,
    game_id: 1,
    date: '2024-10-01',
    season: 2025,
    venue: 'Test Arena',
    attendance: 18000,
    officials: '',
    season_series: '',
    spread: 0,
    over_under: 0,
    favorite_moneyline: 0,
    team_id: 1,
    team_name: 'Boston Bruins',
    home_away: 'home',
    team_record: '2-1-0',
    score: 4,
    shots: 30,
    power_play_goals: 1,
    power_play_opportunities: 3,
    faceoff_win_pct: 50,
    hits: 20,
    blocked_shots: 10,
    pim: 6,
    giveaways: 10,
    takeaways: 5,
    cum_wins: 2,
    cum_games: 3,
    season_win_pct: 0.667,
    rest_days: 2,
    rolling_score_3: 3.0,
    rolling_power_play_goals_3: 1.0,
    rolling_power_play_opportunities_3: 3.0,
    rolling_faceoff_win_pct_3: 50.0,
    rolling_hits_3: 20.0,
    rolling_blocked_shots_3: 10.0,
    rolling_shots_3: 30.0,
    rolling_pim_3: 6.0,
    rolling_giveaways_3: 10.0,
    rolling_takeaways_3: 5.0,
    rolling_pp_efficiency_3: 0.333,
    rolling_score_10: 3.0,
    rolling_power_play_goals_10: 1.0,
    rolling_power_play_opportunities_10: 3.0,
    rolling_faceoff_win_pct_10: 50.0,
    rolling_hits_10: 20.0,
    rolling_blocked_shots_10: 10.0,
    rolling_shots_10: 30.0,
    rolling_pim_10: 6.0,
    rolling_giveaways_10: 10.0,
    rolling_takeaways_10: 5.0,
    rolling_pp_efficiency_10: 0.333,
    opp_team_id: 2,
    opp_team_name: 'Opponent',
    opp_won: 0,
    opp_rest_days: 2,
    opp_season_win_pct: 0.5,
    opp_score: 2,
    opp_shots: 28,
    opp_rolling_score_3: 2.0,
    opp_rolling_score_10: 2.0,
    opp_rolling_power_play_goals_3: 0.0,
    opp_rolling_power_play_goals_10: 0.0,
    opp_rolling_power_play_opportunities_3: 2.0,
    opp_rolling_power_play_opportunities_10: 2.0,
    opp_rolling_faceoff_win_pct_3: 50.0,
    opp_rolling_faceoff_win_pct_10: 50.0,
    opp_rolling_hits_3: 15.0,
    opp_rolling_hits_10: 15.0,
    opp_rolling_blocked_shots_3: 8.0,
    opp_rolling_blocked_shots_10: 8.0,
    opp_rolling_shots_3: 28.0,
    opp_rolling_shots_10: 28.0,
    opp_rolling_pim_3: 4.0,
    opp_rolling_pim_10: 4.0,
    opp_rolling_giveaways_3: 8.0,
    opp_rolling_giveaways_10: 8.0,
    opp_rolling_takeaways_3: 4.0,
    opp_rolling_takeaways_10: 4.0,
    opp_rolling_pp_efficiency_3: 0.0,
    opp_rolling_pp_efficiency_10: 0.0,
    save_pct: 0.929,
    is_home: 1,
    rest_advantage: 0,
    ...overrides,
  };
}

describe('data-service', () => {
  const games = [
    makeGame({won: 1, score: 4, opp_score: 2, is_home: 1, rest_days: 2}),
    makeGame({won: 0, score: 1, opp_score: 3, is_home: 0, rest_days: 2}),
    makeGame({
      won: 1,
      score: 3,
      opp_score: 1,
      is_home: 1,
      rest_days: 3,
      team_name: 'Montreal Canadiens',
      team_id: 10,
    }),
    makeGame({
      won: 0,
      score: 2,
      opp_score: 5,
      is_home: 0,
      rest_days: 3,
      team_name: 'Montreal Canadiens',
      team_id: 10,
    }),
  ];

  test('getTeamNames returns unique sorted team names', () => {
    const names = getTeamNames(games);
    expect(names).toEqual(['Boston Bruins', 'Montreal Canadiens']);
  });

  test('computeTeamStats computes correct stats', () => {
    const stats = computeTeamStats(games, 'Boston Bruins');
    expect(stats).toBeDefined();
    expect(stats!.teamName).toBe('Boston Bruins');
    expect(stats!.gamesPlayed).toBe(2);
    expect(stats!.wins).toBe(1);
    expect(stats!.losses).toBe(1);
    expect(stats!.points).toBe(2);
    expect(stats!.goalsFor).toBe(5);
    expect(stats!.goalsAgainst).toBe(5);
    expect(stats!.goalDiff).toBe(0);
    expect(stats!.winPct).toBe(0.5);
  });

  test('computeTeamStats returns null for missing team', () => {
    const stats = computeTeamStats(games, 'Nonexistent');
    expect(stats).toBeNull();
  });

  test('computeHomeAwayStats computes home/away split', () => {
    const ha = computeHomeAwayStats(games, 'Boston Bruins');
    expect(ha).toBeDefined();
    expect(ha!.homeGames).toBe(1);
    expect(ha!.awayGames).toBe(1);
    expect(ha!.homeWins).toBe(1);
    expect(ha!.awayWins).toBe(0);
  });

  test('computeRestDayStats groups by rest days', () => {
    const rd = computeRestDayStats(games, 'Boston Bruins');
    // Both games have rest_days=2
    expect(rd).toHaveLength(1);
    expect(rd[0].restDays).toBe(2);
    expect(rd[0].games).toBe(2);
    expect(rd[0].wins).toBe(1);
  });
});
