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

/** Represents a single row of NHL game data from the sheet. */
export interface NhlGameData {
  won: number;
  game_id: number;
  date: string;
  season: number;
  venue: string;
  attendance: number;
  officials: string;
  season_series: string;
  spread: number;
  over_under: number;
  favorite_moneyline: number;
  team_id: number;
  team_name: string;
  home_away: string;
  team_record: string;
  score: number;
  shots: number;
  power_play_goals: number;
  power_play_opportunities: number;
  faceoff_win_pct: number;
  hits: number;
  blocked_shots: number;
  pim: number;
  giveaways: number;
  takeaways: number;
  cum_wins: number;
  cum_games: number;
  season_win_pct: number;
  rest_days: number;
  rolling_score_3: number;
  rolling_power_play_goals_3: number;
  rolling_power_play_opportunities_3: number;
  rolling_faceoff_win_pct_3: number;
  rolling_hits_3: number;
  rolling_blocked_shots_3: number;
  rolling_shots_3: number;
  rolling_pim_3: number;
  rolling_giveaways_3: number;
  rolling_takeaways_3: number;
  rolling_pp_efficiency_3: number;
  rolling_score_10: number;
  rolling_power_play_goals_10: number;
  rolling_power_play_opportunities_10: number;
  rolling_faceoff_win_pct_10: number;
  rolling_hits_10: number;
  rolling_blocked_shots_10: number;
  rolling_shots_10: number;
  rolling_pim_10: number;
  rolling_giveaways_10: number;
  rolling_takeaways_10: number;
  rolling_pp_efficiency_10: number;
  opp_team_id: number;
  opp_team_name: string;
  opp_won: number;
  opp_rest_days: number;
  opp_season_win_pct: number;
  opp_score: number;
  opp_shots: number;
  opp_rolling_score_3: number;
  opp_rolling_score_10: number;
  opp_rolling_power_play_goals_3: number;
  opp_rolling_power_play_goals_10: number;
  opp_rolling_power_play_opportunities_3: number;
  opp_rolling_power_play_opportunities_10: number;
  opp_rolling_faceoff_win_pct_3: number;
  opp_rolling_faceoff_win_pct_10: number;
  opp_rolling_hits_3: number;
  opp_rolling_hits_10: number;
  opp_rolling_blocked_shots_3: number;
  opp_rolling_blocked_shots_10: number;
  opp_rolling_shots_3: number;
  opp_rolling_shots_10: number;
  opp_rolling_pim_3: number;
  opp_rolling_pim_10: number;
  opp_rolling_giveaways_3: number;
  opp_rolling_giveaways_10: number;
  opp_rolling_takeaways_3: number;
  opp_rolling_takeaways_10: number;
  opp_rolling_pp_efficiency_3: number;
  opp_rolling_pp_efficiency_10: number;
  save_pct: number;
  is_home: number;
  rest_advantage: number;
}

/** Aggregated team stats for the analysis report. */
export interface TeamStats {
  teamName: string;
  gamesPlayed: number;
  wins: number;
  losses: number;
  points: number;
  winPct: number;
  goalsFor: number;
  goalsAgainst: number;
  goalDiff: number;
  avgShotsFor: number;
  avgShotsAgainst: number;
  shotDiff: number;
  powerPlayPct: number;
  penaltyKillPct: number;
  avgSavePct: number;
  faceoffWinPct: number;
  avgHits: number;
  avgBlockedShots: number;
  avgPim: number;
}

/** Home / Away split stats. */
export interface HomeAwayStats {
  homeGames: number;
  homeWins: number;
  homeWinPct: number;
  homeGoalsFor: number;
  homeGoalsAgainst: number;
  awayGames: number;
  awayWins: number;
  awayWinPct: number;
  awayGoalsFor: number;
  awayGoalsAgainst: number;
}

/** Rest day analysis. */
export interface RestDayStats {
  restDays: number;
  games: number;
  wins: number;
  winPct: number;
  avgGoalsFor: number;
  avgGoalsAgainst: number;
}
