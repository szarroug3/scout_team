from collections import defaultdict
from time import sleep

from match import ParsedMatch, UnparsedMatch
from api import dota_api_call, get_player_name, opendota_api_call


class Player(object):
    def __init__(self, account_id, player_names, heroes):
        self.account_id = account_id
        self.name = get_player_name(self.account_id, player_names)
        self.heroes = self.get_heroes(heroes)
        self.recent_heroes = self.get_recent_heroes(heroes)

    def get_heroes(self, heroes):
        data = opendota_api_call('players', self.account_id, 'heroes')

        hero_data = []
        for hero in sorted(data, key=lambda h: h['games'], reverse=True)[:5]:
            hero_id = hero['hero_id']
            win_rate = hero['win'] * 100 / hero['games']
            hero_data.append({'name': heroes[int(hero_id)], 'games': hero['games'],
                              'win_rate': '{:.1f}%'.format(win_rate)})
        return hero_data

    def get_recent_heroes(self, heroes):
        data = opendota_api_call('players', self.account_id, 'recentMatches')

        hero_data = defaultdict(lambda: defaultdict(int))
        for match in data:
            hero_id = match['hero_id']
            name = heroes[hero_id]
            hero_data[name]['count'] += 1
            if match['player_slot'] < 100 and not match['radiant_win']:
                hero_data[name]['wins'] += 1
            if match['player_slot'] > 100 and match['radiant_win']:
                hero_data[name]['wins'] += 1
        return hero_data


class Team(object):
    def __init__(self, team_id, player_names, heroes, league_id, api_key):
        if not league_id:
            raise Exception('league_id required for team scouting')

        self.team_id = team_id
        self.player_names = player_names
        self.heroes = heroes
        self.api_key = api_key

        self.parsed_matches = []
        self.unparsed_matches = []
        self.pick_count = defaultdict(lambda: defaultdict(int))
        self.ban_count = defaultdict(int)
        self.banned_against_count = defaultdict(int)
        self.players = {}

        self.get_team_data()
        self.parse_matches(self.get_team_matches(league_id=league_id))

    def get_team_data(self):
        data = dota_api_call('GetTeamInfoByTeamID', self.api_key, start_at_team_id=self.team_id, teams_requested=1)
        if not data['teams']:
            raise Exception('Team with ID {} not found.'.format(self.team_id))

        self.name = data['teams'][0]['name']
        for key, value in data['teams'][0].items():
            if not key.startswith('player') or not key.endswith('account_id'):
                continue
            value = str(value)
            self.players[value] = Player(value, self.player_names, self.heroes)
            break
            sleep(.5)

    def get_team_matches(self, league_id):
        matches = set()
        while True:
            if matches:
                matches_min = min(matches, key=lambda m: m[0])[0] - 1
            else:
                matches_min = None

            new_matches = self.get_matches(matches, league_id=league_id, start_at_match_id=matches_min)
            # TODO: remove this
            return new_matches
            if len(matches) == len(new_matches):
                return matches
            matches = new_matches
        return matches

    def get_matches(self, matches, league_id, start_at_match_id=None):
        params = {'matches_requested': 1000, 'league_id': league_id}
        if start_at_match_id:
            params['start_at_match_id'] = start_at_match_id

        data = dota_api_call('GetMatchHistory', self.api_key, **params)
        for match in data['matches']:
            if int(self.team_id) in [match['dire_team_id'], match['radiant_team_id']]:
                matches.add(match['match_id'])
        return matches

    def parse_matches(self, matches):
        for match in sorted(matches, reverse=True):
            data = opendota_api_call('matches', str(match))
            if data.get('picks_bans'):
                match_details = ParsedMatch(data, self.team_id, self.player_names, self.heroes, self.api_key)
                for hero in match_details.picks.values():
                    self.pick_count[hero['name']]['count'] += 1
                    if match_details.win:
                        self.pick_count[hero['name']]['wins'] += 1
                for hero in match_details.banned_against.values():
                    self.banned_against_count[hero['name']] += 1
                for hero in match_details.bans.values():
                    self.ban_count[hero['name']] += 1
                self.parsed_matches.append(match_details)
            else:
                match_details = UnparsedMatch(data, self.team_id, self.player_names, self.heroes, self.players.keys())
                self.unparsed_matches.append(match_details)
                for hero in match_details.picks.values():
                    self.pick_count[hero['name']]['count'] += 1
                    if match_details.win:
                        self.pick_count[hero['name']]['wins'] += 1


