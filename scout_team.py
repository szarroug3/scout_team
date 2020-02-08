import requests
import xlsxwriter

from argparse import ArgumentParser
from collections import defaultdict
from datetime import datetime
from itertools import zip_longest
from time import sleep
from urllib.parse import urlencode


DOTA_BASE_URL = 'https://api.steampowered.com/IDOTA2Match_570/{{func_name}}/V1/?key={api_key}&{{params}}'
OPENDOTA_BASE_URL = 'https://api.opendota.com/api/{func_name}/{params}'
LANE = {1: 'S', 2: 'M', 3: 'O', 4: 'J'}
COLORS = ['white', 'yellow', 'orange', 'cyan', 'silver', '#FF7F50', '#FFD700', '#ADFF2F',
          '#40E0D0', '#00BFFF', '#D8BFD8', '#FFC0CB', '#FAEBD7', '#E6E6FA', '#FFD700']


def dota_api_call(func_name, **params):
    resp = requests.get(DOTA_BASE_URL.format(func_name=func_name, params=urlencode(params)))

    if not resp.ok:
        raise Exception('Something went wrong: GET {}: {} {}'.format(func_name, resp.status_code, resp.reason))
    return resp.json().get('result', {})


def opendota_api_call(func_name, *params):
    resp = requests.get(OPENDOTA_BASE_URL.format(func_name=func_name, params='/'.join(params)))

    if not resp.ok:
        raise Exception('Something went wrong: GET {}: {} {}'.format(func_name, resp.status_code, resp.reason))

    return resp.json()


def get_heroes():
    data = opendota_api_call('heroes')
    return {hero['id']: hero['localized_name'] for hero in data}


def get_enemy_captain(match_id, side, player_names):
    data = dota_api_call('GetMatchDetails', match_id=match_id)
    captain_id = data.get('dire_captain') if side == 'radiant' else data.get('radiant_captain')
    if captain_id:
        return get_player_name(captain_id, player_names)
    return ''


def get_player_names(match, picks, player_names):
    for player in match['players']:
        hero_id = player['hero_id']
        account_id = player['account_id']
        if hero_id not in picks:
            continue
        picks[hero_id]['player_name'] = get_player_name(account_id, player_names)

    return picks, player_names


def get_player_name(account_id, player_names):
    account_id = str(account_id)
    if account_id not in player_names:
        data = opendota_api_call('players', account_id)
        player_names[account_id] = data['profile']['personaname']
    return player_names[account_id]


class Match(object):
    def __init__(self, data, team_id, players):
        self.data = data
        self.team_id = team_id
        self.match_id = data['match_id']
        self.side = self.get_side(team_id, str(data['dire_team_id']))
        self.win = self.get_win(data['radiant_win'])
        self.picks = {}

    def get_side(self, team_id, dire_team_id):
        if dire_team_id == team_id:
            return 'dire'
        return 'radiant'

    def get_win(self, radiant_win):
        if self.side == 'radiant' and radiant_win:
            return True
        if self.side == 'dire' and not radiant_win:
            return True
        return False


class ParsedMatch(Match):
    def __init__(self, data, team_id, players, heroes):
        Match.__init__(self, data, team_id, players)
        self.team_side_number = 0 if self.side == 'radiant' else 1
        self.first_pick = None
        self.bans = {}
        self.banned_against = {}
        self.enemy_captain = get_enemy_captain(self.match_id, self.side, players)
        self.get_picks_bans(heroes)
        self.get_player_info()

    def get_picks_bans(self, heroes):
        for picks_bans in self.data['picks_bans']:
            # enemy pick, don't care
            if picks_bans['team'] != self.team_side_number and picks_bans['is_pick']:
                continue

            if self.first_pick is None:
                if picks_bans['team'] == self.team_side_number and picks_bans['is_pick']:
                    self.first_pick = picks_bans['order'] == 6

            hero_id = picks_bans['hero_id']
            hero_name = heroes[hero_id]
            if picks_bans['is_pick']:
                self.picks[hero_id] = {'name': hero_name, 'order': picks_bans['order']}
            else:
                if picks_bans['team'] == self.team_side_number:
                    self.bans[hero_id] = {'name': hero_name, 'order': picks_bans['order']}
                else:
                    self.banned_against[hero_id] = {'name': hero_name, 'order': picks_bans['order']}

    def get_player_info(self):
        for player in self.data['players']:
            hero_id = player['hero_id']
            if hero_id not in self.picks:
                continue
            self.picks[hero_id].update({'lane': LANE[player['lane_role']],
                                        'roaming': player['is_roaming']})


class UnparsedMatch(Match):
    def __init__(self, data, team_id, players, heroes, account_ids):
        Match.__init__(self, data, team_id, players)
        self.get_player_info(heroes, account_ids)

    def get_player_info(self, heroes, account_ids):
        for player in self.data['players']:
            hero_id = player['hero_id']
            if str(player['account_id']) not in account_ids:
                continue
            self.picks[hero_id] = ({'name': heroes[hero_id],
                                    'lane': LANE[player['lane_role']],
                                    'roaming': player['is_roaming']})


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
            if match['player_slot'] > 100 and not match['radiant_win']:
                hero_data[name]['wins'] += 1
            if match['player_slot'] < 100 and match['radiant_win']:
                hero_data[name]['wins'] += 1
        return hero_data


class Team(object):
    def __init__(self, team_id, player_names, heroes, league_id):
        if not league_id:
            raise Exception('league_id required for team scouting')

        self.team_id = team_id
        self.player_names = player_names
        self.heroes = heroes
        self.parsed_matches = []
        self.unparsed_matches = []
        self.pick_count = defaultdict(lambda: defaultdict(int))
        self.ban_count = defaultdict(int)
        self.banned_against_count = defaultdict(int)
        self.players = []

        self.get_team_data()
        self.parse_matches(self.get_team_matches(league_id=league_id))

    def get_team_data(self):
        data = dota_api_call('GetTeamInfoByTeamID', start_at_team_id=self.team_id, teams_requested=1)
        if not data['teams']:
            raise Exception('Team with ID {} not found.'.format(self.team_id))

        self.name = data['teams'][0]['name']
        for key, value in data['teams'][0].items():
            if not key.startswith('player') or not key.endswith('account_id'):
                continue
            value = str(value)
            self.players.append(value)
            sleep(.5)

    def get_team_matches(self, league_id):
        matches = set()
        while True:
            if matches:
                matches_min = min(matches, key=lambda m: m[0])[0] - 1
            else:
                matches_min = None

            new_matches = self.get_matches(matches, league_id=league_id, start_at_match_id=matches_min)
            if len(matches) == len(new_matches):
                return matches
            matches = new_matches
        return matches

    def get_matches(self, matches, league_id, start_at_match_id=None):
        params = {'matches_requested': 1000, 'league_id': league_id}
        if start_at_match_id:
            params['start_at_match_id'] = start_at_match_id

        data = dota_api_call('GetMatchHistory', **params)
        for match in data['matches']:
            if int(self.team_id) in [match['dire_team_id'], match['radiant_team_id']]:
                matches.add(match['match_id'])
        return matches

    def parse_matches(self, matches):
        for match in sorted(matches, reverse=True):
            try:
                data = opendota_api_call('matches', str(match))
            except:
                print('match {} could not be found using dotabuff api'.format(match))
            if data.get('picks_bans'):
                match_details = ParsedMatch(data, self.team_id, self.player_names, self.heroes)
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
                match_details = UnparsedMatch(data, self.team_id, self.player_names, self.heroes, self.players)
                self.unparsed_matches.append(match_details)
                for hero in match_details.picks.values():
                    self.pick_count[hero['name']]['count'] += 1
                    if match_details.win:
                        self.pick_count[hero['name']]['wins'] += 1


class XlsxWriter(object):
    def __init__(self, file):
        self.workbook = xlsxwriter.Workbook(file)
        self.worksheet = self.workbook.add_worksheet()
        self.colors = self.create_colors()
        self.used_colors = set()
        self.row = 0

    def close(self):
        self.workbook.close()

    def create_colors(self):
        colors = {}
        for num, color in enumerate(COLORS, start=1):
            colors[num] = self.workbook.add_format({'bg_color': color})
        return colors

    def write_matches(self, team):
        if not team.parsed_matches and not team.unparsed_matches:
            return

        if team.parsed_matches:
            self.worksheet.write(self.row, 0, team.name)
            self.row += 1
            self.worksheet.write(self.row, 0, 'PICKS')
            self.worksheet.write(self.row, 5, 'PICK')
            self.worksheet.write(self.row, 6, 'RESULT')
            self.worksheet.write(self.row, 7, 'SIDE')
            self.worksheet.write(self.row, 8, 'CAPTAIN')
            self.worksheet.write(self.row, 9, 'BANNED AGAINST')
            self.worksheet.write(self.row, 16, 'BANS')
            self.worksheet.write(self.row, 23, 'DOTABUFF')
            self.row += 1
            self.write_parsed_matches(team)

        if team.unparsed_matches:
            self.worksheet.write(self.row, 0, 'PICKS')
            self.worksheet.write(self.row, 5, 'RESULT')
            self.worksheet.write(self.row, 6, 'SIDE')
            self.worksheet.write(self.row, 7, 'DOTABUFF')
            self.row += 1
            self.write_unparsed_matches(team)

    def write_parsed_matches(self, team):
        for match in team.parsed_matches:
            self.write_parsed_match(match, team)
            self.row += 1

        self.row += 1

    def write_parsed_match(self, match, team):
        for column, pick in enumerate(sorted(match.picks.values(), key=lambda p: p['order'])):
            count = team.pick_count[pick['name']]['count']
            color = self.colors.get(count, self.colors[1])
            self.used_colors.add(count)
            self.write_hero(column, color, **pick)

        self.worksheet.write(self.row, 5, 'FP' if match.first_pick else 'SP')
        self.worksheet.write(self.row, 6, 'W' if match.win else 'L')
        self.worksheet.write(self.row, 7, match.side)
        self.worksheet.write(self.row, 8, match.enemy_captain)

        for column, ban in enumerate(sorted(match.banned_against.values(), key=lambda p: p['order']), start=9):
            color = self.colors.get(team.banned_against_count[ban['name']], self.colors[1])
            self.used_colors.add(team.banned_against_count[ban['name']])
            self.write_hero(column, color, **ban)

        for column, ban in enumerate(sorted(match.bans.values(), key=lambda p: p['order']), start=16):
            color = self.colors.get(team.ban_count[ban['name']], self.colors[1])
            self.used_colors.add(team.ban_count[ban['name']])
            self.write_hero(column, color, **ban)

        self.worksheet.write(self.row, 23, 'http://www.dotabuff.com/matches/{}'.format(match.match_id))

    def write_unparsed_matches(self, team):
        for match in team.unparsed_matches:
            self.write_unparsed_match(match, team)
            self.row += 1

        self.row += 1

    def write_unparsed_match(self, match, team):
        for column, pick in enumerate(sorted(match.picks.values(), key=lambda p: p['lane'])):
            count = team.pick_count[pick['name']]['count']
            color = self.colors.get(count, self.colors[1])
            self.used_colors.add(count)
            self.write_hero(column, color, **pick)

        self.worksheet.write(self.row, 5, 'W' if match.win else 'L')
        self.worksheet.write(self.row, 6, match.side)
        self.worksheet.write(self.row, 7, 'http://www.dotabuff.com/matches/{}'.format(match.match_id))

    def write_summary(self, team):
        self.worksheet.write(self.row, 0, 'SUMMARY')
        self.row += 1

        self.worksheet.write(self.row, 0, 'PICK')
        self.worksheet.write(self.row, 1, 'COUNT')
        self.worksheet.write(self.row, 2, 'WINS')
        self.worksheet.write(self.row, 3, 'WIN RATE')

        self.worksheet.write(self.row, 5, 'BANNED AGAINST')
        self.worksheet.write(self.row, 6, 'COUNT')

        self.worksheet.write(self.row, 8, 'BAN')
        self.worksheet.write(self.row, 9, 'COUNT')
        self.row += 1

        pick_row, banned_against_row, ban_row = self.row, self.row, self.row
        pick_count = sorted(team.pick_count.items(), key=lambda p: p[1]['count'], reverse=True)
        banned_against_count = sorted(team.banned_against_count.items(), key=lambda b: b[1], reverse=True)
        ban_count = sorted(team.ban_count.items(), key=lambda b: b[1], reverse=True)

        for pick, banned_against, ban in zip_longest(pick_count, banned_against_count, ban_count):
            if pick:
                win_rate = '{:.1f}%'.format(pick[1]['wins'] * 100 / pick[1]['count'])
                if pick[1]['wins'] == pick[1]['count']:
                    self.worksheet.write(pick_row, 0, pick[0], self.colors[2])
                    self.worksheet.write(pick_row, 1, pick[1]['count'], self.colors[2])
                    self.worksheet.write(pick_row, 2, pick[1]['wins'], self.colors[2])
                    self.worksheet.write(pick_row, 3, win_rate, self.colors[2])
                else:
                    self.worksheet.write(pick_row, 0, pick[0])
                    self.worksheet.write(pick_row, 1, pick[1]['count'])
                    self.worksheet.write(pick_row, 2, pick[1]['wins'])
                    self.worksheet.write(pick_row, 3, win_rate)
                pick_row += 1
            if banned_against:
                self.worksheet.write(banned_against_row, 5, banned_against[0])
                self.worksheet.write(banned_against_row, 6, banned_against[1])
                banned_against_row += 1
            if ban:
                self.worksheet.write(ban_row, 8, ban[0])
                self.worksheet.write(ban_row, 9, ban[1])
                ban_row += 1

        self.row = max(pick_row, banned_against_row, ban_row)
        self.row += 1


    def write_hero(self, column, color, name=None, lane=None, roaming=False, **kwargs):
        data = name
        if lane:
            data += ' ' + lane

        if roaming:
            data += ' (R)'

        self.worksheet.write(self.row, column, data, color)

    def write_legend(self):
        if not self.used_colors:
            return

        self.worksheet.write(self.row, 0, 'LEGEND')
        self.row += 1

        for column, count in enumerate(sorted(self.used_colors)):
            color = self.colors.get(count, self.colors[1])
            self.worksheet.write(self.row, column, count, color)
        self.row += 2

    def write_players(self, players, highlight_heroes):
        self.worksheet.write(self.row, 0, 'PLAYERS')
        self.row += 1

        for player in sorted(players, key=lambda p: p.name.lower()):
            self.worksheet.write(self.row, 0, player.name)
            self.row += 1

            self.worksheet.write(self.row, 0, 'HERO')
            self.worksheet.write(self.row, 1, 'GAMES')
            self.worksheet.write(self.row, 2, 'WIN RATE')
            self.row += 1

            for hero in player.heroes:
                if hero['name'].lower() in highlight_heroes:
                    self.worksheet.write(self.row, 0, hero['name'], self.colors[2])
                    self.worksheet.write(self.row, 1, hero['games'], self.colors[2])
                    self.worksheet.write(self.row, 2, hero['win_rate'], self.colors[2])
                else:
                    self.worksheet.write(self.row, 0, hero['name'])
                    self.worksheet.write(self.row, 1, hero['games'])
                    self.worksheet.write(self.row, 2, hero['win_rate'])
                self.row += 1
            self.row += 1

            self.worksheet.write(self.row, 0, 'HERO')
            self.worksheet.write(self.row, 1, 'RECENT GAMES')
            self.worksheet.write(self.row, 2, 'RECENT WIN RATE')
            self.row += 1

            for hero, data in sorted(player.recent_heroes.items(), key=lambda h: h[1]['count'], reverse=True):
                win_rate = '{:.1f}%'.format(data['wins'] * 100 / data['count'])
                if hero.lower() in highlight_heroes:
                    self.worksheet.write(self.row, 0, hero, self.colors[2])
                    self.worksheet.write(self.row, 1, data['count'], self.colors[2])
                    self.worksheet.write(self.row, 2, win_rate, self.colors[2])
                else:
                    self.worksheet.write(self.row, 0, hero)
                    self.worksheet.write(self.row, 1, data['count'])
                    self.worksheet.write(self.row, 2, win_rate)
                self.row += 1
            self.row += 1
                


        self.row += 1


def get_args():
    parser = ArgumentParser()
    parser.add_argument('api_key', help='steam api key -- you can get a key from https://steamcommunity.com/dev/apikey')
    parser.add_argument('-l', '--league-id', help='(required for team scouting only) the league to use for teams -- you can get this from dotabuff')
    parser.add_argument('-t', '--team-id', action='append', default=[], help='(optional -- supports multiples) teams to scout -- you can get this from dotabuff')
    parser.add_argument('-p', '--player', action='append', default=[], help='(optional -- supports multiples) players to scout -- you can get this from dotabuff')
    parser.add_argument('-c', '--counterpick-heroes', help='(optional) file of heroes to highlight if found in recent matches; file should be comma separated')
    parser.add_argument('-f', '--file', help='the file to save results to')
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    current_time = datetime.utcnow().strftime('%Y%m%d%H%M%S')
    file = args.file if args.file else '{}.xlsx'.format(current_time)
    DOTA_BASE_URL = DOTA_BASE_URL.format(api_key=args.api_key)
    player_names = {}
    players = {}
    player_ids = set(args.player)
    heroes = get_heroes()
    highlight_heroes = []

    if args.counterpick_heroes:
        with open(args.counterpick_heroes) as f:
            lowered_hero_names = [h.lower() for h in heroes.values()]
            for hero in f.read().split(','):
                filtered_hero = hero.strip().lower()
                if filtered_hero not in lowered_hero_names:
                    print('Warning: {} not a valid hero name and will be ignored'.format(hero))
                    continue
                highlight_heroes.append(filtered_hero)

    try:
        writer = XlsxWriter(file)
        for team_id in args.team_id:
            team = Team(team_id, player_names, heroes, league_id=args.league_id)
            player_ids.update(team.players)
            writer.write_matches(team)
            writer.write_summary(team)
        writer.write_legend()

        for player in player_ids:
            players[player] = Player(player, player_names, heroes)

        writer.write_players(players.values(), highlight_heroes)
        writer.close()
    except Exception as e:
        print(e)