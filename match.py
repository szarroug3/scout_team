from api import get_enemy_captain


LANE = {1: 'S', 2: 'M', 3: 'O', 4: 'J'}


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
    def __init__(self, data, team_id, players, heroes, api_key):
        Match.__init__(self, data, team_id, players)
        self.team_side_number = 0 if self.side == 'radiant' else 1
        self.first_pick = None
        self.bans = {}
        self.banned_against = {}
        self.enemy_captain = get_enemy_captain(self.match_id, self.side, players, api_key)
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

