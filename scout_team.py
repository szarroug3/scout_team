from argparse import ArgumentParser
from datetime import datetime

from api import get_heroes
from dota import Player, Team
from xlsx import XlsxWriter




def get_args():
    parser = ArgumentParser()
    parser.add_argument('api_key', help='steam api key -- you can get a key from https://steamcommunity.com/dev/apikey')
    parser.add_argument('-l', '--league-id', help='(required for team scouting only) the league to use for teams -- you can get this from dotabuff')
    parser.add_argument('-t', '--team-id', action='append', default=[], help='(optional -- supports multiples) teams to scout -- you can get this from dotabuff')
    parser.add_argument('-p', '--player', action='append', default=[], help='(optional -- supports multiples) players to scout -- you can get this from dotabuff')
    parser.add_argument('-f', '--file', help='the file to save results to')
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    current_time = datetime.utcnow().strftime('%Y%m%d%H%M%S')
    file = args.file if args.file else '{}.xlsx'.format(current_time)
    player_names = {}
    players = {}
    heroes = get_heroes()

    if True:
    # try:
        writer = XlsxWriter(file)
        for team_id in args.team_id:
            team = Team(team_id, player_names, heroes, args.league_id, args.api_key)
            players.update(team.players)
            writer.write_matches(team)
            writer.write_summary(team)
        writer.write_legend()

        for player in args.player:
            if player not in players:
                players[player] = Player(player, player_names, heroes)

        writer.write_players(players.values())
        writer.close()
    # except Exception as e:
         #print(e)