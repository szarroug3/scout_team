import os

from argparse import ArgumentParser
from configparser import ConfigParser
from datetime import datetime

from api import get_heroes
from dota import Player, Team
from writer import XlsxWriter, HtmlWriter


def get_args():
    parser = ArgumentParser()
    parser.add_argument('-a', '--api-key', help='steam api key -- you can get a key from https://steamcommunity.com/dev/apikey')
    parser.add_argument('-l', '--league-id', help='(required for team scouting only) the league to use for teams -- you can get this from dotabuff')
    parser.add_argument('-t', '--team-id', action='append', default=[], help='(optional -- supports multiples) teams to scout -- you can get this from dotabuff')
    parser.add_argument('-p', '--player', action='append', default=[], help='(optional -- supports multiples) players to scout -- you can get this from dotabuff')
    parser.add_argument('-c', '--counterpick-heroes', help='(optional) list or file of heroes to highlight if found in recent matches; list should be comma separated')
    parser.add_argument('-f', '--file_format', help='the file format to use for output file name')
    parser.add_argument('-x', '--xlsx', action='store_true', help='write to an xlsx file instead of html')
    return parser.parse_args()


def read_config():
    config = ConfigParser()
    config.read('.config')
    main = config['MAIN']
    return main


if __name__ == '__main__':
    args = get_args()
    main = read_config()

    api_key = args.api_key or main.get('API_KEY')
    if not api_key:
        print('API key must be specified in config file or input parameters')
        exit()

    league_id = args.league_id or main.get('LEAGUE_ID')
    counterpick_heroes = args.counterpick_heroes or main.get('COUNTERPICK_HEROES')
    xlsx = args.xlsx or main.getboolean('XLSX')
    file_format = args.file_format or main.get('FILE_FORMAT') or '{{name}}_{{date}}.{format}'.format(format='xlsx' if xlsx else 'html')

    date = datetime.utcnow().strftime('%Y%m%d%H%M%S')
    file = file_format.format(name=args.team_id, date=date)
    player_names = {}
    players = {}
    heroes = get_heroes(api_key)
    highlight_heroes = []

    if counterpick_heroes:
        if os.path.exists(counterpick_heroes):
            with open(counterpick_heroes) as f:
                counterpick_heroes = f.read()

        lowered_hero_names = [h.lower() for h in heroes.values()]
        for hero in counterpick_heroes.split(','):
            filtered_hero = hero.strip().lower()
            if filtered_hero not in lowered_hero_names:
                print('Warning: {} not a valid hero name and will be ignored'.format(hero))
                continue
            highlight_heroes.append(filtered_hero)

    if True:
    # try:
        writer = XlsxWriter(file) if xlsx else HtmlWriter(file)
        for team_id in args.team_id:
            team = Team(team_id, player_names, heroes, league_id, api_key)
            players.update(team.players)
            writer.write_matches(team)
            writer.write_summary(team)
        writer.write_legend()

        for player in args.player:
            break
            if player not in players:
                players[player] = Player(player, player_names, heroes)

        writer.write_players(players.values(), highlight_heroes)
        writer.close()
    # except Exception as e:
         #print(e)
