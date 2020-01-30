import requests

from urllib.parse import urlencode


DOTA_BASE_URL = 'https://api.steampowered.com/IDOTA2Match_570/{func_name}/V1/?key={api_key}&{params}'
OPENDOTA_BASE_URL = 'https://api.opendota.com/api/{func_name}/{params}'


def dota_api_call(func_name, api_key, **params):
    resp = requests.get(DOTA_BASE_URL.format(func_name=func_name, api_key=api_key, params=urlencode(params)))

    if not resp.ok:
        raise Exception('Something went wrong: GET {}: {} {}'.format(func_name, resp.status_code, resp.reason))
    return resp.json().get('result', {})


def opendota_api_call(func_name, *params):
    print(OPENDOTA_BASE_URL.format(func_name=func_name, params='/'.join(params)))
    resp = requests.get(OPENDOTA_BASE_URL.format(func_name=func_name, params='/'.join(params)))

    if not resp.ok:
        raise Exception('Something went wrong: GET {}: {} {}'.format(func_name, resp.status_code, resp.reason))

    return resp.json()


def get_heroes():
    data = opendota_api_call('heroes')
    return {hero['id']: hero['localized_name'] for hero in data}


def get_enemy_captain(match_id, side, player_names, api_key):
    data = dota_api_call('GetMatchDetails', api_key, match_id=match_id)
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

