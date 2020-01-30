import xlsxwriter

from itertools import zip_longest
from lxml import etree


COLORS = ['white', 'yellow', 'orange', 'cyan', 'silver', '#FF7F50', '#FFD700', '#ADFF2F',
          '#40E0D0', '#00BFFF', '#D8BFD8', '#FFC0CB', '#FAEBD7', '#E6E6FA', '#FFD700']


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
        if team.parsed_matches:
            self.worksheet.write(self.row, 0, team.name)
            self.row += 1
            self.worksheet.write(self.row, 0, 'PICKS')
            self.worksheet.write(self.row, 5, 'PICK')
            self.worksheet.write(self.row, 6, 'RESULT')
            self.worksheet.write(self.row, 7, 'SIDE')
            self.worksheet.write(self.row, 8, 'CAPTAIN')
            self.worksheet.write(self.row, 9, 'ENEMY BANS')
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

        for column, enemy_ban in enumerate(sorted(match.enemy_bans.values(), key=lambda p: p['order']), start=9):
            color = self.colors.get(team.enemy_ban_count[enemy_ban['name']], self.colors[1])
            self.used_colors.add(team.enemy_ban_count[ban['name']])
            self.write_hero(column, color, **enemy_ban)

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

        self.worksheet.write(self.row, 5, 'ENEMY BANS')
        self.worksheet.write(self.row, 6, 'COUNT')

        self.worksheet.write(self.row, 8, 'BAN')
        self.worksheet.write(self.row, 9, 'COUNT')
        self.row += 1

        pick_row, enemy_bans_row, ban_row = self.row, self.row, self.row
        pick_count = sorted(team.pick_count.items(), key=lambda p: p[1]['count'], reverse=True)
        enemy_ban_count = sorted(team.enemy_ban_count.items(), key=lambda b: b[1], reverse=True)
        ban_count = sorted(team.ban_count.items(), key=lambda b: b[1], reverse=True)

        for pick, enemy_bans, ban in zip_longest(pick_count, enemy_ban_count, ban_count):
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
            if enemy_bans:
                self.worksheet.write(enemy_bans_row, 5, enemy_bans[0])
                self.worksheet.write(enemy_bans_row, 6, enemy_bans[1])
                enemy_bans_row += 1
            if ban:
                self.worksheet.write(ban_row, 8, ban[0])
                self.worksheet.write(ban_row, 9, ban[1])
                ban_row += 1

        self.row = max(pick_row, enemy_bans_row, ban_row)
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

    def write_players(self, players):
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
                self.worksheet.write(self.row, 0, hero)
                self.worksheet.write(self.row, 1, data['count'])
                self.worksheet.write(self.row, 2, win_rate)
                self.row += 1
            self.row += 1



        self.row += 1


class HtmlWriter(object):
    def __init__(self, file):
        self.file = file
        parser = etree.HTMLParser(remove_blank_text=True)
        self.html = etree.parse('template.html', parser)
        for element in self.html.iter():
            element.tail = None

        self.body = self.html.xpath('body')[0]
        self.classes = self.create_classes()

    def close(self):
        with open(self.file, 'w') as f:
            f.write(etree.tostring(self.html, method='html', pretty_print=True).decode('utf-8'))

    def create_classes(self):
        return {2: 'two', 3: 'three', 4: 'four', 5: 'five', 6: 'six', 7: 'seven', 8: 'eight',
                9: 'nine', 10: 'ten', 11: 'eleven', 12: 'twelve', 13: 'thirteen', 14: 'fourteen',
                15: 'fifteen', 16: 'sixteen'}

    def write_matches(self, team):
        if team.parsed_matches:
            table = self.write_parsed_matches_header()
            for match in team.parsed_matches:
                tbody = etree.SubElement(table, 'tbody')
                tr = etree.SubElement(tbody, 'tr')
                self.write_parsed_match(tr, match, team)

                tbody = etree.SubElement(table, 'tbody')
                tr = etree.SubElement(tbody, 'tr')
                tr.set('class', 'hide')
                self.write_hero_data_header(tr)
                self.write_parsed_hero_data(tr, match, team)

    def write_parsed_matches_header(self):
        table = etree.SubElement(self.body, 'table')
        header = etree.SubElement(table, 'tr')
        row = etree.SubElement(header, 'th', colspan='5')
        row.text = 'Picks'

        for text in ['Result', 'Side', 'Captain']:
            row = etree.SubElement(header, 'th')
            row.text = text

        row = etree.SubElement(header, 'th', text='Enemy Bans', colspan='6')
        row.text = 'Enemy Bans'
        row = etree.SubElement(header, 'th', colspan='6')
        row.text = 'Bans'

        return table

    def write_parsed_match(self, tr, match, team):
        for pick in sorted(match.picks.values(), key=lambda p: p['order']):
            count = team.pick_count[pick['name']]['count']
            class_name = self.classes.get(count)
            self.write_hero(tr, class_name, str(match.match_id), **pick)

        self.write_match_details(tr, match)

        for enemy_ban in sorted(match.enemy_bans.values(), key=lambda p: p['order']):
            count = team.enemy_ban_count[enemy_ban['name']]
            class_name = self.classes.get(count)
            self.write_hero(tr, class_name, str(match.match_id), **enemy_ban)

        for ban in sorted(match.bans.values(), key=lambda p: p['order']):
            count = team.ban_count[ban['name']]
            class_name = self.classes.get(count)
            self.write_hero(tr, class_name, str(match.match_id), **ban)

    def write_hero(self, tr, class_name, match_id, name=None, lane=None, roaming=False, **kwargs):
        data = name
        if lane:
            data += ' ' + lane

        if roaming:
            data += ' (R)'

        subbed_name = name.replace(' ', '_')
        td = etree.SubElement(tr, 'td', id=subbed_name)
        if class_name:
            td.set('class', '{} {}'.format(subbed_name, class_name))
        else:
            td.set('class', subbed_name)

        cell = etree.SubElement(td, 'label')
        cell.text = data
        cell.set('for', match_id)

        cell = etree.SubElement(td, 'input', type='checkbox', name=match_id, id=match_id)
        cell.set('data-toggle', 'toggle')

    def write_match_details(self, tr, match):
        for data in ['W' if match.win else 'L', match.side, match.enemy_captain]:
            td = etree.SubElement(tr, 'td')

            cell = etree.SubElement(td, 'label')
            cell.text = data
            cell.set('for', str(match.match_id))

            cell = etree.SubElement(td, 'input', type='checkbox', name=str(match.match_id), id=str(match.match_id))
            cell.set('data-toggle', 'toggle')

    def write_hero_data_header(self, tr):
        cell = etree.SubElement(tr, 'td')
        cell.set('class', 'borderless')

        for label in ['Player', 'Hero', 'Lane', 'Kills', 'Deaths', 'Assists',
                      'Last Hits', 'Denies', 'GPM', 'XPM', 'Obs', 'Sentries']:
            cell = etree.SubElement(tr, 'td')
            cell.text = label

    def write_parsed_hero_data(self, tr, match, team):
        pass

    def write_summary(self, team):
        pass

    def write_legend(self):
        pass

    def write_players(self, players):
        pass
