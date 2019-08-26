from datetime import datetime
import math
import xlrd
import os.path
import copy
import csv
# import pandas
import pickle


class Player:
    players = []
    # Player のコンストラクタ

    def __init__(self, team, name, date, sex):
        self.team = team
        self.name = name
        # 属性
        self.attribute = Player.get_attri(Player.culc_age(date))
        self.sex = sex

    # 種目をセットする
    def set_player_item(self, item, distance, time):
        self.item = item
        self.distance = distance
        self.time = Player.culc_time(time)

    def get_Splayer_data(self):
        # Splayers用のデータを返す
        return self.team, self.name, self.attribute, self.time

    def get_register_data(self):
        # Readerクラスに登録するデータを返す
        team, name, attribute, time = self.get_Splayer_data()
        self.splayer = Splayer(team, name, attribute, time)
        return self.sex, self.item, self.distance, self.splayer

    # 特殊型式からタイムを取得するメソッド
    @staticmethod
    def culc_time(t):
        div_t = t.split("-")
        time = 0
        if len(div_t) == 2:
            time = int(div_t[0]) + (int(div_t[1])/100)
        else:
            time = (int(div_t[0])*60) + int(div_t[1]) + (int(div_t[2])/100)
        return time

    @staticmethod
    def culc_age(birthdayStr):
        if not (birthdayStr.isdigit() and len(birthdayStr) == 8):
            return -1
        dStr = datetime.now().strftime("%Y%m%d")
        return math.floor((int(dStr)-int(birthdayStr))/10000)

    @staticmethod
    def get_attri(age):
        if 7 <= age and age <= 12:
            return "小"+str(age-6)
        elif 13 <= age and age <= 15:
            return "中"+str(age-12)
        elif 16 <= age and age <= 18:
            return "高"+str(age-15)
        else:
            return "一般"

    @staticmethod
    def set_player(player):
        # 作成したPlayerをクラス配列に加える
        Player.players.append(player)

    @staticmethod
    def get_players():
        # 現在のplayersを返す
        return Player.players


class Splayer:
    def __init__(self, team, name, attribute, time):
        self.team = team
        self.name = name
        self.attribute = attribute
        self.time = time

    def print_name_time(self):
        print(self.name+" : "+str(self.time))


class Relay:
    teams = []

    def __init__(self, team, team_name, members, sex):
        self.team = team
        self.team_name = team_name
        self.members = members
        # 属性
        # self.attribute = Player.get_attri(Player.culc_age(date))
        self.sex = sex

    # 種目をセットする
    def set_relay_item(self, item, distance, time):
        self.item = item
        self.distance = distance
        self.time = Player.culc_time(time)

    def get_register_data(self):
        return self.sex, self.item, self.distance

    def print_name_time(self):
        print(self.team_name+" "+self.members+" : "+str(self.time))

    @staticmethod
    def set_team(team):
        # 作成したrelayをクラス配列に加える
        Relay.teams.append(team)


class Reader:
    def __init__(self):
        # 各種目での参加者を格納する
        self.items = {}

    # 参加者の種目登録を担う
    def set_item(self, sex, item, distance, splayer):
        # キーがあるか検証し，なければ作成する
        if sex in self.items:
            if item in self.items[sex]:
                if distance in self.items[sex][item]:
                    self.items[sex][item][distance].append(splayer)
                    return
                else:
                    self.items[sex][item][distance] = [splayer]
                    return
            else:
                self.items[sex][item] = {}
        else:
            self.items[sex] = {}
        # 再度登録しようとする精神
        self.set_item(sex, item, distance, splayer)

    def get_item(self, sex, item, distance):
        return self.items[sex][item][distance]

    # 参加者の登録
    def register_Players(self):
        for player in Player.players:
            sex, item, distance, splayer = player.get_register_data()
            self.set_item(sex, item, distance, splayer)

    # リレー参加者の登録
    def register_Relay(self):
        for team in Relay.teams:
            sex, item, distance = team.get_register_data()
            self.set_item(sex, item, distance, team)

    def sort_item(self):
        for sex in self.items.keys():
            for item in self.items[sex].keys():
                for dis in self.items[sex][item].keys():
                    self.sort(self.items[sex][item][dis])

    def sort(self, players):
        for i in range(len(players)):
            for j in range(len(players)-1, i, -1):
                if players[j].time > players[j-1].time:
                    players[j], players[j-1] = players[j-1], players[j]
        return players

    def sheet0_read(self, sheet):
        # チーム名，人数を取得
        team, player_num = self.get_teamdata(sheet)
        # 人数分のfor
        for i in range(player_num):
            # 人名取得
            player_name = str(sheet.cell(4+i, 1).value).replace("　", " ")
            # 参加種目に応じてエントリー
            item_index = 0
            # 生年月日を取得
            date = self.sheet0_get_date(sheet, i)
            # 性別取得
            sex = str(sheet.cell(4+i, 6).value)
            # もし参照セルがnullでなければ種目が存在する
            while(sheet.cell(4+i, 7+3*item_index).value):
                player = Player(team, player_name, date, sex)
                item, distance, time = self.sheet0_get_item(
                    sheet, i, item_index)
                player.set_player_item(item, distance, time)
                item_index += 1
                Player.set_player(player)

    def sheet1_read(self, sheet):
        # チーム名，人数を取得
        team, relay_num = self.get_teamdata(sheet)
        # 人数分のfor
        # print(relay_num)
        for i in range(relay_num):
            # メンバー全体の名前を連結して取得
            members = self.sheet1_get_member(sheet, i).replace("　", " ")
            team_name = str(sheet.cell(4+i, 1).value)
            sex = str(sheet.cell(4+i, 6).value)
            relay = Relay(team, team_name, members, sex)
            item, distance, time = self.sheet1_get_relay_item(
                sheet, i)
            relay.set_relay_item(item, distance, time)
            Relay.set_team(relay)

    # チーム名，人数を取得する

    def get_teamdata(self, sheet):
        r = 0
        while sheet.cell(4+r, 1).value != "":
            r += 1
        return str(sheet.cell(0, 3).value), r

    # 生年月日を取得する
    def sheet0_get_date(self, sheet, index):
        date = 0
        for i in range(3):
            date += int(sheet.cell(4+index, 3+i).value)*100**(2-i)
        return str(date)

    # 種目情報を取得する
    def sheet0_get_item(self, sheet, row, column):
        distance = str(int(sheet.cell(4+row, 7+3*column).value))
        item = str(sheet.cell(4+row, 7+3*column+1).value)
        time = str(sheet.cell(4+row, 7+3*column+2).value)
        return item, distance, time

    def sheet1_get_member(self, sheet, r):
        c = 2
        member = str(sheet.cell(4+r, c).value)
        for i in range(3):
            c += 1
            member = member + "・" + str(sheet.cell(4+r, c).value)
        return member

    def sheet1_get_relay_item(self, sheet, row):
        distance = str(int(sheet.cell(4+row, 8).value))
        item = str(sheet.cell(4+row, 9).value)
        time = str(sheet.cell(4+row, 10).value)
        return item, distance, time

    def print(self):
        for sex in self.items.keys():
            for item in self.items[sex].keys():
                for dis in self.items[sex][item].keys():
                    print(sex+" "+item+" "+dis)
                    for players in self.items[sex][item][dis]:
                        players.print_name_time()
                    print()


if __name__ == '__main__':
    name = os.path.dirname(os.path.abspath(__name__))

    # 絶対パスと相対パスをくっつける
    joined_path = os.path.join(name, '../data')
    # 正規化して絶対パスにする
    data_path = os.path.normpath(joined_path)
    data_path = str(name) + "/data"
    # 全てのファイル名を取得した
    data_list = os.listdir(data_path)
    # data_list = ["愛媛県マスターズ協会.xlsx"]
    reader = Reader()
    # 全てのファイルに対して実行
    # Playerクラスを全員+a分作成する
    for team in data_list:
        print(data_path+"/"+team)
        book = xlrd.open_workbook(data_path+"/"+team)
        sheet = book.sheet_by_index(0)
        sheet1 = book.sheet_by_index(1)
        reader.sheet0_read(sheet)
        reader.sheet1_read(sheet1)

    # Playerクラスから Readerクラスに登録する
    reader.register_Players()

    # Relayクラスから Readerクラスに登録する
    reader.register_Relay()

    # エントリータイムでsort()する
    reader.sort_item()

    # エントリー者の表示
    reader.print()
    with open('reader.pickle', mode='wb') as f:
        pickle.dump(reader, f)
