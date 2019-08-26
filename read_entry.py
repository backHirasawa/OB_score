import datetime
import math
import xlrd
import os.path
import copy
import csv
# import pandas
import pickle
import re


class Player:
    # 全参加者を保持するクラス変数
    players = []

    def __init__(self, sex, name, age, department):
        self.sex = sex
        self.name = name
        # 学年/卒業年度(str)
        self.age = Player._change_age(age)
        # 学科
        self.department = Player._change_department(department)
        # 泳いだ総距離
        self.sum_distance = 0
        # 獲得したポイント
        self.point = 0

    # 泳いだ距離を加算する
    def add_distance(self, distance):
        self.sum_distance += distance

    # 獲得したポイントを加算する
    def add_point(self, point):
        self.point += point

    # 学年/卒業年度を変換する
    @staticmethod
    def _change_age(age):
        # 本科生
        if age <= 5:
            # 3以下ならそのまま
            if age <= 3:
                # 型をstringに変換しておく
                return str(age)
            # 4,5年はOB1,2なので
            else:
                age -= 3
                return "OB"+str(age)
        # 卒業生
        else:
            today = datetime.date.today()
            age = today.year - age + 2
            return "OB"+str(age)

    # 学科を変換する
    @staticmethod
    def _change_department(department):
        if department == "機械":
            return "M"
        elif department == "電気":
            return "E"
        elif department == "電子制御":
            return "S"
        elif department == "情報":
            return "I"
        else:
            return "C"

    # 作成したPlayerをクラス配列に加える
    @staticmethod
    def set_player(player):
        Player.players.append(player)

    # 現在の選手全員を返す
    @staticmethod
    def get_players():
        return Player.players

    @staticmethod
    def print_players():
        for player in Player.players:
            print(player.name+" : "+str(player.age))


class Splayer:
    splayers = []

    def __init__(self, player, item, distance, time):
        # sortや登録に使う
        self.player = player
        self.item = item
        self.distance = distance
        self.time = time
        self.player.add_distance(int(self.distance))

    def print_name_time(self):
        print(self.name+" : "+str(self.time))

    def get_contents(self):
        return self.player.name, self.player.age, self.player.department

    @staticmethod
    def set_splayer(splayer):
        Splayer.splayers.append(splayer)


class Reader:
    def __init__(self):
        # 各種目での参加者を格納する
        self.items = {}

    # 参加者の種目登録を行う
    def set_item(self, item, distance, splayer):
        # キーがあるか検証し，なければ作成する
        if item in self.items:
            if distance in self.items[item]:
                self.items[item][distance].append(splayer)
                return
            else:
                self.items[item][distance] = [splayer]
                return
        else:
            self.items[item] = {}
        # 再帰して再登録
        self.set_item(item, distance, splayer)

    def get_item(self, item, distance):
        return self.items[item][distance]

    # 参加者の登録
    def register(self):
        for splayer in Splayer.splayers:
            # sex, item, distance, splayer = player.get_register_data()
            self.set_item(splayer.item, splayer.distance, splayer)

    # エントリータイムで sort する
    def sort_item(self):
        for item in self.items.keys():
            for distance in self.items[item].keys():
                # まずはタイムゼロの人のタイムを平均タイムで登録する
                self.items[item][distance] = self._time_average(
                    self.items[item][distance])
                # sortする
                self.items[item][distance] = self._sort(
                    self.items[item][distance])

    # タイム平均化を適用
    def _time_average(self, splayers):
        sum_time = 0
        non_zero_num = 0
        for splayer in splayers:
            sum_time += splayer.time
            # 母数調整
            if splayer.time != 0:
                non_zero_num += 1
        # 平均を計算
        average = sum_time/non_zero_num
        for splayer in splayers:
            if splayer.time == 0:
                splayer.time = average
        return splayers

    def _sort(self, splayers):
        for i in range(len(splayers)):
            for j in range(len(splayers)-1, i, -1):
                if splayers[j].time > splayers[j-1].time:
                    splayers[j], splayers[j-1] = splayers[j-1], splayers[j]
        return splayers

    # シートから読み取る
    def read_sheet(self, sheet):
        # 人数を取得
        player_num = sheet.nrows - 1
        item_base = 5
        # 人数分の回す
        for i in range(player_num):
            # 選手データの取得
            sex, name, age, department = self._read_player_data(sheet, i)
            # 参加種目の取得
            item_index = 0
            # 参加者作成
            player = Player(sex, name, age, department)
            # 選手欄に登録
            Player.set_player(player)
            # 参照セルがnullでなければ種目が存在する
            while(sheet.cell(1+i, item_base+2*item_index).value) and item_index <= 2:
                item, distance, time = self._read_item(
                    sheet, i, item_index, item_base)
                # 各種目登録用の分身
                splayer = Splayer(player, item, distance, time)
                # 分身を登録
                Splayer.set_splayer(splayer)
                item_index += 1

    def _read_player_data(self, sheet, i):
        sex = sheet.cell(1+i, 1).value
        name = re.sub(" |　", "", sheet.cell(1+i, 2).value)
        age = int(sheet.cell(1+i, 3).value)
        department = sheet.cell(1+i, 4).value
        return sex, name, age, department

    # 種目情報を取得する
    def _read_item(self, sheet, i, item_index, item_base):
        item_dis = str(sheet.cell(1+i, item_base+2*item_index).value)
        # MG処理(選手として登録．以下のkeyで参照可能)
        if item_dis == "MG":
            item = "MG"
            distance = "0"
            # zero_divを回避
            time = 1
        else:
            item_dis = re.sub("m|ｍ", " ", item_dis).split()
            # print(item_dis)
            distance = item_dis[0]
            item = item_dis[1]

            if Reader.is_float(sheet.cell(1+i, item_base+2*item_index+1).value):
                time = float(sheet.cell(1+i, item_base+2*item_index+1).value)
            # もしエントリータイムが書いていなければ
            else:
                # 後に平均化する
                # print("time : None")
                time = 0

        return item, distance, time

    def print_item(self, item, distance):
        for splayer in self.items[item][distance]:
            print(splayer.player.name+" :"+str(splayer.time))

    @staticmethod
    def is_float(s):
        try:
            float(s)
        except:
            return False
        return True


if __name__ == "__main__":
    file_name = "entry_list.xlsx"
    reader = Reader()

    book = xlrd.open_workbook(file_name)
    sheet = book.sheet_by_index(0)
    reader.read_sheet(sheet)

    # Splayerクラスから， Readerクラスに登録する
    reader.register()
    # エントリータイムで sort する
    reader.sort_item()

    reader.print_item("平泳ぎ", "50")

    # reader.print()
    with open('reader.pickle', mode='wb') as f:
        pickle.dump(reader, f)
