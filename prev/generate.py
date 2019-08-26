from datetime import datetime
import math
import xlrd
import xlsxwriter
import os.path
import copy
import csv
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

    # 記述する内容を返す
    def get_content(self):
        name = self.name
        team = "("+self.team+")"
        attribute = self.attribute
        return name, team, attribute


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

    # 記述する内容を返す
    def get_content(self):
        name = self.team_name
        members = "("+self.members+")"
        return name, members

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
                if players[j].time < players[j-1].time:
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
        distance = str(sheet.cell(4+row, 7+3*column).value)
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
        distance = str(sheet.cell(4+row, 8).value)
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


class Writer:
    def __init__(self, reader, sheet, book):
        # 記述したい対象を格納したクラスを格納
        self.reader = reader
        # 記述するsheetの実体
        self.sheet = sheet
        # book
        self.book = book
        self.item_format = self.book.add_format()
        # self.item_format.set_font()
        self.item_format.set_bottom(5)
        self.item_format.set_top(5)

        self.time_format = self.book.add_format()
        self.time_format.set_bottom(1)

        self.set_property()
        # 左右のポインタ(0~3)
        self.left_p = 0
        self.right_p = 0
        # 右側のcolumn
        self.right_side_col = 6
        self.right_bias = 0
        # 現在のページ数
        self.page = 1
        # ページ数による y座標の重み
        self.page_bias = 0
        # 書き始める座標
        self.write_p = [0, 0]
        # 取り扱う種目がリレーか否か
        self.is_relay = False
        self.next_is_relay = False

    # excel に記述する
    def write_excel(self, order, program):
        time_str = "[   ]      :      . "
        time_title = "時間"

        for i, prog in enumerate(program):
            lane_half = int(self.lane/2)
            # リレーか否か
            if "リレー" in prog:
                self.is_relay = True
            else:
                self.is_relay = False
            sex = order[i][0]
            item = order[i][1]
            distance = order[i][2]
            # この種目の参加者リスト
            player_list = self.reader.items[sex][item][distance]
            # 行われるレースの数
            races = math.ceil(len(player_list)/self.lane)
            print("races: "+str(races))
            # 1レースに出場する選手の数
            race_lane = Writer.culc_race_lane(len(player_list))
            # 種目名を記述
            self.sheet.write(
                self.write_p[0]+1, self.write_p[1], prog, self.item_format)
            if self.is_relay:
                for j in range(1, 11):
                    self.sheet.write(
                        self.write_p[0]+1, self.write_p[1]+j, "", self.item_format)
            else:
                for j in range(1, 5):
                    self.sheet.write(
                        self.write_p[0]+1, self.write_p[1]+j, "", self.item_format)

            # レースの数繰り返す
            start = 0
            for race in range(races):
                # このレースに出場する参加者リスト(スライス)
                one_race_list = player_list[start: start+race_lane[race]]
                start += race_lane[race]
                # 速い順にする
                one_race_list.reverse()
                row = self.write_p[0]
                column = self.write_p[1]
                # レース番号の記述
                self.sheet.write(row+3, column, "【"+str(race+1)+"】")
                # リレーならば
                if self.is_relay:
                    self.sheet.write(3+row, 9, time_title)
                    j_ = 0
                    for j, relay_team in enumerate(one_race_list):
                        # 記述内容の取得
                        team_name, members = relay_team.get_content()
                        self.sheet.write(
                            3+row + self.lane_locate[j], 0, str(self.lane_locate[j])+". "+team_name)
                        self.sheet.write(
                            3+row + self.lane_locate[j], 3, members, self.time_format)
                        for p in range(1, 6):
                            self.sheet.write(
                                3+row + self.lane_locate[j], 3+p, "", self.time_format)
                        j_ += 1
                    for k in range(j_, self.lane):
                        self.sheet.write(
                            3+row + self.lane_locate[k], 0, str(self.lane_locate[k])+". ")
                        for p in range(6):
                            self.sheet.write(
                                3+row + self.lane_locate[k], 3+p, "", self.time_format)

                    for k in range(self.lane):
                        self.sheet.write(
                            3+row + self.lane_locate[k], 9, time_str, self.time_format)
                        self.sheet.write(
                            3+row + self.lane_locate[k], 10, "", self.time_format)

                    if race == races-1 and len(program)-1 > i:
                        if "リレー" in program[i+1]:
                            self.next_is_relay = True
                        else:
                            self.next_is_relay = False
                            # self.next_p()

                    self.next_p()
                # 個人戦ならば
                else:
                    self.sheet.write(3+row, column+3, time_title)
                    j_ = 0
                    for j, splayer in enumerate(one_race_list):
                        # 記述内容の取得
                        name, team, attribute = splayer.get_content()
                        self.sheet.write(
                            3+row + self.lane_locate[j], column, str(self.lane_locate[j])+". "+name)
                        self.sheet.write(
                            3+row + self.lane_locate[j], column+1, team)
                        self.sheet.write(
                            3+row + self.lane_locate[j], column+2, attribute)
                        j_ += 1
                    for k in range(j_, self.lane):
                        self.sheet.write(
                            3+row + self.lane_locate[k], column, str(self.lane_locate[k])+". ")
                    for k in range(self.lane):
                        self.sheet.write(
                            3+row + self.lane_locate[k], column+3, time_str, self.time_format)
                        self.sheet.write(
                            3+row + self.lane_locate[k], column+4, "", self.time_format)
                    if race == races-1 and len(program)-1 > i:
                        if "リレー" in program[i+1]:
                            self.next_is_relay = True
                        else:
                            self.next_is_relay = False
                            # self.next_p()
                    self.next_p()

    # 使用するレーン数を置く

    def set_lane_locate(self, lane):
        self.lane = lane
        lane_locate_ = []
        for i in range(lane):
            lane_locate_.append(i*(-1)**(i+1))
        lane_locate_[0] += int(self.lane/2)
        for i in range(1, lane):
            lane_locate_[i] += lane_locate_[i-1]
        self.lane_locate = lane_locate_

    # sheetのプロパティを変更する
    def set_property(self):
        for i in range(1000):
            self.sheet.set_row(i, 19)
        self.sheet.set_column("A:B", 13)
        self.sheet.set_column("C:C", 5)
        self.sheet.set_column("D:E", 6.7)

        self.sheet.set_column("F:F", 3)

        self.sheet.set_column("G:H", 13)
        self.sheet.set_column("I:I", 5)
        self.sheet.set_column("J:K", 6.7)

        # ポインタを次に進める

    # 次の writer_p を計算する
    def next_p(self):
        # もしリレーなら
        if self.next_is_relay:
            # リレー且つ左側を占拠していなければ
            if self.left_p < 3:
                self.left_p += 1
                self.right_p = self.left_p+1
                self.write_p = [(4+self.lane)*self.left_p + self.page_bias, 0]
            # 左側が全て埋まっていれば
            else:
                self.next_page()
        # 次が個人戦
        else:
            # もし個人戦で左側を占拠していなければ
            if self.left_p < 3:
                self.left_p += 1
                self.write_p = [(4+self.lane)*self.left_p + self.page_bias, 0]
            # 現在 left_p = 3 で左側が埋まった場合
            else:
                # 右側が埋まってなければ
                if self.right_p <= 3:
                    self.right_bias = self.right_side_col
                    # # 直前の種目がリレーだったら
                    # if self.is_relay and self.right_p < 3:
                    #     self.right_p += 1
                    self.write_p = [
                        (4+self.lane)*self.right_p + self.page_bias, self.right_side_col]
                    self.right_p += 1
                # 右側も埋まっていた
                else:
                    self.next_page()

    # 次のページを使う際の初期化
    def next_page(self):
        self.left_p = 0
        self.right_p = 0
        self.right_bias = 0
        self.page_bias = 4*(4+self.lane)*(self.page)
        self.write_p = [self.page_bias, 0]
        self.page += 1

    # 使用するレーン数を決める
    @staticmethod
    def culc_race_lane(num):
        num_list = []
        lane = 6
        while num != 0:
            # 以下 lane = 6 の例
            # もし参加者が12人以上なら
            if num >= 2*lane:
                num_list.append(lane)
                num = num - lane
            # 参加者が7~11人なら
            elif lane < num:
                # もし参加者が偶数
                if num % 2 == 0:
                    num = int(num/2)
                    num_list.append(num)
                    num_list.append(num)
                # 奇数
                else:
                    num = int(num/2)
                    num_list.append(num+1)
                    num_list.append(num)
                num = 0
            # 参加者が1~6人なら
            else:
                num_list.append(num)
                num = 0

        num_list.reverse()
        return num_list


if __name__ == "__main__":
    with open("reader.pickle", mode="rb") as f:
        reader = pickle.load(f)

    # reader.print()

    # order.csv を読み込む
    order = []
    with open("order.csv", 'r', encoding="utf-8") as r:
        read = csv.reader(r)
        for i in read:
            order.append(i[0].split())

    # program.csvを読み込む
    program = []
    with open("program.csv", 'r', encoding="utf-8") as r:
        read = csv.reader(r)
        for i in read:
            program.append(i[0])
    # 記述するexcelファイル
    workbook = xlsxwriter.Workbook('program2.xlsx')
    sheet = workbook.add_worksheet()
    writer = Writer(reader, sheet, workbook)
    # 使用するレーン数
    lane_num = 6
    # lane_num セット
    writer.set_lane_locate(lane_num)
    print(writer.lane_locate)
    writer.write_excel(order, program)
    workbook.close()
