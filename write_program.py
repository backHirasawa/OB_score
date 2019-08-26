from read_entry import Reader
from read_entry import Player
from read_entry import Splayer

import pickle
import xlrd
import xlsxwriter
import os.path
import copy
import csv
import re
import math


class Writer:
    def __init__(self, reader, sheet, book):
        # 記述したい対象を格納したクラスを格納
        self.reader = reader
        # 記述するsheetの実体
        self.sheet = sheet
        # book
        self.book = book
        self.item_format = self.book.add_format()
        self.item_format.set_bottom(5)
        self.item_format.set_top(5)

        self.time_format = self.book.add_format()
        self.time_format.set_bottom(1)

        self.time_format = self.book.add_format()
        self.time_format.set_bottom(1)

        self._set_property()
        # 左右のポインタ
        self.left_p = 0
        self.right_p = 0

        # →側のcolumn
        self.right_side_col = 6
        self.right_bias = 0

        # 現在のページ数
        self.page = 1
        # ページ数による y 座標の重み
        self.page_bias = 0

        # 書き始める座標
        self.write_p = [0, 0]

    def write_excel(self, order, program):
        time_str = "[   ]      :      . "
        time_title = "時間"

        for i, prog in enumerate(program):
            lane_half = math.ceil(self.lane/2)
            item = order[i][0]
            distance = order[i][1]
            # 参加者リスト取得
            splayers = self.reader.get_item(item, distance)
            # この競技で行われるレースの数
            races = math.ceil(len(splayers)/self.lane)
            race_lane = Writer.culc_race_lane(len(splayers))
            self.sheet.write(
                self.write_p[0]+1,
                self.write_p[1],
                prog,
                self.item_format
            )

            # 競技名の罫線の距離を延ばす
            for j in range(1, 5):
                self.sheet.write(
                    self.write_p[0]+1,
                    self.write_p[1]+j,
                    "",
                    self.item_format
                )

            start = 0
            for race in range(races):
                # 各レース出場者をスライス
                one_race_list = splayers[start: start + race_lane[race]]
                start += race_lane[race]
                # 速い順にする
                one_race_list.reverse()
                row = self.write_p[0]
                column = self.write_p[1]
                # レース番号を記述
                self.sheet.write(
                    row+3,
                    column,
                    "【"+str(race+1)+"】"
                )

                self.sheet.write(
                    3+row,
                    column+3,
                    time_title
                )
                # 謎の変数
                j_ = 0

                # 1レースに出場する選手の数だけ回す
                for j, splayer in enumerate(one_race_list):
                    # 記述内容の取得
                    name, age, department = splayer.get_contents()
                    self.sheet.write(
                        3+row + self.lane_locate[j],
                        column,
                        str(self.lane_locate[j])+". "+name
                    )
                    self.sheet.write(
                        3+row + self.lane_locate[j],
                        column+1,
                        age
                    )
                    self.sheet.write(
                        3+row + self.lane_locate[j],
                        column+2,
                        department
                    )
                    j_ += 1

                # ドットの記述
                for k in range(j_, self.lane):
                    self.sheet.write(
                        3+row + self.lane_locate[k],
                        column,
                        str(self.lane_locate[k])+". "
                    )
                # 時間部分の記述
                for k in range(self.lane):
                    self.sheet.write(
                        3+row + self.lane_locate[k],
                        column+3,
                        time_str, self.time_format
                    )
                    self.sheet.write(
                        3+row + self.lane_locate[k],
                        column+4,
                        "",
                        self.time_format
                    )
                self._next_p()

    def set_lane_locate(self, lane):
        self.lane = lane
        lane_locate_ = []
        for i in range(lane):
            lane_locate_.append(i*(-1)**(i+1))
        lane_locate_[0] += int(self.lane/2)
        for i in range(1, lane):
            lane_locate_[i] += lane_locate_[i-1]
        self.lane_locate = lane_locate_

    def _set_property(self):
        for i in range(1000):
            self.sheet.set_row(i, 19)
        self.sheet.set_column("A:B", 13)
        self.sheet.set_column("C:C", 5)
        self.sheet.set_column("D:E", 6.7)

        self.sheet.set_column("F:F", 3)

        self.sheet.set_column("G:H", 13)
        self.sheet.set_column("I:I", 5)
        self.sheet.set_column("J:K", 6.7)

    def _next_p(self):
        # 個人戦で左側を占拠していなければ
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
                self._next_page()

    # 次のページを使う際の初期化

    def _next_page(self):
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

    order = []
    with open("order_OB.csv", 'r', encoding="utf-8") as r:
        read = csv.reader(r)
        for i in read:
            order.append(i[0].split())
        # print(order)

    # 競技名を作成
    program = []
    for i, item_dis in enumerate(order):
        pro = "No."+str(i+1)+" "+item_dis[1]+"m "+item_dis[0]
        program.append(pro)
        print(pro)

        # 記述するexcelファイル
    workbook = xlsxwriter.Workbook('program2.xlsx')
    sheet = workbook.add_worksheet()
    writer = Writer(reader, sheet, workbook)
    # 使用するレーン数
    lane_num = 6
    writer.set_lane_locate(lane_num)
    print(writer.lane_locate)
    writer.write_excel(order, program)
    workbook.close()
