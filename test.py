import xlrd


def get_player_num(sheet):
    # r = sheet.nrows - 1
    # for i in range(r):
    # 	column =  sh
    return 0


if __name__ == "__main__":
    file_name = "entry_list.xlsx"
    book = xlrd.open_workbook(file_name)
    sheet = book.sheet_by_index(0)
    r = get_player_num(sheet)
    print(r)
