import csv

order = []
with open("program.csv", 'r', encoding="utf-8") as r:
    read = csv.reader(r)
    for i in read:
        order.append(i[0])
        # print(i)

print(order)
