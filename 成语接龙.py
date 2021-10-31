import random
from openpyxl import Workbook
import sys
import os


def resource_path(relative_path):
    if getattr(sys, "frozen", False):  # 是否Bundle Resource
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


filename = resource_path(os.path.join("res", "chengyu.txt"))
# print(filename)
with open(filename, encoding="utf-8") as f:
    idioms = f.readlines()
wb1 = Workbook()
ws1 = wb1.active
ws1.title = "成语接龙"
times, round = map(int, input("请输入要玩的局数和回合数，用空格分隔开：").split())
counter = 1
while counter < times + 1:
    idiom, list1 = random.choice(idioms), []
    print(f"第{counter}局开始，我来出题：{idiom}")
    list1.append(idiom)
    for _ in range(round):
        user_in = input("输入成语：")
        if len(user_in) in list(range(4)):
            user_in = input("成语有这么短的吗，不要骗我呦，重新输入：")
        if user_in[0] == idiom[-2]:
            list1.append(user_in)
            char = list(user_in)[-1]
            if char in [i[0] for i in idioms]:
                idiom = random.choice([i for i in idioms if i.startswith(char)])
            else:
                if input("这个我没接上，输入n 回合继续：").lower() == "n":
                    list1.append("")
                    idiom = random.choice(idioms)
            list1.append(idiom)
            print(f"{idiom}")
        else:
            print("乱接就有点不讲武德了。")
    else:
        ws1.append(list1)
        counter += 1
file = input("游戏结束，请输入要保存的文件名：") + ".xlsx"
wb1.save(filename=file)
