"""
读取txt文件，取有效数据，绘制分流排名分布图。
"""
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

filename = 'data/分流专业排名2020.txt'
fig_title = "工科试验班2019级专业分流推导排名分布图（仅供参考）"

group_distance = 20  # 界限排名属于排名较高的组
max_rank = 256

# 先读取一个原始表，每个数据结构专业-[所有有效排名]
ranks = {}
with open(filename,errors='ignore') as fp:
    for line in fp:
        ls = line.strip().split()
        if len(ls)!=2:
            continue
        try:
            rank=int(ls[1])
        except ValueError:
            continue
        else:
            ranks.setdefault(ls[0],[]).append(rank)

for lst in ranks.values():
    lst.sort()

groups = {}
lst=[]
for major,rank_values in ranks.items():
    lst = []
    groups[major]=lst
    # precondition: rank_values已经升序排序
    for rank in rank_values:
        # 如果跳过了一些组，先全部补充0. post condition: 当前排名所在组的数据存在且是最后一个。
        rank_group = rank//group_distance
        while len(lst) <= rank_group:
            lst.append(0)
        lst[rank_group]+=1
    while len(lst) <= max_rank//group_distance:
        lst.append(0)
    for i in range(len(lst)):
        lst[i]/=len(rank_values)
        lst[i]*=100

indexes = []
fromindex=0
while fromindex <= max_rank:
    indexes.append(fromindex)
    fromindex+=group_distance

from pprint import pprint as printf
printf(list(groups.values()))
print(indexes)

df = pd.DataFrame(data=groups,index=indexes)

plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

cmap = sns.cm.rocket_r
sns.heatmap(df,annot=True,cmap=cmap)
plt.tight_layout()
plt.title(fig_title)
plt.show()
