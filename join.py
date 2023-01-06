import glob
import itertools
import re

import numpy as np
import pandas as pd

# ファイルの取得
files = glob.glob("./submit/*.xlsx")
mat = np.zeros((len(files), 30))
names = []

p = re.compile(
    r"[_＿]+([^0-9^a-z^A-Z]+)[_＿]+([^0-9^a-z^A-Z^\.^\-]+)"
)  # 名前の抽出用正規表現（日本語の名前のみ対応）

for idx, file in enumerate(files):
    names.append(" ".join(p.search(file).groups()))
    df = pd.read_excel(file, sheet_name="集計用", header=None)
    mat[idx] = df.values.T.flatten()


# 　データフレームの作成
df = pd.DataFrame(
    mat.T,
    columns=names,
    index=pd.MultiIndex.from_tuples(
        list(
            itertools.product(
                ["月", "火", "水", "木", "金"],
                [
                    "2限   10:00〜11:25",
                    "昼休み 11:25〜12:15",
                    "3限   12:15〜13:30",
                    "4限   13:45〜15:00",
                    "5限   15:15〜16:30",
                    "6限   16:45〜18:30",
                ],
            )
        )
    ),
)

# 希望表として出力
df.to_excel("./data/request_table.xlsx")
