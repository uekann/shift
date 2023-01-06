import pandas as pd

df_input = pd.read_excel("./data/shift_table.xlsx", index_col=[0, 1], header=0)
shift = df_input.values
names = list(df_input.columns)
times = list(df_input.index)

# シフト表1の作成
values1 = [
    [names[j] for j, x in enumerate(shift[i]) if x == 1] for i, t in enumerate(times)
]
df_output1 = pd.DataFrame(
    values1,
    index=pd.MultiIndex.from_tuples(times),
    columns=[f"人{i+1}" for i in range(sum(shift[0]))],
)
df_output1.to_csv("./output/shift1.csv")


# シフト表2の作成
values2 = [
    ["".join(t)[:4] for i, t in enumerate(times) if shift[i][j]]
    for j, m in enumerate(names)
]
df_output2 = pd.DataFrame(
    values2,
    index=names,
    columns=[f"時間{i+1}" for i in range(max([sum(j) for j in zip(*shift)]))],
)
df_output2.to_csv("./output/shift2.csv")
