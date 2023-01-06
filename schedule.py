import pandas as pd
import pulp

N = 2  # 一度に入る人数

while True:
    # 希望表の取得
    df_input = pd.read_excel("./data/request_table.xlsx", index_col=[0, 1], header=0)
    members = list(df_input.columns)
    times = list(df_input.index)
    request_table = list(df_input.values)

    # 問題定義
    problem = pulp.LpProblem("schedule", sense=pulp.LpMaximize)

    # 変数定義
    x_tm = [
        [
            pulp.LpVariable(f"x_{i}_{j}", cat=pulp.LpBinary)
            for j, m in enumerate(members)
        ]
        for i, t in enumerate(times)
    ]
    # x_mt = list(zip(*x_tm))  # x_tmの転置行列

    # 目的関数定義
    problem += pulp.lpSum(
        [pulp.lpDot(request_table[i], x_tm[i]) for i, _ in enumerate(times)]
    )

    # 制約関数定義
    # 一度に入るのがN人になるようにする制約
    for i, t in enumerate(times):
        problem += pulp.lpSum(x_tm[i]) == N

    # 公平にシフトに入るようにする制約 (全員が一定回数入る)
    for i, m in enumerate(members):
        problem += pulp.lpSum(list(zip(*x_tm))[i]) >= N * len(times) // len(members)
        problem += pulp.lpSum(list(zip(*x_tm))[i]) <= N * len(times) // len(members) + 1

    # 解決
    try:
        status = problem.solve()
    except Exception as e:
        print("Can't solve. Try again.\n↓Error massage↓")
        print(e)
        if input("Try again?[y/n]") == "y":
            continue
        else:
            exit()
    else:
        print(f"Successfully resolved! [status : {pulp.LpStatus[status]}]")

    # 希望度との検証
    values = [list(map(lambda x: x.value(), x_tm[i])) for i, _ in enumerate(times)]
    mismatch = []
    for i, t in enumerate(times):
        for j, m in enumerate(members):
            if request_table[i][j] == 1 and values[i][j]:
                mismatch.append((m, t))

    if mismatch:
        print("Looks like someone's hopes were not met...\n↓Mismatch list↓")
        print(*map(lambda x: x[0] + " " + x[1][:3], mismatch), sep="\n")
        if input("Try again?[y/n]") == "y":
            continue
        else:
            exit()

    df_output = pd.DataFrame(
        values, columns=members, index=pd.MultiIndex.from_tuples(times)
    )
    df_output.to_excel("./data/shift_table.xlsx")
    exit()
