import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib

matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False

k1 = 1.17
k2 = 1.41
k3 = 1.22
w_ref = 1.26

# 初始化用于保存结果的 DataFrame，确保使用浮点类型
time_index = range(1, 101)
columns = [f'WT{n}' for n in range(1, 101)]

# 创建空的 DataFrame 来保存结果，指定 dtype 为 float
Tshaft_WF1 = pd.DataFrame(index=time_index, columns=columns, dtype=float)
Tshaft_WF2 = pd.DataFrame(index=time_index, columns=columns, dtype=float)
Ft_WF1 = pd.DataFrame(index=time_index, columns=columns, dtype=float)
Ft_WF2 = pd.DataFrame(index=time_index, columns=columns, dtype=float)

# 初始化列表来存储NRMSE
wf1_tshaft_nrmses = []
wf1_ft_nrmses = []
wf2_tshaft_nrmses = []
wf2_ft_nrmses = []


# 定义计算主轴扭矩的函数
def calculate_Tshaft(Pref, Vw):
    Tshaft = np.where(
        Vw < 10.2,
        (10.2 * Pref) / (w_ref * Vw),
        np.where(
            Vw <= 12.2,
            Pref / w_ref,
            (12.2 * Pref) / (w_ref * Vw)
        )
    )
    return Tshaft.astype(float)


# 定义计算塔架推力的函数
def calculate_Ft(Pref, Vw):
    Ft = np.where(
        Vw < 10.2,
        k1 * Pref / Vw,
        np.where(
            Vw <= 12.2,
            k2 * Pref / Vw,
            k3 * Pref / Vw
        )
    )
    return Ft.astype(float)


# 处理 WF1_data.xlsx
for n in range(1, 101):
    sheet_name = f'WT_{n}'
    try:
        df = pd.read_excel('WF1.xlsx', sheet_name=sheet_name)
        Pref = df['Pref'].values
        Vw = df['Vw'].values

        # 计算主轴扭矩和塔架推力
        Tshaft_WF1[f'WT{n}'] = calculate_Tshaft(Pref, Vw)
        Ft_WF1[f'WT{n}'] = calculate_Ft(Pref, Vw)

        # 计算 NRMSE（假设有真实值 y_true）
        y_true = np.random.rand(len(Pref))  # 示例真实值，替换为实际值
        nrmse_tshaft = np.sqrt(np.mean((Tshaft_WF1[f'WT{n}'] - y_true) ** 2)) / np.mean(y_true)
        nrmse_ft = np.sqrt(np.mean((Ft_WF1[f'WT{n}'] - y_true) ** 2)) / np.mean(y_true)

        wf1_tshaft_nrmses.append(nrmse_tshaft)
        wf1_ft_nrmses.append(nrmse_ft)

    except Exception as e:
        print(f'读取 WF1.xlsx 中的 {sheet_name} 时出错：{e}')
        continue

# 处理 WF2_data.xlsx
for n in range(1, 101):
    sheet_name = f'WT_{n}'
    try:
        df = pd.read_excel('WF2.xlsx', sheet_name=sheet_name)
        Pref = df['Pref'].values
        Vw = df['Vw'].values

        # 计算主轴扭矩和塔架推力
        Tshaft_WF2[f'WT{n}'] = calculate_Tshaft(Pref, Vw)
        Ft_WF2[f'WT{n}'] = calculate_Ft(Pref, Vw)

        # 计算 NRMSE（假设有真实值 y_true）
        y_true = np.random.rand(len(Pref))  # 示例真实值，替换为实际值
        nrmse_tshaft = np.sqrt(np.mean((Tshaft_WF2[f'WT{n}'] - y_true) ** 2)) / np.mean(y_true)
        nrmse_ft = np.sqrt(np.mean((Ft_WF2[f'WT{n}'] - y_true) ** 2)) / np.mean(y_true)

        wf2_tshaft_nrmses.append(nrmse_tshaft)
        wf2_ft_nrmses.append(nrmse_ft)

    except Exception as e:
        print(f'读取 WF2.xlsx 中的 {sheet_name} 时出错：{e}')
        continue
