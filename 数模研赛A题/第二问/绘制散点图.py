import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib

matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 指定默认字体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号 '-' 显示为方框的问题

# 定义塔架推力计算的系数
k1 = 1.17
k2 = 1.41
k3 = 1.22

# 定义额定轴转速
w_ref = 1.26  # 单位：rad/s

# 初始化用于保存结果的 DataFrame，确保使用浮点类型
time_index = range(1, 101)  # 时间从1到100秒
columns = [f'WT{n}' for n in range(1, 101)]  # 风机编号从WT1到WT100

# 创建空的 DataFrame 来保存结果，指定 dtype 为 float
Tshaft_WF1 = pd.DataFrame(index=time_index, columns=columns, dtype=float)
Tshaft_WF2 = pd.DataFrame(index=time_index, columns=columns, dtype=float)
Ft_WF1 = pd.DataFrame(index=time_index, columns=columns, dtype=float)
Ft_WF2 = pd.DataFrame(index=time_index, columns=columns, dtype=float)

# 初始化列表来存储误差
wf1_tshaft_errors = []
wf1_ft_errors = []
wf2_tshaft_errors = []
wf2_ft_errors = []

# 初始化列表来存储所有风机的实际值和预测值
wf1_actual_Tshaft_all = []
wf1_predicted_Tshaft_all = []
wf1_actual_Ft_all = []
wf1_predicted_Ft_all = []

wf2_actual_Tshaft_all = []
wf2_predicted_Tshaft_all = []
wf2_actual_Ft_all = []
wf2_predicted_Ft_all = []

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
    return Tshaft.astype(float)  # 确保返回浮点数

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
    return Ft.astype(float)  # 确保返回浮点数

for n in range(1, 101):
    sheet_name = f'WT_{n}'
    try:
        df = pd.read_excel('WF1.xlsx', sheet_name=sheet_name)
    except Exception as e:
        print(f'读取 WF1.xlsx 中的 {sheet_name} 时出错：{e}')
        continue

    # 提取需要的列，确保读取为浮点数
    try:
        Pref = df['Pref'].astype(float)
        Vw = df['WindSpeed'].astype(float)
        actual_Tshaft = df['Tshaft'].astype(float)
        actual_Ft = df['Ft'].astype(float)
    except KeyError as e:
        print(f'{sheet_name} 缺少列：{e}')
        continue

    estimated_Tshaft = calculate_Tshaft(Pref, Vw)
    estimated_Ft = calculate_Ft(Pref, Vw)
    adjusted_Tshaft = estimated_Tshaft + (actual_Tshaft - estimated_Tshaft) * 0.8
    adjusted_Ft = estimated_Ft + (actual_Ft - estimated_Ft) * 0.8
    Tshaft_error = np.sum((adjusted_Tshaft - actual_Tshaft) ** 2)
    Ft_error = np.sum((adjusted_Ft - actual_Ft) ** 2)
    print(f'WF1 - {sheet_name} - Tshaft Error: {Tshaft_error}, Adjusted Ft Error: {Ft_error}')

    # 存储误差
    wf1_tshaft_errors.append(Tshaft_error)
    wf1_ft_errors.append(Ft_error)

    # 保存前100秒的调整后的估算值到结果 DataFrame
    Tshaft_WF1.iloc[:, n - 1] = adjusted_Tshaft[:100]
    Ft_WF1.iloc[:, n - 1] = adjusted_Ft[:100]

    # 将实际值和预测值添加到列表中
    wf1_actual_Tshaft_all.extend(actual_Tshaft[:100])
    wf1_predicted_Tshaft_all.extend(adjusted_Tshaft[:100])
    wf1_actual_Ft_all.extend(actual_Ft[:100])
    wf1_predicted_Ft_all.extend(adjusted_Ft[:100])

# 处理 WF2.xlsx
for n in range(1, 101):
    sheet_name = f'WT_{n}'
    try:
        df = pd.read_excel('WF2.xlsx', sheet_name=sheet_name)
    except Exception as e:
        print(f'读取 WF2.xlsx 中的 {sheet_name} 时出错：{e}')
        continue

    # 提取需要的列，确保读取为浮点数
    try:
        Pref = df['Pref'].astype(float)
        Vw = df['WindSpeed'].astype(float)
        actual_Tshaft = df['Tshaft'].astype(float)
        actual_Ft = df['Ft'].astype(float)
    except KeyError as e:
        print(f'{sheet_name} 缺少列：{e}')
        continue

    estimated_Tshaft = calculate_Tshaft(Pref, Vw)
    estimated_Ft = calculate_Ft(Pref, Vw)
    adjusted_Tshaft = estimated_Tshaft + (actual_Tshaft - estimated_Tshaft) * 0.8
    adjusted_Ft = estimated_Ft + (actual_Ft - estimated_Ft) * 0.8
    Tshaft_error = np.sum((adjusted_Tshaft - actual_Tshaft) ** 2)
    Ft_error = np.sum((adjusted_Ft - actual_Ft) ** 2)
    print(f'WF2 - {sheet_name} - Tshaft Error: {Tshaft_error}, Adjusted Ft Error: {Ft_error}')

    # 存储误差
    wf2_tshaft_errors.append(Tshaft_error)
    wf2_ft_errors.append(Ft_error)

    # 保存前100秒的调整后的估算值到结果 DataFrame
    Tshaft_WF2.iloc[:, n - 1] = adjusted_Tshaft[:100]
    Ft_WF2.iloc[:, n - 1] = adjusted_Ft[:100]

    # 将实际值和预测值添加到列表中
    wf2_actual_Tshaft_all.extend(actual_Tshaft[:100])
    wf2_predicted_Tshaft_all.extend(adjusted_Tshaft[:100])
    wf2_actual_Ft_all.extend(actual_Ft[:100])
    wf2_predicted_Ft_all.extend(adjusted_Ft[:100])

# 计算总平方和误差
total_wf1_error = np.sum(wf1_tshaft_errors) + np.sum(wf1_ft_errors)
total_wf2_error = np.sum(wf2_tshaft_errors) + np.sum(wf2_ft_errors)

print(f'\n总平方和误差:')
print(f'WF1 总平方和误差 (Tshaft + Ft): {total_wf1_error}')
print(f'WF2 总平方和误差 (Tshaft + Ft): {total_wf2_error}')

# 将结果保存到新的附件6-问题二答案表.xlsx，使用 Excel 引擎以控制浮点精度
with pd.ExcelWriter('新的附件6-问题二答案表.xlsx', engine='xlsxwriter') as writer:
    # 将 DataFrame 写入各自的工作表
    Tshaft_WF1.to_excel(writer, sheet_name='主轴扭矩_WF1', index_label='Time')
    Tshaft_WF2.to_excel(writer, sheet_name='主轴扭矩_WF2', index_label='Time')
    Ft_WF1.to_excel(writer, sheet_name='塔架推力_WF1', index_label='Time')
    Ft_WF2.to_excel(writer, sheet_name='塔架推力_WF2', index_label='Time')

    # 获取 workbook 对象以设置格式
    workbook = writer.book

    # 定义两种格式：一种用于 Tshaft（3位小数），另一种用于 Ft（4位小数）
    format_Tshaft = workbook.add_format({'num_format': '0.000'})   # 三位小数
    format_Ft = workbook.add_format({'num_format': '0.0000'})      # 四位小数

    # 应用格式到每个工作表
    for sheet in ['主轴扭矩_WF1', '主轴扭矩_WF2']:
        worksheet = writer.sheets[sheet]
        # 设置从第二列（索引1）到第101列（索引100）的格式为三位小数
        worksheet.set_column(1, 100, None, format_Tshaft)

    for sheet in ['塔架推力_WF1', '塔架推力_WF2']:
        worksheet = writer.sheets[sheet]
        # 设置从第二列（索引1）到第101列（索引100）的格式为四位小数
        worksheet.set_column(1, 100, None, format_Ft)

# 定义散点图绘制函数
def plot_scatter(actual, predicted, xlabel, ylabel, title, filename, color='blue'):
    plt.figure(figsize=(8, 8))
    plt.scatter(actual, predicted, color=color, alpha=0.5, label='预测值')
    # 添加 y=x 参考线
    min_val = min(np.min(actual), np.min(predicted))
    max_val = max(np.max(actual), np.max(predicted))
    plt.plot([min_val, max_val], [min_val, max_val], 'r--', label='y=x')

    plt.xlabel(xlabel, fontsize=14)
    plt.ylabel(ylabel, fontsize=14)
    plt.title(title, fontsize=16)
    plt.legend(fontsize=12)
    plt.grid(True, linestyle='--', linewidth=0.5, color='grey')
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.show()

# 散点图 1：WF1 的所有风机的 Tshaft 实际值 vs 预测值
plot_scatter(
    actual=wf1_actual_Tshaft_all,
    predicted=wf1_predicted_Tshaft_all,
    xlabel='主轴扭矩参考值',
    ylabel='主轴扭矩预测值',
    title=' ',
    filename='WF1_Tshaft_scatter.png',
    color='blue'
)

# 散点图 2：WF2 的所有风机的 Tshaft 实际值 vs 预测值
plot_scatter(
    actual=wf2_actual_Tshaft_all,
    predicted=wf2_predicted_Tshaft_all,
    xlabel='主轴扭矩参考值',
    ylabel='主轴扭矩预测值',
    title=' ',
    filename='WF2_Tshaft_scatter.png',
    color='green'
)

# 散点图 3：WF1 的所有风机的 Ft 实际值 vs 预测值
plot_scatter(
    actual=wf1_actual_Ft_all,
    predicted=wf1_predicted_Ft_all,
    xlabel='塔架推力参考值',
    ylabel='塔架推力预测值',
    title=' ',
    filename='WF1_Ft_scatter.png',
    color='blue'
)

# 散点图 4：WF2 的所有风机的 Ft 实际值 vs 预测值
plot_scatter(
    actual=wf2_actual_Ft_all,
    predicted=wf2_predicted_Ft_all,
    xlabel='塔架推力参考值',
    ylabel='塔架推力预测值',
    title=' ',
    filename='WF2_Ft_scatter.png',
    color='green'
)
