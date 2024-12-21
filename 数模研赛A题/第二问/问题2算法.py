import numpy as np
import pandas as pd

# 定义常量
rho = 1.225  # 空气密度，单位：kg/m³
A = np.pi * (63) ** 2  # 假设风轮扫掠面积，半径为63米
C_t = 0.5  # 推力系数
omega = 1.2  # 假设的角速度，单位：rad/s

# 读取 WF1 和 WF2 的 xlsx 文件，包含所有风机的 sheet
wf1_data = pd.ExcelFile('WF1_data.xlsx')
wf2_data = pd.ExcelFile('WF2_data.xlsx')

# 读取附件6-问题二答案表的Excel文件
attachment_6 = pd.ExcelFile('附件6答案表.xlsx')

# 读取主轴扭矩和塔架推力的 sheet 表
main_shaft_sheet = pd.read_excel(attachment_6, sheet_name='主轴扭矩')
tower_thrust_sheet = pd.read_excel(attachment_6, sheet_name='塔架推力')

time_data = np.arange(1, 2001)
main_shaft_sheet = main_shaft_sheet.reindex(range(2000))
tower_thrust_sheet = tower_thrust_sheet.reindex(range(2000))

# 填充时间列
main_shaft_sheet['Time'] = time_data
tower_thrust_sheet['Time'] = time_data

# 创建存储计算结果的列表
main_shaft_results = []
tower_thrust_results = []

# 遍历 WF1 的所有工作表，每个工作表代表一台风机
for sheet_name in wf1_data.sheet_names:
    wf1_turbine = pd.read_excel(wf1_data, sheet_name=sheet_name)  # 读取每个风机的数据

    # 提取 Pref 和 WindSpeed 列的数据
    wind_speed_wf1 = wf1_turbine['WindSpeed']
    power_ref_wf1 = wf1_turbine['Pref']

    # 计算主轴扭矩和塔架推力
    main_shaft_torque_wf1 = power_ref_wf1 / omega
    tower_thrust_wf1 = C_t * 0.5 * rho * A * wind_speed_wf1 ** 2

    # 将计算结果添加到列表中
    main_shaft_results.append(main_shaft_torque_wf1.values)  # 计算结果存储为数组
    tower_thrust_results.append(tower_thrust_wf1.values)

# 遍历 WF2 的所有工作表，每个工作表代表一台风机
for sheet_name in wf2_data.sheet_names:
    wf2_turbine = pd.read_excel(wf2_data, sheet_name=sheet_name)  # 读取每个风机的数据

    # 提取 Pref 和 WindSpeed 列的数据
    wind_speed_wf2 = wf2_turbine['WindSpeed']
    power_ref_wf2 = wf2_turbine['Pref']

    # 计算主轴扭矩和塔架推力
    main_shaft_torque_wf2 = power_ref_wf2 / omega
    tower_thrust_wf2 = C_t * 0.5 * rho * A * wind_speed_wf2 ** 2

    # 将计算结果添加到列表中
    main_shaft_results.append(main_shaft_torque_wf2.values)  # 计算结果存储为数组
    tower_thrust_results.append(tower_thrust_wf2.values)

# 将结果转换为 DataFrame，并按照风机编号插入附件6的相应工作表
# 假设附件6中的列名从第二列开始代表各个风机 WT1, WT2, ... WTn

# 生成新的DataFrame用于主轴扭矩
for i, col in enumerate(main_shaft_sheet.columns[1:], start=0):
    if i < len(main_shaft_results):  # 确保索引不超出范围
        main_shaft_sheet[col] = main_shaft_results[i]  # 每个风机的数据插入对应的列

# 生成新的DataFrame用于塔架推力
for i, col in enumerate(tower_thrust_sheet.columns[1:], start=0):
    if i < len(tower_thrust_results):  # 确保索引不超出范围
        tower_thrust_sheet[col] = tower_thrust_results[i]  # 每个风机的数据插入对应的列

# 将更新后的数据保存回附件6，设置 if_sheet_exists='replace' 来覆盖现有的工作表
with pd.ExcelWriter('附件6答案表.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    main_shaft_sheet.to_excel(writer, sheet_name='主轴扭矩', index=False)
    tower_thrust_sheet.to_excel(writer, sheet_name='塔架推力', index=False)

print("结果已成功保存到 '附件6答案表.xlsx'")
