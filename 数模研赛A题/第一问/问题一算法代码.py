import pandas as pd
import numpy as np

# Load the provided fatigue data from the excel file
file_path_attachment_1 = '/mnt/data/附件1-疲劳评估数据.xlsx'

# 读取主轴扭矩和塔架推力的数据
main_shaft_torque_df = pd.read_excel(file_path_attachment_1, sheet_name='主轴扭矩')
tower_thrust_df = pd.read_excel(file_path_attachment_1, sheet_name='塔架推力')

# 移除最后两行的参考值 (等效疲劳载荷和累积疲劳损伤值)
cleaned_main_shaft_torque_df = main_shaft_torque_df.iloc[:-2, :]
cleaned_tower_thrust_df = tower_thrust_df.iloc[:-2, :]


# 累积疲劳损伤的计算 (基于Palmgren-Miner理论)
def calculate_fatigue_damage(load_series, C=1e6, m=3):
    cumulative_damage = np.zeros(load_series.shape[0])

    # 参考S-N曲线的疲劳寿命公式
    for i in range(1, load_series.shape[0]):
        stress_amplitude = abs(load_series[i] - load_series[i - 1])
        N_F = C / (stress_amplitude ** m if stress_amplitude != 0 else 1)

        # 损伤增量计算
        damage_increment = 1 / N_F
        cumulative_damage[i] = cumulative_damage[i - 1] + damage_increment

    return cumulative_damage


# Example: Calculate cumulative fatigue damage for one turbine (WT1)
wt1_main_shaft_damage = calculate_fatigue_damage(cleaned_main_shaft_torque_df['WT1'])
wt1_tower_thrust_damage = calculate_fatigue_damage(cleaned_tower_thrust_df['WT1'])

# 将结果汇总为DataFrame
time_series = cleaned_main_shaft_torque_df['T(s)']
result_df = pd.DataFrame({
    'Time (s)': time_series,
    'Main Shaft Torque Damage (WT1)': wt1_main_shaft_damage,
    'Tower Thrust Damage (WT1)': wt1_tower_thrust_damage
})

# 保存计算结果到文件
result_df.to_excel('/mnt/data/风机主轴与塔架累积疲劳损伤计算结果.xlsx', index=False)
