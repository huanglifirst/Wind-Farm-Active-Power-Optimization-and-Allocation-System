import numpy as np
import pandas as pd
from scipy.optimize import minimize

# 读取Excel文件的数据
wf1_data_path = 'WF1_data.xlsx'
wf2_data_path = 'WF2_data.xlsx'

# 加载 WF1_data.xlsx 和 WF2_data.xlsx 文件中的所有工作表
wf1_data = pd.read_excel(wf1_data_path, sheet_name=None)
wf2_data = pd.read_excel(wf2_data_path, sheet_name=None)


# 定义塔架推力的计算函数
def calculate_tower_thrust(pref, wind_speed, k1, k2, k3):
    V_ref = 11.2
    Delta_V = 1.0
    if wind_speed == 0:
        return 0.0
    elif wind_speed < (V_ref - Delta_V):
        return k1 * (pref / wind_speed)
    elif (V_ref - Delta_V) <= wind_speed <= (V_ref + Delta_V):
        return k2 * (pref / wind_speed)
    else:
        return k3 * (pref / wind_speed)


# 误差分析函数（平方和差异）
def calculate_error(estimated, reference):
    return np.sum((reference - estimated) ** 2)


# 目标函数：仅最小化塔架推力误差平方和
def objective_function(k_values, wt_data):
    k1, k2, k3 = k_values
    pref = wt_data['Pref'].values
    wind_speed = wt_data['WindSpeed'].values
    ft_ref = wt_data['Ft'].values

    # Vector化计算塔架推力
    ft_theoretical = np.where(
        wind_speed < 10.2,
        k1 * pref / wind_speed,
        np.where(
            wind_speed <= 12.2,
            k2 * pref / wind_speed,
            k3 * pref / wind_speed
        )
    )

    # 处理风速为零的情况
    ft_theoretical = np.where(wind_speed == 0, 0.0, ft_theoretical)

    # 计算误差平方和
    thrust_diff_squared = (ft_theoretical - ft_ref) ** 2

    # 返回总误差平方和
    return np.sum(thrust_diff_squared)


# BFGS优化函数
def optimize_k_values(wt_data):
    # 设置初始参数值
    initial_k_values = [1.2, 1.5, 1.3]

    # 设置约束条件 0.5 ≤ k ≤ 3.0
    bounds = [(0.5, 3.0), (0.5, 3.0), (0.5, 3.0)]

    # 使用L-BFGS-B算法进行优化
    result = minimize(objective_function, initial_k_values, args=(wt_data,), method='L-BFGS-B', bounds=bounds)

    if result.success:
        return result.x
    else:
        print("优化未收敛:", result.message)
        return initial_k_values


# 处理WF1 和 WF2 的数据，优化k值并预测塔架推力
results = []

for file_name, data_dict in zip(['WF1', 'WF2'], [wf1_data, wf2_data]):
    for wt in data_dict.keys():
        wt_data = data_dict[wt].loc[:, ['Pref', 'WindSpeed', 'Tshaft', 'Ft']].copy()  # 读取所有时刻的数据

        # 过滤风速为零的数据
        wt_data = wt_data[wt_data['WindSpeed'] > 0].copy()

        # 如果过滤后没有数据，跳过
        if wt_data.empty:
            print(f"{file_name}_{wt} 数据集中无有效风速数据，跳过优化。")
            continue

        # 优化k值
        print(f"优化 {file_name}_{wt} 的 k1, k2, k3 参数...")
        optimized_k_values = optimize_k_values(wt_data)
        print(
            f"{file_name}_{wt} 优化后的参数: k1 = {optimized_k_values[0]:.4f}, k2 = {optimized_k_values[1]:.4f}, k3 = {optimized_k_values[2]:.4f}")

        # 预测塔架推力
        wind_speed = wt_data['WindSpeed'].values
        pref = wt_data['Pref'].values
        ft_theoretical = np.where(
            wind_speed < 10.2,
            optimized_k_values[0] * pref / wind_speed,
            np.where(
                wind_speed <= 12.2,
                optimized_k_values[1] * pref / wind_speed,
                optimized_k_values[2] * pref / wind_speed
            )
        )
        ft_theoretical = np.where(wind_speed == 0, 0.0, ft_theoretical)
        wt_data['Pred_Ft'] = ft_theoretical

        # 计算误差平方和
        ft_error = calculate_error(wt_data['Pred_Ft'].values, wt_data['Ft'].values)

        # 保存结果
        results.append({
            'Wind_Turbine': f"{file_name}_{wt}",
            'k1': optimized_k_values[0],
            'k2': optimized_k_values[1],
            'k3': optimized_k_values[2],
            'Ft_Error_Sum_of_Squares': ft_error
        })

# 转换为DataFrame以便展示
results_df = pd.DataFrame(results)

# 显示误差对比结果
print("Wind Turbine Error Comparison:")
print(results_df)

# 如果需要将结果保存到Excel，可以使用以下代码
# results_df.to_excel('Wind_Turbine_Error_Comparison.xlsx', index=False)
