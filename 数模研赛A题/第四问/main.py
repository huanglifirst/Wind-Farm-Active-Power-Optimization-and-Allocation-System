import numpy as np
import pandas as pd
from scipy.optimize import minimize
import matplotlib
import os


matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False

eta = 0.95  # 机械效率
alpha = 0.1  # 系统响应参数 (Nm·s/W)
gamma_V = 0.1  # 测量噪声水平 (10%)
delta_t_max = 10  # 最大通信延迟 (秒)
P_max = 5  # 风机额定功率 (MW)
rho = 1.225  # 空气密度 (kg/m^3)
C_T = 0.7  # 推力系数
R = 64  # 风轮半径 (米)
A = np.pi * R**2  # 风轮扫掠面积 (m^2)
lambda_1 = 0.5  # 主轴疲劳损伤权重
lambda_2 = 0.5  # 塔架疲劳损伤权重
k1 = 1e-11  # 主轴疲劳损伤系数 (小于1e-10)
k2 = 1e-11  # 塔架疲劳损伤系数 (小于1e-10)

# 数据目录和文件名
data_dir = '.'
attachment4_file = '附件4-噪声和延迟作用下的采集数据.xlsx'
attachment4_data = pd.read_excel(os.path.join(data_dir, attachment4_file), sheet_name=None)

# 提取数据
Ft_data = attachment4_data['Ft'].values
Pout_data = attachment4_data['Pout'].values
Pref_data = attachment4_data['Pref'].values
Vwin_data = attachment4_data['Vwin'].values
wgenW_data = attachment4_data['wgenM'].values

Pref_data = Pref_data.T
Vwin_data = Vwin_data.T
wgenW_data = wgenW_data.T

# 初始化参数
num_turbines = 10
time_steps = Pref_data.shape[1]
P_t = np.sum(Pref_data, axis=0)

# 初始化数据结构
P_ref = Pref_data.copy()
V_w = Vwin_data.copy()
omega_f = wgenW_data.copy()
T_shaft = np.zeros((num_turbines, time_steps))
F_tower = np.zeros((num_turbines, time_steps))
D_s = np.zeros((num_turbines, time_steps))
D_t = np.zeros((num_turbines, time_steps))
D_total = np.zeros(time_steps)

# 数据预处理：处理测量噪声和通信延迟
def preprocess_data(V_w_raw):
    # 应用指数平滑法降低噪声影响
    V_w_filtered = np.zeros_like(V_w_raw)
    V_w_filtered[:, 0] = V_w_raw[:, 0]
    alpha_smoothing = 0.5
    for t in range(1, V_w_raw.shape[1]):
        delayed_indices = np.isnan(V_w_raw[:, t])
        V_w_filtered[delayed_indices, t] = V_w_filtered[delayed_indices, t - 1]
        V_w_filtered[~delayed_indices, t] = alpha_smoothing * V_w_raw[~delayed_indices, t] + \
                                            (1 - alpha_smoothing) * V_w_filtered[~delayed_indices, t - 1]
    return V_w_filtered

# 处理数据中的NaN值，将缺失值填充为前一个有效值
def fill_nan_with_previous(data):
    data_filled = data.copy()
    for i in range(data.shape[0]):
        for t in range(1, data.shape[1]):
            if np.isnan(data_filled[i, t]):
                data_filled[i, t] = data_filled[i, t - 1]
    return data_filled

V_w = fill_nan_with_previous(V_w)
V_w_estimated = preprocess_data(V_w)
omega_f = fill_nan_with_previous(omega_f)

for t in range(time_steps):
    # 打印 P_t[t] 和 P_ref_avg
    P_ref_avg = P_t[t] / num_turbines
    print(f"Time {t}, P_t[t] = {P_t[t]:.4f}, P_ref_avg = {P_ref_avg:.4f}")

    # 定义优化问题
    def objective(P_ref_t):
        # 计算每台风机的主轴扭矩和塔架推力
        T_shaft_t = np.zeros(num_turbines)
        F_tower_t = np.zeros(num_turbines)
        D_s_t = np.zeros(num_turbines)
        D_t_t = np.zeros(num_turbines)
        for i in range(num_turbines):
            # 获取omega_f
            omega_f_i_t = omega_f[i, t]
            if t == 0:
                omega_f_i_prev = omega_f[i, t]
                P_ref_i_prev = P_ref_t[i]
            else:
                omega_f_i_prev = omega_f[i, t - 1]
                P_ref_i_prev = P_ref[i, t - 1]
            # 主轴扭矩计算
            T_shaft_i_t = eta * omega_f_i_t * P_ref_t[i] + \
                          alpha * (omega_f_i_t * P_ref_t[i] - omega_f_i_prev * P_ref_i_prev)
            T_shaft_t[i] = T_shaft_i_t
            # 考虑不确定性的最差情况下的风速
            V_w_i_t = V_w_estimated[i, t] * (1 + gamma_V)
            # 塔架推力计算
            F_tower_i_t = 0.5 * rho * C_T * A * V_w_i_t**2 * (P_ref_t[i] / P_max)
            F_tower_t[i] = F_tower_i_t
            # 疲劳损伤计算
            D_s_i_t = k1 * T_shaft_i_t**2
            D_t_i_t = k2 * F_tower_i_t**2
            D_s_t[i] = D_s_i_t
            D_t_t[i] = D_t_i_t
        # 总疲劳损伤
        D_total_t = np.sum(lambda_1 * D_s_t + lambda_2 * D_t_t)
        return D_total_t

    constraints = []
    # 功率平衡约束
    def power_balance_constraint(P_ref_t):
        return np.sum(P_ref_t) - P_t[t]
    constraints.append({'type': 'eq', 'fun': power_balance_constraint})

    bounds = []
    deviation_limit = 2.0  # 放宽偏差限制 (MW)
    for i in range(num_turbines):
        # 功率限制
        bounds.append((0, P_max))
        def deviation_constraint_upper(P_ref_t, i=i):
            return deviation_limit - (P_ref_t[i] - P_ref_avg)
        def deviation_constraint_lower(P_ref_t, i=i):
            return deviation_limit - (P_ref_avg - P_ref_t[i])
        constraints.append({'type': 'ineq', 'fun': deviation_constraint_upper})
        constraints.append({'type': 'ineq', 'fun': deviation_constraint_lower})

    # 初始猜测
    if t == 0:
        P_ref_init = np.full(num_turbines, P_ref_avg)
    else:
        P_ref_init = P_ref[:, t - 1]

    # 优化求解器
    result = minimize(objective, P_ref_init, method='SLSQP', bounds=bounds, constraints=constraints, options={'ftol':1e-6, 'disp': True, 'maxiter':1000})
    if result.success:
        P_ref_opt = result.x
    else:
        print(f"优化在时间 {t} 失败: {result.message}")
        P_ref_opt = P_ref_init

    # 存储优化结果
    P_ref[:, t] = P_ref_opt

    # 更新疲劳损伤和其他变量
    for i in range(num_turbines):
        omega_f_i_t = omega_f[i, t]
        if t == 0:
            omega_f_i_prev = omega_f[i, t]
            P_ref_i_prev = P_ref[i, t]
        else:
            omega_f_i_prev = omega_f[i, t - 1]
            P_ref_i_prev = P_ref[i, t - 1]
        # 主轴扭矩计算
        T_shaft_i_t = eta * omega_f_i_t * P_ref[i, t] + \
                      alpha * (omega_f_i_t * P_ref[i, t] - omega_f_i_prev * P_ref_i_prev)
        T_shaft[i, t] = T_shaft_i_t
        # 塔架推力计算
        V_w_i_t = V_w_estimated[i, t]
        F_tower_i_t = 0.5 * rho * C_T * A * V_w_i_t**2 * (P_ref[i, t] / P_max)
        F_tower[i, t] = F_tower_i_t
        # 疲劳损伤计算
        D_s_i_t = k1 * T_shaft_i_t**2
        D_t_i_t = k2 * F_tower_i_t**2
        D_s[i, t] = D_s_i_t
        D_t[i, t] = D_t_i_t
    # 时间 t 的总疲劳损伤
    D_total[t] = np.sum(lambda_1 * D_s[:, t] + lambda_2 * D_t[:, t])

# 结果展示和保存
import matplotlib.pyplot as plt

# 绘制累积疲劳损伤
plt.figure(figsize=(10, 6))
plt.plot(range(time_steps), np.cumsum(D_total), label='优化后累积疲劳损伤')
plt.xlabel('时间 (秒)')
plt.ylabel('累积疲劳损伤')
plt.title('累积疲劳损伤随时间的变化')
plt.legend()
plt.grid(True)
plt.show()

# 示例风机的功率参考值
turbine_sample = 0
plt.figure(figsize=(10, 6))
plt.plot(range(time_steps), P_ref[turbine_sample, :], label=f'风机 {turbine_sample + 1} 的功率参考值')
plt.xlabel('时间 (秒)')
plt.ylabel('功率参考值 (MW)')
plt.title(f'风机 {turbine_sample + 1} 的功率参考值随时间的变化')
plt.legend()
plt.grid(True)
plt.show()

# 保存结果到CSV文件
output_dir = 'OptimizationResults'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 保存每台风机的优化功率参考值
for i in range(num_turbines):
    turbine_id = i + 1
    df_output = pd.DataFrame({
        'Time': range(time_steps),
        'P_ref_optimized': P_ref[i, :]
    })
    df_output.to_csv(os.path.join(output_dir, f'Turbine_{turbine_id}_P_ref.csv'), index=False)

# 保存累积疲劳损伤
df_damage = pd.DataFrame({
    'Time': range(time_steps),
    'D_total_cumulative': np.cumsum(D_total)
})
df_damage.to_csv(os.path.join(output_dir, 'Cumulative_Fatigue_Damage.csv'), index=False)

print("优化完成。结果已保存到 'OptimizationResults' 目录中。")
