import matplotlib.pyplot as plt

# 假设我们选择一个风机进行可视化
selected_wind_turbine = results_df['Wind_Turbine'].iloc[0]
selected_data = wf1_data[selected_wind_turbine.split('_')[0]].loc[selected_wind_turbine.split('_')[1], ['Pref', 'WindSpeed', 'Tshaft', 'Ft']].copy()
selected_data = selected_data[selected_data['WindSpeed'] > 0].copy()

# 获取优化后的k值
optimized_k = results_df[results_df['Wind_Turbine'] == selected_wind_turbine][['k1', 'k2', 'k3']].values[0]

# 计算预测的Ft
selected_data['Pred_Ft'] = np.where(
    selected_data['WindSpeed'] < 10.2,
    optimized_k[0] * selected_data['Pref'] / selected_data['WindSpeed'],
    np.where(
        selected_data['WindSpeed'] <= 12.2,
        optimized_k[1] * selected_data['Pref'] / selected_data['WindSpeed'],
        optimized_k[2] * selected_data['Pref'] / selected_data['WindSpeed']
    )
)
selected_data['Pred_Ft'] = np.where(selected_data['WindSpeed'] == 0, 0.0, selected_data['Pred_Ft'])

# 绘制塔架推力散点图
plt.figure(figsize=(12, 6))

plt.subplot(1, 2, 1)
plt.scatter(selected_data['Ft'], selected_data['Pred_Ft'], alpha=0.5, label='Predicted vs Reference')
plt.plot([selected_data['Ft'].min(), selected_data['Ft'].max()],
         [selected_data['Ft'].min(), selected_data['Ft'].max()], 'r--', label='Ideal Fit')
plt.xlabel('Reference Ft')
plt.ylabel('Predicted Ft')
plt.title(f'Scatter Plot for {selected_wind_turbine}')
plt.legend()
plt.grid(True)

# 绘制残差图
residuals_ft = selected_data['Pred_Ft'] - selected_data['Ft']

plt.subplot(1, 2, 2)
plt.scatter(selected_data['WindSpeed'], residuals_ft, alpha=0.5)
plt.axhline(0, color='r', linestyle='--')
plt.xlabel('Wind Speed (m/s)')
plt.ylabel('Residual Ft')
plt.title(f'Residual Plot for {selected_wind_turbine}')
plt.grid(True)

plt.tight_layout()
plt.show()
