import matplotlib.pyplot as plt
# 创建主轴扭矩和塔架推力的累积疲劳损伤对比图
plt.figure(figsize=(10, 6))
plt.plot(time_series, wt1_main_shaft_damage, label='Main Shaft Torque Damage (WT1)', color='blue')
plt.plot(time_series, wt1_tower_thrust_damage, label='Tower Thrust Damage (WT1)', color='green')
plt.title('Cumulative Fatigue Damage for Main Shaft Torque and Tower Thrust (WT1)')
plt.xlabel('Time (s)')
plt.ylabel('Cumulative Fatigue Damage')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('/mnt/data/Cumulative_Damage_WT1_Comparison.png')  # 保存图片