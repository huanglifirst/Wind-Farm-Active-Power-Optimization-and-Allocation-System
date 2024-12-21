import matplotlib.pyplot as plt

# 创建主轴扭矩的累积疲劳损伤图
plt.figure(figsize=(10, 6))
plt.plot(time_series, wt1_main_shaft_damage, label='Main Shaft Torque Damage (WT1)', color='blue')
plt.title('Cumulative Fatigue Damage for Main Shaft Torque (WT1)')
plt.xlabel('Time (s)')
plt.ylabel('Cumulative Fatigue Damage')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('/mnt/data/Main_Shaft_Torque_Cumulative_Damage_WT1.png')  # 保存图片





# 关闭所有图像窗口
plt.close('all')
