import matplotlib.pyplot as plt

# 创建塔架推力的累积疲劳损伤图
plt.figure(figsize=(10, 6))
plt.plot(time_series, wt1_tower_thrust_damage, label='Tower Thrust Damage (WT1)', color='green')
plt.title('Cumulative Fatigue Damage for Tower Thrust (WT1)')
plt.xlabel('Time (s)')
plt.ylabel('Cumulative Fatigue Damage')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('/mnt/data/Tower_Thrust_Cumulative_Damage_WT1.png')  # 保存图片