import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 定义联合园区总装机容量
total_capacities = {
    'pv': 1350,  # 总光伏装机容量(kW)
    'wind': 1500  # 总风电装机容量(kW)
}

# 电价参数
electricity_prices = {
    'pv': 0.4,  # 光伏购电成本 (元/kWh)
    'wind': 0.5,  # 风电购电成本 (元/kWh)
    'grid': 1.0  # 主网购电价格 (元/kWh)
}

# 读取负荷数据
load_data = pd.read_excel('C:/Users/HP/Desktop/附件1：各园区典型日负荷数据.xlsx')
# 读取风光数据（跳过前2行非数据行）
renewable_data = pd.read_excel('C:/Users/HP/Desktop/附件2：各园区典型日风光发电数据.xlsx', skiprows=2, header=None,
                               names=['时间', 'A_pv', 'B_wind', 'C_pv', 'C_wind'])

# 确保时间列对齐
load_data['时间（h）'] = load_data['时间（h）'].astype(str)
renewable_data['时间'] = renewable_data['时间'].astype(str)

# 合并数据
date = pd.merge(load_data, renewable_data, left_on='时间（h）', right_on='时间', how='inner')

# 计算联合园区总量
date['总负荷(kW)'] = date['园区A负荷(kW)'] + date['园区B负荷(kW)'] + date['园区C负荷(kW)']
date['总光伏(kW)'] = date['A_pv'] * total_capacities['pv']  # 光伏归一化值乘以总装机容量
date['总风电(kW)'] = date['B_wind'] * total_capacities['wind']  # 风电归一化值乘以总装机容量
date['总发电(kW)'] = date['总光伏(kW)'] + date['总风电(kW)']

# 初始化联合园区结果列
date['总弃电量(kW)'] = 0.0
date['总网购电量(kW)'] = 0.0
date['弃光(kW)'] = 0.0
date['弃风(kW)'] = 0.0
date['光伏利用量(kW)'] = 0.0
date['风电利用量(kW)'] = 0.0

# 计算每个时刻的能源分配（优先使用光伏策略）
for i, row in date.iterrows():
    total_load = row['总负荷(kW)']
    pv_gen = row['总光伏(kW)']
    wind_gen = row['总风电(kW)']

    # 1. 优先使用光伏发电
    pv_used = min(pv_gen, total_load)
    load_after_pv = total_load - pv_used

    # 2. 然后使用风电
    wind_used = min(wind_gen, load_after_pv)
    load_remain = load_after_pv - wind_used

    # 3. 剩余负荷需要网购
    grid_purchase = max(0, load_remain)

    # 4. 计算弃电
    pv_curtail = pv_gen - pv_used
    wind_curtail = wind_gen - wind_used

    # 记录结果
    date.at[i, '光伏利用量(kW)'] = pv_used
    date.at[i, '风电利用量(kW)'] = wind_used
    date.at[i, '弃光(kW)'] = pv_curtail
    date.at[i, '弃风(kW)'] = wind_curtail
    date.at[i, '总弃电量(kW)'] = pv_curtail + wind_curtail
    date.at[i, '总网购电量(kW)'] = grid_purchase

# 计算总量（按小时累加）
total_load_energy = date['总负荷(kW)'].sum()  # 总负荷电量(kWh)
total_curtailment = date['总弃电量(kW)'].sum()  # 总弃电量(kWh)
total_grid_purchase = date['总网购电量(kW)'].sum()  # 总网购电量(kWh)
total_pv_used = date['光伏利用量(kW)'].sum()  # 光伏利用电量(kWh)
total_wind_used = date['风电利用量(kW)'].sum()  # 风电利用电量(kWh)

# 计算成本
renewable_cost = total_pv_used * electricity_prices['pv'] + total_wind_used * electricity_prices['wind']  # 可再生能源成本(元)
grid_cost = total_grid_purchase * electricity_prices['grid']  # 网购电成本(元)
total_cost = renewable_cost + grid_cost  # 总供电成本(元)
avg_cost = total_cost / total_load_energy if total_load_energy > 0 else 0  # 单位电量成本(元/kWh)

# 打印联合园区结果
print("\n联合园区运行经济性分析结果:")
print(f"总负荷电量(kWh): {total_load_energy:.2f}")
print(f"总弃风弃光电量(kWh): {total_curtailment:.2f}")
print(f"总网购电量(kWh): {total_grid_purchase:.2f}")
print(f"光伏利用量(kWh): {total_pv_used:.2f}")
print(f"风电利用量(kWh): {total_wind_used:.2f}")
print(f"可再生能源成本(元): {renewable_cost:.2f}")
print(f"网购电成本(元): {grid_cost:.2f}")
print(f"总供电成本(元): {total_cost:.2f}")
print(f"单位电量平均供电成本(元/kWh): {avg_cost:.4f}")

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 绘制联合园区能源曲线图
plt.figure(figsize=(14, 8))
hours = range(len(date))

# 绘制负荷曲线
plt.plot(hours, date['总负荷(kW)'], 'k-', label='总负荷', linewidth=2.5, zorder=10)

# 绘制光伏发电曲线
plt.plot(hours, date['总光伏(kW)'], color='gold', label='光伏发电', linewidth=2)

# 绘制风电发电曲线
plt.plot(hours, date['总风电(kW)'], color='royalblue', label='风电发电', linewidth=2)

# 绘制总发电曲线
plt.plot(hours, date['总发电(kW)'], 'g--', label='总发电', linewidth=1.8, alpha=0.8)

# 绘制弃光曲线
plt.plot(hours, date['弃光(kW)'], 'r--', label='弃光', linewidth=1.8, alpha=0.7)

# 绘制弃风曲线
plt.plot(hours, date['弃风(kW)'], 'm--', label='弃风', linewidth=1.8, alpha=0.7)

# 设置图表属性
plt.title('联合园区 - 能源曲线图', fontsize=18)
plt.xlabel('时间 (小时)', fontsize=14)
plt.ylabel('功率 (kW)', fontsize=14)

# 添加图例
plt.legend(loc='best', fontsize=12)

# 设置网格
plt.grid(True, linestyle='--', alpha=0.6)

# 设置x轴刻度
plt.xticks(hours, date['时间（h）'], rotation=45, fontsize=10)
plt.xlim(0, len(hours) - 1)

# 添加装机容量说明
plt.figtext(0.5, 0.01, f"注: 光伏总装机容量={total_capacities['pv']}kW, 风电总装机容量={total_capacities['wind']}kW",
            ha='center', fontsize=10, alpha=0.7)

plt.tight_layout(rect=[0, 0.03, 1, 0.95])

# 保存图表
plt.savefig('联合园区_能源曲线.png', dpi=300, bbox_inches='tight')
plt.show()