import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 定义所有园区的装机容量（没有的设为0）
capacities = {
    'A': {'pv': 750, 'wind': 0},  # 园区A只有光伏
    'B': {'pv': 0, 'wind': 1000},  # 园区B只有风电
    'C': {'pv': 600, 'wind': 500}  # 园区C有光伏和风电
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

# 原始数据中没有A_wind和B_pv列，我们添加它们并设为0
date['A_wind'] = 0.0
date['B_pv'] = 0.0

# 为每个园区计算实际发电功率
for area in ['A', 'B', 'C']:
    # 光伏功率 = 装机容量 × 归一化值
    date[f'{area}_pv_power'] = capacities[area]['pv'] * date[f'{area}_pv']
    # 风电功率 = 装机容量 × 归一化值
    date[f'{area}_wind_power'] = capacities[area]['wind'] * date[f'{area}_wind']
    # 总可再生能源功率
    date[f'{area}_total_power'] = date[f'{area}_pv_power'] + date[f'{area}_wind_power']

# 初始化结果列
for area in ['A', 'B', 'C']:
    date[f'{area}_renew_used'] = 0.0  # 可再生能源实际用量
    date[f'{area}_curtailment'] = 0.0  # 总弃电量
    date[f'{area}_grid_purchase'] = 0.0  # 网购电量
    date[f'{area}_curtail_pv'] = 0.0  # 光伏弃电量
    date[f'{area}_curtail_wind'] = 0.0  # 风电弃电量

# 统一计算每个园区的能源分配
for i, row in date.iterrows():
    for area in ['A', 'B', 'C']:
        load = row[f'园区{area}负荷(kW)']
        total_power = row[f'{area}_total_power']
        pv_power = row[f'{area}_pv_power']
        wind_power = row[f'{area}_wind_power']
        if total_power >= load:
            # 可再生能源满足全部负荷
            date.at[i, f'{area}_renew_used'] = load
            total_curtail = total_power - load
            date.at[i, f'{area}_curtailment'] = total_curtail

            # 优先弃风电：先弃风电，风电不足再弃光伏
            wind_curtail = min(wind_power, total_curtail)
            pv_curtail = total_curtail - wind_curtail

            # 记录弃电分配
            date.at[i, f'{area}_curtail_pv'] = pv_curtail
            date.at[i, f'{area}_curtail_wind'] = wind_curtail
        else:
            # 可再生能源不足
            date.at[i, f'{area}_renew_used'] = total_power
            date.at[i, f'{area}_grid_purchase'] = load - total_power

            # 设置弃电量为0
            date.at[i, f'{area}_curtailment'] = 0
            date.at[i, f'{area}_curtail_pv'] = 0
            date.at[i, f'{area}_curtail_wind'] = 0

# 计算总量（按小时累加）
results = {}
for area in ['A', 'B', 'C']:
    total_load = date[f'园区{area}负荷(kW)'].sum()
    total_curtail = date[f'{area}_curtailment'].sum()
    grid_purchase = date[f'{area}_grid_purchase'].sum()

    # 计算可再生能源成本
    pv_used = date[f'{area}_pv_power'] - date[f'{area}_curtail_pv']
    wind_used = date[f'{area}_wind_power'] - date[f'{area}_curtail_wind']
    renew_cost = pv_used.sum() * 0.4 + wind_used.sum() * 0.5

    # 网购电成本
    grid_cost = grid_purchase * 1.0

    total_cost = renew_cost + grid_cost
    avg_cost = total_cost / total_load if total_load > 0 else 0

    results[area] = {
        '总负荷电量(kWh)': total_load,
        '弃电量(kWh)': total_curtail,
        '网购电量(kWh)': grid_purchase,
        '可再生能源成本(元)': renew_cost,
        '网购电成本(元)': grid_cost,
        '总供电成本(元)': total_cost,
        '单位电量成本(元/kWh)': avg_cost
    }

# 打印结果
for area, res in results.items():
    print(f"\n园区{area}结果:")
    for k, v in res.items():
        print(f"{k}: {v:.2f}" if isinstance(v, float) else f"{k}: {v}")

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号


# 创建统一的绘图函数
def plot_area_energy(area):
    """绘制园区能源曲线图"""
    plt.figure(figsize=(14, 8))
    hours = range(len(date))

    # 获取当前园区的装机容量
    pv_cap = capacities[area]['pv']
    wind_cap = capacities[area]['wind']

    # 绘制负荷曲线
    plt.plot(hours, date[f'园区{area}负荷(kW)'], 'k-', label='负荷', linewidth=2.5, zorder=10)

    # 绘制光伏曲线
    plt.plot(hours, date[f'{area}_pv_power'], color='gold', label='光伏发电', linewidth=2)

    # 绘制风电曲线
    plt.plot(hours, date[f'{area}_wind_power'], color='royalblue', label='风电发电', linewidth=2)

    # 绘制总发电曲线
    plt.plot(hours, date[f'{area}_total_power'], 'g--', label='总发电', linewidth=1.8, alpha=0.8)

    # 绘制弃光曲线
    plt.plot(hours, date[f'{area}_curtail_pv'], 'r--', label='弃光', linewidth=1.8, alpha=0.7)

    # 绘制弃风曲线
    plt.plot(hours, date[f'{area}_curtail_wind'], 'm--', label='弃风', linewidth=1.8, alpha=0.7)

    # 设置图表属性
    plt.title(f'园区{area} - 能源曲线图', fontsize=18)
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
    plt.figtext(0.5, 0.01, f"注: 光伏装机容量={pv_cap}kW, 风电装机容量={wind_cap}kW",
                ha='center', fontsize=10, alpha=0.7)

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])

    # 保存图表
    plt.savefig(f'园区{area}_能源曲线.png', dpi=300, bbox_inches='tight')
    plt.show()


# 为所有园区绘制曲线图
for area in ['A', 'B', 'C']:
    plot_area_energy(area)