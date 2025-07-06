import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 定义所有园区的装机容量（没有的设为0）
capacities = {
    'A': {'pv': 750, 'wind': 0},  # 园区A只有光伏
    'B': {'pv': 0, 'wind': 1000},  # 园区B只有风电
    'C': {'pv': 600, 'wind': 500}  # 园区C有光伏和风电
}

# 电价参数
electricity_prices = {
    'pv': 0.4,  # 光伏购电成本 (元/kWh)
    'wind': 0.5,  # 风电购电成本 (元/kWh)
    'grid': 1.0  # 主网购电价格 (元/kWh)
}

# 储能配置参数（50kW/100kWh）
storage_config = {
    'power': 50,  # 充放电功率限制 (kW)
    'capacity': 100,  # 总容量 (kWh)
    'soc_min': 10,  # SOC下限 (%)
    'soc_max': 90,  # SOC上限 (%)
    'efficiency': 0.95,  # 充放电效率
    'power_cost': 800,  # 功率单价 (元/kW)
    'energy_cost': 1800,  # 能量单价 (元/kWh)
    'lifetime': 10  # 运行寿命 (年)
}

# 计算储能投资成本及其日分摊成本
storage_investment = storage_config['power'] * storage_config['power_cost'] + \
                     storage_config['capacity'] * storage_config['energy_cost']
storage_daily_cost = storage_investment / (storage_config['lifetime'] * 365)

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

# 初始化储能状态 (初始SOC设为90%)
storage_soc = {area: 90.0 for area in ['A', 'B', 'C']}  # 单位：%

# 初始化结果列
for area in ['A', 'B', 'C']:
    date[f'{area}_renew_used'] = 0.0  # 可再生能源实际用量（用于负荷）
    date[f'{area}_renew_to_storage'] = 0.0  # 可再生能源充入储能量
    date[f'{area}_storage_discharge'] = 0.0  # 储能放电量（供给负荷）
    date[f'{area}_curtailment'] = 0.0  # 总弃电量
    date[f'{area}_grid_purchase'] = 0.0  # 网购电量
    date[f'{area}_curtail_pv'] = 0.0  # 光伏弃电量
    date[f'{area}_curtail_wind'] = 0.0  # 风电弃电量
    date[f'{area}_soc'] = 0.0  # 储能SOC
    date[f'{area}_pv_used'] = 0.0  # 光伏实际利用量（用于负荷）
    date[f'{area}_wind_used'] = 0.0  # 风电实际利用量（用于负荷）

# 计算每个园区的能源分配
for i, row in date.iterrows():
    for area in ['A', 'B', 'C']:
        load = row[f'园区{area}负荷(kW)']
        pv_gen = row[f'{area}_pv_power']
        wind_gen = row[f'{area}_wind_power']
        total_gen = pv_gen + wind_gen

        # 当前储能状态（转换为kWh）
        soc_pct = storage_soc[area]
        soc_kwh = soc_pct / 100 * storage_config['capacity']

        # 1. 优先使用光伏发电
        pv_used = min(pv_gen, load)
        load_after_pv = load - pv_used

        # 2. 然后使用风电
        wind_used = min(wind_gen, load_after_pv)
        load_remain = load_after_pv - wind_used
        renew_used = pv_used + wind_used

        # 3. 计算剩余可再生能源
        pv_surplus = pv_gen - pv_used
        wind_surplus = wind_gen - wind_used
        total_surplus = pv_surplus + wind_surplus

        # 4. 剩余可再生能源处理
        if total_surplus > 0:
            # 计算最大充电功率 (考虑SOC上限和效率)
            max_charge_kw = min(
                storage_config['power'],
                (storage_config['soc_max'] / 100 * storage_config['capacity'] - soc_kwh) / storage_config['efficiency']
            )
            charge_kw = min(total_surplus, max_charge_kw)

            # 实际充入储能的电量（考虑效率）
            charge_actual = charge_kw * storage_config['efficiency']

            # 更新SOC
            soc_kwh += charge_actual
            renew_to_storage = charge_kw

            # 弃电分配（优先弃风电）
            curtail_total = total_surplus - charge_kw
            wind_curtail = min(wind_surplus, curtail_total)
            pv_curtail = curtail_total - wind_curtail
        else:
            renew_to_storage = 0.0
            charge_actual = 0.0
            pv_curtail = pv_surplus
            wind_curtail = wind_surplus
            curtail_total = pv_curtail + wind_curtail

        # 5. 负荷不足部分由储能补充
        if load_remain > 0:
            # 计算最大放电功率 (考虑SOC下限和效率)
            max_discharge_kw = min(
                storage_config['power'],
                (soc_kwh - storage_config['soc_min'] / 100 * storage_config['capacity']) * storage_config['efficiency']
            )
            discharge_kw = min(load_remain, max_discharge_kw)

            # 实际放出的电量（考虑效率）
            discharge_actual = discharge_kw / storage_config['efficiency']

            # 更新SOC
            soc_kwh -= discharge_actual
            load_remain -= discharge_kw
        else:
            discharge_kw = 0.0
            discharge_actual = 0.0

        # 6. 剩余负荷由电网补充
        grid_purchase = load_remain if load_remain > 0 else 0.0

        # 7. 更新储能状态（百分比）
        storage_soc[area] = soc_kwh / storage_config['capacity'] * 100

        # 8. 记录结果
        date.at[i, f'{area}_renew_used'] = renew_used
        date.at[i, f'{area}_renew_to_storage'] = renew_to_storage
        date.at[i, f'{area}_storage_discharge'] = discharge_kw
        date.at[i, f'{area}_curtail_pv'] = pv_curtail
        date.at[i, f'{area}_curtail_wind'] = wind_curtail
        date.at[i, f'{area}_curtailment'] = curtail_total
        date.at[i, f'{area}_grid_purchase'] = grid_purchase
        date.at[i, f'{area}_soc'] = storage_soc[area]
        date.at[i, f'{area}_pv_used'] = pv_used  # 记录光伏实际利用量
        date.at[i, f'{area}_wind_used'] = wind_used  # 记录风电实际利用量

# 计算总量（按小时累加）
results = {}
for area in ['A', 'B', 'C']:
    total_load = date[f'园区{area}负荷(kW)'].sum()

    # 使用实际利用量计算成本
    pv_used_total = date[f'{area}_pv_used'].sum()
    wind_used_total = date[f'{area}_wind_used'].sum()

    # 可再生能源成本 = 实际利用量 × 成本
    renew_cost = pv_used_total * electricity_prices['pv'] + wind_used_total * electricity_prices['wind']

    # 网购电成本
    grid_cost = date[f'{area}_grid_purchase'].sum() * electricity_prices['grid']

    # 运行成本（不含储能投资）
    total_operation_cost = renew_cost + grid_cost

    # 总成本（含储能投资）
    total_cost = total_operation_cost + storage_daily_cost

    # 单位电量成本
    avg_cost = total_cost / total_load if total_load > 0 else 0

    results[area] = {
        '总负荷电量(kWh)': total_load,
        '弃电量(kWh)': date[f'{area}_curtailment'].sum(),
        '网购电量(kWh)': date[f'{area}_grid_purchase'].sum(),
        '光伏利用量(kWh)': pv_used_total,
        '风电利用量(kWh)': wind_used_total,
        '可再生能源成本(元)': renew_cost,
        '网购电成本(元)': grid_cost,
        '储能日分摊成本(元)': storage_daily_cost,
        '总运行成本(元)': total_operation_cost,
        '总供电成本(元)': total_cost,
        '单位电量成本(元/kWh)': avg_cost,
        '储能充入电量(kWh)': date[f'{area}_renew_to_storage'].sum(),
        '储能供电量(kWh)': date[f'{area}_storage_discharge'].sum(),
        '储能投资成本(元)': storage_investment
    }

# 打印结果
print(f"储能配置: {storage_config['power']}kW/{storage_config['capacity']}kWh")
print(f"储能投资成本: {storage_investment:.2f}元")
print(f"每日分摊成本: {storage_daily_cost:.2f}元/天")

for area, res in results.items():
    print(f"\n园区{area}结果 (配置储能):")
    for k, v in res.items():
        print(f"{k}: {v:.2f}" if isinstance(v, float) else f"{k}: {v}")

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号


# 创建统一的绘图函数
def plot_area_energy(area):
    """绘制园区能源曲线图"""
    plt.figure(figsize=(14, 10))
    hours = range(len(date))

    # 获取当前园区的装机容量
    pv_cap = capacities[area]['pv']
    wind_cap = capacities[area]['wind']

    # 创建子图
    ax1 = plt.subplot(2, 1, 1)

    # 绘制负荷曲线
    ax1.plot(hours, date[f'园区{area}负荷(kW)'], 'k-', label='负荷', linewidth=2.5, zorder=10)

    # 绘制光伏曲线
    if pv_cap > 0:
        ax1.plot(hours, date[f'{area}_pv_power'], color='gold', label='光伏发电', linewidth=2)

    # 绘制风电曲线
    if wind_cap > 0:
        ax1.plot(hours, date[f'{area}_wind_power'], color='royalblue', label='风电发电', linewidth=2)

    # 绘制总发电曲线
    ax1.plot(hours, date[f'{area}_total_power'], 'g--', label='总发电', linewidth=1.8, alpha=0.8)

    # 绘制弃光曲线
    if pv_cap > 0:
        ax1.plot(hours, date[f'{area}_curtail_pv'], 'r--', label='弃光', linewidth=1.8, alpha=0.7)

    # 绘制弃风曲线
    if wind_cap > 0:
        ax1.plot(hours, date[f'{area}_curtail_wind'], 'm--', label='弃风', linewidth=1.8, alpha=0.7)

    # 设置图表属性
    ax1.set_title(f'园区{area} - 能源曲线图', fontsize=18)
    ax1.set_ylabel('功率 (kW)', fontsize=14)
    ax1.legend(loc='upper left', fontsize=12)
    ax1.grid(True, linestyle='--', alpha=0.6)

    # 设置x轴刻度
    ax1.set_xticks(hours)
    ax1.set_xticklabels(date['时间（h）'], rotation=45, fontsize=10)
    ax1.set_xlim(0, len(hours) - 1)

    # 添加储能曲线子图
    ax2 = plt.subplot(2, 1, 2, sharex=ax1)

    # 绘制储能SOC曲线
    ax2.plot(hours, date[f'{area}_soc'], 'b-', label='储能SOC', linewidth=2.5)
    ax2.set_ylabel('SOC (%)', fontsize=14)
    ax2.set_ylim(0, 100)

    # 绘制储能充放电曲线（双Y轴）
    ax3 = ax2.twinx()
    ax3.bar(hours, date[f'{area}_renew_to_storage'], color='g', alpha=0.5, label='储能充电')
    ax3.bar(hours, -date[f'{area}_storage_discharge'], color='r', alpha=0.5, label='储能放电')
    ax3.set_ylabel('充放电功率 (kW)', fontsize=14)

    # 设置图例
    lines, labels = ax2.get_legend_handles_labels()
    bars, bar_labels = ax3.get_legend_handles_labels()
    ax3.legend(lines + bars, labels + bar_labels, loc='upper right', fontsize=12)

    ax2.grid(True, linestyle='--', alpha=0.6)
    ax2.set_xlabel('时间 (小时)', fontsize=14)

    # 添加装机容量说明
    plt.figtext(0.5, 0.01,
                f"注: 光伏装机容量={pv_cap}kW, 风电装机容量={wind_cap}kW, 储能配置={storage_config['power']}kW/{storage_config['capacity']}kWh",
                ha='center', fontsize=10, alpha=0.7)

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])

    # 保存图表
    plt.savefig(f'园区{area}_能源曲线_储能配置.png', dpi=300, bbox_inches='tight')
    plt.show()


# 为所有园区绘制曲线图
for area in ['A', 'B', 'C']:
    plot_area_energy(area)