import pandas as pd
import numpy as np

# 定义所有园区的初始装机容量
initial_capacities = {
    'A': {'pv': 750, 'wind': 0},  # 园区A只有光伏
    'B': {'pv': 0, 'wind': 1000},  # 园区B只有风电
    'C': {'pv': 600, 'wind': 500}  # 园区C有光伏和风电
}

# 成本参数
cost_params = {
    'pv': 2500,  # 光伏单位成本 (元/kW)
    'wind': 3000,  # 风电单位成本 (元/kW)
    'ess_power': 800,  # 储能功率单价 (元/kW)
    'ess_energy': 1800  # 储能能量单价 (元/kWh)
}

# 储能技术参数
ess_params = {
    'soc_min': 10,  # SOC下限 (%)
    'soc_max': 90,  # SOC上限 (%)
    'efficiency': 0.95,  # 充放电效率
}

# 电价参数
electricity_prices = {
    'wind': 0.5,  # 风电购电成本 (元/kWh)
    'pv': 0.4,  # 光伏购电成本 (元/kWh)
    'grid': 1.0  # 主网购电价格 (元/kWh)
}

# 投资回报期 (年)
payback_period = 5

# 读取负荷数据
load_data = pd.read_excel('C:/Users/HP/Desktop/附件1：各园区典型日负荷数据.xlsx')
# 读取风光数据
renewable_data = pd.read_excel('C:/Users/HP/Desktop/附件2：各园区典型日风光发电数据.xlsx', skiprows=2, header=None,
                               names=['时间', 'A_pv', 'B_wind', 'C_pv', 'C_wind'])

# 确保时间列对齐
load_data['时间（h）'] = load_data['时间（h）'].astype(str)
renewable_data['时间'] = renewable_data['时间'].astype(str)

# 合并数据
data = pd.merge(load_data, renewable_data, left_on='时间（h）', right_on='时间', how='inner')

# 最大负荷增长50%
for area in ['A', 'B', 'C']:
    data[f'园区{area}负荷(kW)'] = data[f'园区{area}负荷(kW)'] * 1.5

# 计算总负荷
data['总负荷(kW)'] = data['园区A负荷(kW)'] + data['园区B负荷(kW)'] + data['园区C负荷(kW)']

# 计算日总负荷电量 (kWh)
daily_load_total = data['总负荷(kW)'].sum()

# 初始总容量
initial_total_pv = initial_capacities['A']['pv'] + initial_capacities['C']['pv']
initial_total_wind = initial_capacities['B']['wind'] + initial_capacities['C']['wind']


# 为联合园区优化风光储配置
def optimize_joint():
    """为联合园区优化风光储配置"""
    # 确定循环范围
    pv_range = range(max(0, initial_total_pv - 500), initial_total_pv + 1000, 200)
    wind_range = range(max(0, initial_total_wind - 500), initial_total_wind + 1000, 200)
    ess_power_range = range(0, 501, 100)
    ess_capacity_range = range(0, 1001, 200)

    best_cost = float('inf')
    best_config = {}
    best_results = {}

    # 遍历所有可能的配置组合
    for pv_cap in pv_range:
        for wind_cap in wind_range:
            # 跳过无效组合
            if pv_cap == 0 and wind_cap == 0:
                continue

            for ess_power in ess_power_range:
                for ess_capacity in ess_capacity_range:
                    # 跳过无效储能配置
                    if ess_power > 0 and ess_capacity == 0:
                        continue

                    # 计算投资成本
                    investment_cost = (pv_cap * cost_params['pv'] +
                                       wind_cap * cost_params['wind'] +
                                       ess_power * cost_params['ess_power'] +
                                       ess_capacity * cost_params['ess_energy'])

                    # 初始化储能状态
                    storage_soc = 90.0  # 初始SOC 90%

                    # 初始化运行结果
                    total_pv_used = 0
                    total_wind_used = 0
                    grid_purchase = 0
                    pv_curtail = 0
                    wind_curtail = 0

                    # 模拟24小时运行
                    for _, row in data.iterrows():
                        # 计算总负荷
                        total_load = row['总负荷(kW)']

                        # 计算各园区可再生能源出力
                        pv_gen_a = pv_cap * (initial_capacities['A']['pv'] / initial_total_pv) * row['A_pv']
                        pv_gen_c = pv_cap * (initial_capacities['C']['pv'] / initial_total_pv) * row['C_pv']
                        wind_gen_b = wind_cap * (initial_capacities['B']['wind'] / initial_total_wind) * row['B_wind']
                        wind_gen_c = wind_cap * (initial_capacities['C']['wind'] / initial_total_wind) * row['C_wind']

                        total_pv_gen = pv_gen_a + pv_gen_c
                        total_wind_gen = wind_gen_b + wind_gen_c

                        # 当前储能状态 (kWh)
                        soc_kwh = storage_soc / 100 * ess_capacity if ess_capacity > 0 else 0

                        # 1. 优先使用光伏发电
                        pv_used_hr = min(total_pv_gen, total_load)
                        load_after_pv = total_load - pv_used_hr

                        # 2. 然后使用风电
                        wind_used_hr = min(total_wind_gen, load_after_pv)
                        load_remain = load_after_pv - wind_used_hr

                        # 更新可再生能源利用量
                        total_pv_used += pv_used_hr
                        total_wind_used += wind_used_hr

                        # 3. 计算剩余可再生能源
                        pv_surplus = total_pv_gen - pv_used_hr
                        wind_surplus = total_wind_gen - wind_used_hr
                        total_surplus = pv_surplus + wind_surplus

                        # 4. 如果有剩余可再生能源，尝试存入储能
                        if total_surplus > 0 and ess_capacity > 0:
                            # 计算最大充电功率
                            soc_max_kwh = ess_params['soc_max'] / 100 * ess_capacity
                            max_charge_kw = min(
                                ess_power,
                                (soc_max_kwh - soc_kwh) / ess_params['efficiency']
                            )
                            charge_kw = min(total_surplus, max_charge_kw)
                            charge_actual = charge_kw * ess_params['efficiency']

                            # 更新SOC
                            soc_kwh += charge_actual

                            # 计算弃电（优先弃风）
                            curtail_total = total_surplus - charge_kw
                            wind_curtail_hr = min(wind_surplus, curtail_total)
                            pv_curtail_hr = curtail_total - wind_curtail_hr

                            # 累加弃电量
                            wind_curtail += wind_curtail_hr
                            pv_curtail += pv_curtail_hr
                        else:
                            # 没有储能或没有剩余可再生能源时，全部弃电
                            wind_curtail += wind_surplus
                            pv_curtail += pv_surplus

                        # 5. 如果负荷仍有剩余，尝试从储能放电
                        if load_remain > 0 and ess_capacity > 0:
                            # 计算最大放电功率
                            soc_min_kwh = ess_params['soc_min'] / 100 * ess_capacity
                            max_discharge_kw = min(
                                ess_power,
                                (soc_kwh - soc_min_kwh) * ess_params['efficiency']
                            )
                            discharge_kw = min(load_remain, max_discharge_kw)
                            discharge_actual = discharge_kw / ess_params['efficiency']

                            # 更新SOC
                            soc_kwh -= discharge_actual
                            load_remain -= discharge_kw

                        # 6. 剩余负荷由电网补充
                        grid_purchase += max(0, load_remain)

                        # 7. 更新储能状态
                        if ess_capacity > 0:
                            storage_soc = soc_kwh / ess_capacity * 100
                            # 确保SOC在范围内
                            storage_soc = max(ess_params['soc_min'], min(ess_params['soc_max'], storage_soc))

                    # 计算运行成本
                    renew_cost = (total_pv_used * electricity_prices['pv'] +
                                  total_wind_used * electricity_prices['wind'])
                    grid_cost = grid_purchase * electricity_prices['grid']
                    daily_operation_cost = renew_cost + grid_cost

                    # 计算5年总成本
                    total_cost = investment_cost + daily_operation_cost * 365 * payback_period

                    # 计算单位电量成本 (元/kWh)
                    total_energy_supplied = daily_load_total * 365 * payback_period
                    cost_per_kwh = total_cost / total_energy_supplied if total_energy_supplied > 0 else 0

                    # 更新最优配置
                    if total_cost < best_cost:
                        best_cost = total_cost
                        best_config = {
                            'pv_capacity': pv_cap,
                            'wind_capacity': wind_cap,
                            'ess_power': ess_power,
                            'ess_capacity': ess_capacity
                        }
                        best_results = {
                            'investment_cost': investment_cost,
                            'daily_operation_cost': daily_operation_cost,
                            'total_cost': total_cost,
                            'daily_pv_used': total_pv_used,
                            'daily_wind_used': total_wind_used,
                            'daily_renew_used': total_pv_used + total_wind_used,
                            'daily_grid_purchase': grid_purchase,
                            'daily_pv_curtail': pv_curtail,
                            'daily_wind_curtail': wind_curtail,
                            'cost_per_kwh': cost_per_kwh
                        }

    return best_config, best_results


# 优化联合园区配置
config, res = optimize_joint()

# 打印结果
print(f"光伏总容量: {config['pv_capacity']} kW")
print(f"风电总容量: {config['wind_capacity']} kW")
print(f"储能总功率: {config['ess_power']} kW")
print(f"储能总容量: {config['ess_capacity']} kWh")
print(f"投资成本: {res['investment_cost']:.2f} 元")
print(f"日运行成本: {res['daily_operation_cost']:.2f} 元")
print(f"总成本(5年): {res['total_cost']:.2f} 元")
print(f"日光伏利用量: {res['daily_pv_used']:.2f} kWh")
print(f"日风电利用量: {res['daily_wind_used']:.2f} kWh")
print(f"日可再生能源利用量: {res['daily_renew_used']:.2f} kWh")
print(f"日网购电量: {res['daily_grid_purchase']:.2f} kWh")
print(f"日弃光电量: {res['daily_pv_curtail']:.2f} kWh")
print(f"日弃风电量: {res['daily_wind_curtail']:.2f} kWh")
print(f"单位电量成本: {res['cost_per_kwh']:.4f} 元/kWh")

# 保存结果
output_data = [{
    '光伏总容量(kW)': config['pv_capacity'],
    '风电总容量(kW)': config['wind_capacity'],
    '储能总功率(kW)': config['ess_power'],
    '储能总容量(kWh)': config['ess_capacity'],
    '投资成本(元)': res['investment_cost'],
    '日运行成本(元)': res['daily_operation_cost'],
    '总成本(元)': res['total_cost'],
    '日光伏利用量(kWh)': res['daily_pv_used'],
    '日风电利用量(kWh)': res['daily_wind_used'],
    '日可再生能源利用量(kWh)': res['daily_renew_used'],
    '日网购电量(kWh)': res['daily_grid_purchase'],
    '日弃光电量(kWh)': res['daily_pv_curtail'],
    '日弃风电量(kWh)': res['daily_wind_curtail'],
    '单位电量成本(元/kWh)': res['cost_per_kwh']
}]

output_df = pd.DataFrame(output_data)
output_df.to_excel('联合运营风光储配置结果.xlsx', index=False)
print("\n结果已保存到 '联合运营风光储配置结果.xlsx'")