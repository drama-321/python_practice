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
    'efficiency': 0.95,  # 充放电效率,
}

# 电价参数
electricity_prices = {
    'wind': 0.5,  # 风电购电成本 (元/kWh)
    'pv': 0.4,  # 光伏购电成本 (元/kWh)
}


# 分时电价函数
def get_grid_price(hour):
    """根据小时返回网购电价"""
    if 7 <= hour <= 22:  # 高峰时段
        return 1.0
    else:  # 低谷时段
        return 0.4


# 投资回报期 (年)
payback_period = 5

# 每月天数（平年）
month_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

# 读取负荷数据
load_data = pd.read_excel('C:/Users/HP/Desktop/附件1：各园区典型日负荷数据.xlsx')

# 最大负荷增长50%
for area in ['A', 'B', 'C']:
    load_data[f'园区{area}负荷(kW)'] = load_data[f'园区{area}负荷(kW)'] * 1.5

# 读取全年12个月风光数据
renewable_data = pd.read_excel('C:/Users/HP/Desktop/附件3：12个月各园区典型日风光发电数据_原.xlsx', skiprows=3,
                               header=None)

# 为每个园区构建12个月的风光数据
area_data = {}
areas = ['A', 'B', 'C']

# 处理12个月的数据
for area in areas:
    area_data[area] = {'pv': np.zeros((12, 24)), 'wind': np.zeros((12, 24))}

for month in range(12):
    column = 1 + month * 4  # 每月4列数据

    # 园区A光伏 (第1列)
    area_data['A']['pv'][month] = pd.to_numeric(
        renewable_data.iloc[1:25, column], errors='coerce'
    ).fillna(0).values

    # 园区B风电 (第2列)
    area_data['B']['wind'][month] = pd.to_numeric(
        renewable_data.iloc[1:25, column + 1], errors='coerce'
    ).fillna(0).values

    # 园区C风电 (第3列)
    area_data['C']['wind'][month] = pd.to_numeric(
        renewable_data.iloc[1:25, column + 2], errors='coerce'
    ).fillna(0).values

    # 园区C光伏 (第4列)
    area_data['C']['pv'][month] = pd.to_numeric(
        renewable_data.iloc[1:25, column + 3], errors='coerce'
    ).fillna(0).values


# 优化函数
def optimize_area_full_year(area):
    """为指定园区优化风光储配置（全年分时电价）"""
    # 获取初始容量
    pv_cap_init = initial_capacities[area]['pv']
    wind_cap_init = initial_capacities[area]['wind']

    # 确定循环范围
    pv_options = [max(0, pv_cap_init - 100), pv_cap_init, pv_cap_init + 100] if pv_cap_init > 0 else [0]
    wind_options = [max(0, wind_cap_init - 100), wind_cap_init, wind_cap_init + 100] if wind_cap_init > 0 else [0]
    ess_power_options = [0, 50, 100]
    ess_capacity_options = [0, 100, 200]

    best_cost = float('inf')
    best_config = {}
    best_results = {}

    # 预计算该园区的日负荷总和
    daily_load_total = load_data[f'园区{area}负荷(kW)'].sum()

    # 遍历所有可能的配置组合
    for pv_cap in pv_options:
        for wind_cap in wind_options:
            # 跳过无效组合
            if pv_cap == 0 and wind_cap == 0:
                continue

            for ess_power in ess_power_options:
                for ess_capacity in ess_capacity_options:
                    # 跳过无效储能配置
                    if ess_power > 0 and ess_capacity == 0:
                        continue

                    # 计算投资成本
                    investment_cost = (
                            pv_cap * cost_params['pv'] +
                            wind_cap * cost_params['wind'] +
                            ess_power * cost_params['ess_power'] +
                            ess_capacity * cost_params['ess_energy']
                    )

                    # 初始化运行结果
                    total_renew_used = 0
                    total_pv_used = 0
                    total_wind_used = 0
                    total_grid_cost = 0
                    total_pv_curtail = 0
                    total_wind_curtail = 0

                    # 全年模拟
                    for month in range(12):
                        # 初始化储能状态（每月第一天从90%开始）
                        storage_soc = 90.0

                        # 获取当月的典型日风光数据
                        if area == 'A':
                            pv_data = area_data[area]['pv'][month]
                            wind_data = np.zeros(24)
                        elif area == 'B':
                            pv_data = np.zeros(24)
                            wind_data = area_data[area]['wind'][month]
                        else:  # 园区C
                            pv_data = area_data[area]['pv'][month]
                            wind_data = area_data[area]['wind'][month]

                        # 模拟该月每天运行
                        for day in range(month_days[month]):
                            # 模拟24小时运行
                            for hour in range(24):
                                # 获取当前小时负荷
                                load = load_data.iloc[hour][f'园区{area}负荷(kW)']

                                # 计算可再生能源出力
                                pv_gen = pv_cap * pv_data[hour]
                                wind_gen = wind_cap * wind_data[hour]

                                # 当前储能状态 (kWh)
                                soc_kwh = storage_soc / 100 * ess_capacity if ess_capacity > 0 else 0

                                # 获取当前电价
                                grid_price = get_grid_price(hour)

                                # 1. 优先使用光伏发电
                                pv_used_hr = min(pv_gen, load)
                                load_after_pv = load - pv_used_hr

                                # 2. 然后使用风电
                                wind_used_hr = min(wind_gen, load_after_pv)
                                load_remain = load_after_pv - wind_used_hr

                                # 更新可再生能源利用量
                                total_renew_used += pv_used_hr + wind_used_hr
                                total_pv_used += pv_used_hr
                                total_wind_used += wind_used_hr

                                # 3. 计算剩余可再生能源
                                pv_surplus = pv_gen - pv_used_hr
                                wind_surplus = wind_gen - wind_used_hr
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

                                    # 计算弃电（优先弃风，因为风电成本更高）
                                    curtail_total = total_surplus - charge_kw
                                    wind_curtail_hr = min(wind_surplus, curtail_total)
                                    pv_curtail_hr = curtail_total - wind_curtail_hr

                                    total_pv_curtail += pv_curtail_hr
                                    total_wind_curtail += wind_curtail_hr
                                else:
                                    # 没有储能或没有剩余可再生能源时，全部弃电
                                    wind_curtail_hr = wind_surplus
                                    pv_curtail_hr = pv_surplus
                                    total_pv_curtail += pv_curtail_hr
                                    total_wind_curtail += wind_curtail_hr

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
                                grid_cost_hr = load_remain * grid_price
                                total_grid_cost += grid_cost_hr

                                # 7. 在低谷时段，如果有储能容量，从电网充电，以0.4的价格购入，即使考虑充电效率成本也只是0.49
                                if grid_price == 0.4 and ess_capacity > 0:
                                    # 计算最大充电功率
                                    soc_max_kwh = ess_params['soc_max'] / 100 * ess_capacity
                                    max_charge_kw = min(
                                        ess_power,
                                        (soc_max_kwh - soc_kwh) / ess_params['efficiency']
                                    )
                                    if max_charge_kw > 0:
                                        charge_kw = max_charge_kw
                                        charge_actual = charge_kw * ess_params['efficiency']

                                        # 更新SOC
                                        soc_kwh += charge_actual
                                        total_grid_cost += charge_kw * grid_price

                                # 更新储能状态
                                if ess_capacity > 0:
                                    storage_soc = soc_kwh / ess_capacity * 100
                                    # 确保SOC在范围内
                                    storage_soc = max(ess_params['soc_min'], min(ess_params['soc_max'], storage_soc))

                    # 计算可再生能源成本（按实际使用量）
                    renew_cost = (total_pv_used * electricity_prices['pv'] +
                                  total_wind_used * electricity_prices['wind'])

                    # 总运行成本（全年）
                    annual_operation_cost = renew_cost + total_grid_cost

                    # 计算5年总成本
                    total_cost = investment_cost + annual_operation_cost * payback_period

                    # 计算单位电量成本 (元/kWh)
                    total_energy_supplied = daily_load_total * 365 * payback_period
                    if total_energy_supplied > 0:
                        cost_per_kwh = total_cost / total_energy_supplied
                    else:
                        cost_per_kwh = 0

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
                            'annual_operation_cost': annual_operation_cost,
                            'total_cost': total_cost,
                            'annual_renew_used': total_renew_used,
                            'annual_pv_used': total_pv_used,
                            'annual_wind_used': total_wind_used,
                            'annual_grid_cost': total_grid_cost,
                            'annual_pv_curtail': total_pv_curtail,
                            'annual_wind_curtail': total_wind_curtail,
                            'cost_per_kwh': cost_per_kwh
                        }

    return best_config, best_results


# 主程序
if __name__ == "__main__":
    results = {}

    # 优化所有三个园区
    for area in ['A', 'B', 'C']:
        config, res = optimize_area_full_year(area)
        results[area] = {'config': config, 'results': res}

    for area in ['A', 'B', 'C']:
        config = results[area]['config']
        res = results[area]['results']
        print(f"\n园区{area}最优配置:")
        print(f"  光伏容量: {config['pv_capacity']} kW")
        print(f"  风电容量: {config['wind_capacity']} kW")
        print(f"  储能功率: {config['ess_power']} kW")
        print(f"  储能容量: {config['ess_capacity']} kWh")
        print(f"  投资成本: {res['investment_cost']:.2f} 元")
        print(f"  年运行成本: {res['annual_operation_cost']:.2f} 元")
        print(f"  总成本(5年): {res['total_cost']:.2f} 元")
        print(f"  年可再生能源利用量: {res['annual_renew_used']:.2f} kWh")
        print(f"  年光伏利用量: {res['annual_pv_used']:.2f} kWh")
        print(f"  年风电利用量: {res['annual_wind_used']:.2f} kWh")
        print(f"  年购电成本: {res['annual_grid_cost']:.2f} 元")
        print(f"  年弃光电量: {res['annual_pv_curtail']:.2f} kWh")
        print(f"  年弃风电量: {res['annual_wind_curtail']:.2f} kWh")
        print(f"  单位电量成本: {res['cost_per_kwh']:.4f} 元/kWh")

    # 保存结果到Excel
    output_data = []
    for area in ['A', 'B', 'C']:
        config = results[area]['config']
        res = results[area]['results']
        output_data.append({
            '园区': area,
            '光伏容量(kW)': config['pv_capacity'],
            '风电容量(kW)': config['wind_capacity'],
            '储能功率(kW)': config['ess_power'],
            '储能容量(kWh)': config['ess_capacity'],
            '投资成本(元)': res['investment_cost'],
            '年运行成本(元)': res['annual_operation_cost'],
            '总成本(元)': res['total_cost'],
            '年可再生能源利用量(kWh)': res['annual_renew_used'],
            '年光伏利用量(kWh)': res['annual_pv_used'],
            '年风电利用量(kWh)': res['annual_wind_used'],
            '年购电成本(元)': res['annual_grid_cost'],
            '年弃光电量(kWh)': res['annual_pv_curtail'],
            '年弃风电量(kWh)': res['annual_wind_curtail'],
            '单位电量成本(元/kWh)': res['cost_per_kwh']
        })

    output_df = pd.DataFrame(output_data)
    output_file = '独立运营风光储配置结果_全年分时电价.xlsx'
    output_df.to_excel(output_file, index=False)
    print(f"\n结果已保存到 '{output_file}'")