import numpy as np
import pandas as pd
import math


# ===================== 数据准备 =====================
def load_units_data():
    """加载机组参数数据（只包含1号和3号机组）"""
    return [
        {'name': '机组1', 'P_max': 600, 'P_min': 180, 'a': 0.226, 'b': 30.42, 'c': 786.80, 'emission': 0.72},
        {'name': '机组3', 'P_max': 150, 'P_min': 45, 'a': 0.785, 'b': 139.6, 'c': 1049.50, 'emission': 0.79}
    ]


def load_demand_data():
    """从Excel文件中读取负荷数据"""
    # 读取Excel文件
    df = pd.read_excel('C:/Users/HP/Desktop/问题一数据.xlsx', sheet_name='Sheet1')
    # 提取负荷功率(p.u.)列并转换为实际功率（MW）
    load_demand = df['负荷功率(p.u.)'].values * 900
    return load_demand.tolist()


def load_wind_data():
    """从Excel文件中读取风电数据（600MW）"""
    # 读取Excel文件
    df = pd.read_excel('C:/Users/HP/Desktop/问题三数据.xlsx', sheet_name='Sheet1')
    # 提取风电功率(p.u.)列并转换为实际功率（MW）
    wind_power = df['风电功率(p.u.)'].values * 600
    return wind_power.tolist()


# ===================== 火电成本计算函数 =====================
def calculate_thermal_cost(units, P_results, carbon_price):
    total_fuel_cost = 0.0  # 煤耗成本（元）
    total_om_cost = 0.0  # 运行维护成本（元）
    total_carbon_cost = 0.0  # 碳捕集成本（元）

    # 煤价（元/kg）
    coal_price = 700 / 1000  # 700元/吨 = 0.7元/kg

    # 遍历每个时间点（15分钟间隔）
    for t in range(len(P_results[0])):
        # 遍历每台机组
        for i, unit in enumerate(units):
            P = P_results[i][t]  # 机组在t时刻的出力（MW）

            # 计算煤耗量（kg/h）
            F_hourly = unit['a'] * P ** 2 + unit['b'] * P + unit['c']

            # 15分钟的煤耗量（kg）
            F_15min = F_hourly * 0.25

            # 煤耗成本（元）
            fuel_cost = F_15min * coal_price
            total_fuel_cost += fuel_cost

            # 运行维护成本（元）= 0.5 * 煤耗成本
            om_cost = 0.5 * fuel_cost
            total_om_cost += om_cost

            # 计算发电量（MWh）
            generation_mwh = P * 0.25  # MW × 0.25h = MWh

            # 计算碳排放量（kg）
            carbon_emission = generation_mwh * 1000 * unit['emission']  # MWh × 1000 kWh/MWh × kg/kWh

            # 碳捕集成本（元）
            carbon_cost = carbon_emission * (carbon_price / 1000)  # 碳捕集单价元/吨 = 元/1000kg
            total_carbon_cost += carbon_cost

    # 总运行成本 = 煤耗成本 + 运行维护成本
    total_operation_cost = total_fuel_cost + total_om_cost

    return total_operation_cost, total_carbon_cost


# ===================== 风电成本计算函数 =====================
def calculate_wind_cost(wind_power, heavy_loads):
    # 风电单位运维成本 (元/kWh)
    wind_om_cost_per_kwh = 0.045

    # 弃风损失单价 (元/kWh)
    wind_curtailment_cost_per_kwh = 0.3
    total_wind_om_cost = 0.0
    total_curtailment_cost = 0.0

    # 遍历每个时间点（15分钟间隔）
    for i in range(len(wind_power)):
        # 风电运维成本 = 实际发电量 × 单位运维成本
        wind_generation_kwh = wind_power[i] * 0.25 * 1000  # MW × 0.25h × 1000 = kWh
        wind_om_cost = wind_generation_kwh * wind_om_cost_per_kwh
        total_wind_om_cost += wind_om_cost

        # 弃风损失 = 弃风电量 × 弃风损失单价
        curtailment_kwh = heavy_loads[i] * 0.25 * 1000  # MW × 0.25h × 1000 = kWh
        curtailment_cost = curtailment_kwh * wind_curtailment_cost_per_kwh
        total_curtailment_cost += curtailment_cost

    return total_wind_om_cost, total_curtailment_cost


# ===================== 失负荷成本计算函数 =====================
def calculate_light_load_cost(light_loads):
    """
    计算失负荷损失成本
    :param light_loads: 失负荷量列表 (MW)
    :return: 失负荷损失成本 (元)
    """
    # 失负荷损失单价 (元/kWh)
    light_load_cost_per_kwh = 8.0
    total_light_load_cost = 0.0

    # 遍历每个时间点（15分钟间隔）
    for light_load in light_loads:
        # 失负荷损失 = 失负荷电量 × 失负荷损失单价
        light_load_kwh = light_load * 0.25 * 1000  # MW × 0.25h × 1000 = kWh
        light_load_cost = light_load_kwh * light_load_cost_per_kwh
        total_light_load_cost += light_load_cost

    return total_light_load_cost


# ===================== 核心算法 =====================
def calculate_lambda(load, units):
    sum_inv_2a = sum(1 / (2 * u['a']) for u in units)
    sum_b_over_2a = sum(u['b'] / (2 * u['a']) for u in units)
    return (load + sum_b_over_2a) / sum_inv_2a


def update_generation(lambda_val, units):
    return [(lambda_val - u['b']) / (2 * u['a']) for u in units]


def economic_dispatch(load, units, max_iter=20, tol=0.01):
    n = len(units)
    fixed = [False] * n
    P = [0.0] * n

    lambda_val = calculate_lambda(load, units)
    P = update_generation(lambda_val, units)

    for _ in range(max_iter):
        new_fixed = fixed.copy()
        violations = False

        for i, u in enumerate(units):
            if fixed[i]:
                continue

            if P[i] < u['P_min']:
                P[i] = u['P_min']
                new_fixed[i] = True
                violations = True
            elif P[i] > u['P_max']:
                P[i] = u['P_max']
                new_fixed[i] = True
                violations = True

        fixed = new_fixed

        if not violations:
            if abs(sum(P) - load) < tol:
                break

        adjustable_units = [u for i, u in enumerate(units) if not fixed[i]]
        adjustable_indices = [i for i in range(n) if not fixed[i]]

        if not adjustable_units:
            break

        fixed_power = sum(P[i] for i in range(n) if fixed[i])
        remaining_load = max(load - fixed_power, 0)

        lambda_val = calculate_lambda(remaining_load, adjustable_units)
        P_adjustable = update_generation(lambda_val, adjustable_units)

        for idx, i in enumerate(adjustable_indices):
            P[i] = P_adjustable[idx]

    return P


# ===================== 主程序 =====================
def main():
    # 加载数据
    units = load_units_data()
    load_demand = load_demand_data()
    wind_power = load_wind_data()  # 加载风电数据（600MW）

    # 初始化结果存储
    P_results = [[] for _ in range(len(units))]  # 火电机组结果
    heavy_loads = [0.0] * len(wind_power)  # 弃风存储，初始化为0
    light_loads = [0.0] * len(wind_power)  # 失负荷存储，初始化为0

    # 执行经济调度
    for i, load_val in enumerate(load_demand):
        # 计算等效负荷（总负荷减去风电）
        equivalent_load = load_val - wind_power[i]

        # 计算弃风量和失负荷量
        min_thermal_output = units[0]['P_min'] + units[1]['P_min']  # 机组1和机组3的最小出力总和
        max_thermal_output = units[0]['P_max'] + units[1]['P_max']  # 机组1和机组3的最大出力总和

        if equivalent_load < min_thermal_output:
            # 出现弃风，因为已达火电下限
            heavy_loads[i] = min_thermal_output - equivalent_load
            equivalent_load = min_thermal_output
        elif equivalent_load > max_thermal_output:
            # 出现失负荷，因为已达火电上限
            light_loads[i] = equivalent_load - max_thermal_output
            equivalent_load = max_thermal_output

        # 对火电机组进行经济调度
        P = economic_dispatch(equivalent_load, units)

        # 保存结果
        for j in range(len(units)):
            P_results[j].append(P[j])

    # 计算总弃风量和总失负荷量
    total_heavy_load = sum(heavy_loads) * 0.25  # 转换为MWh
    total_light_load = sum(light_loads) * 0.25  # 转换为MWh

    # ===================== 成本计算部分 =====================
    # 碳捕集单价列表（元/吨）
    carbon_prices = [0, 60, 80, 100]

    # 计算总负荷电量（MWh）
    total_load_energy = sum(load_demand) * 0.25

    # 准备结果表格
    results_table = []

    # 计算不同碳捕集单价下的成本
    for carbon_price in carbon_prices:
        # 计算火电成本
        operation_cost, carbon_cost = calculate_thermal_cost(units, P_results, carbon_price)

        # 计算风电成本和弃风损失
        wind_om_cost, wind_heavy_load_cost = calculate_wind_cost(wind_power, heavy_loads)

        # 计算失负荷损失
        light_load_cost = calculate_light_load_cost(light_loads)

        # 总发电成本（元）= 火电成本 + 风电成本 + 弃风损失 + 失负荷损失
        total_generation_cost = operation_cost + carbon_cost + wind_om_cost + wind_heavy_load_cost + light_load_cost

        # 单位供电成本（元/kWh）
        unit_supply_cost = total_generation_cost / (total_load_energy * 1000)

        # 转换为万元（用于输出）
        operation_cost_wan = operation_cost / 10000
        carbon_cost_wan = carbon_cost / 10000
        wind_om_cost_wan = wind_om_cost / 10000
        wind_heavy_load_cost_wan = wind_heavy_load_cost / 10000
        light_load_cost_wan = light_load_cost / 10000
        total_generation_cost_wan = total_generation_cost / 10000

        # 保存结果
        results_table.append({
            'carbon_price': carbon_price,
            'operation_cost': operation_cost_wan,
            'carbon_cost': carbon_cost_wan,
            'wind_om_cost': wind_om_cost_wan,
            'heavy_load_energy': total_heavy_load,
            'wind_heavy_load_cost': wind_heavy_load_cost_wan,
            'light_load_energy': total_light_load,
            'light_load_cost': light_load_cost_wan,
            'total_cost': total_generation_cost_wan,
            'unit_cost': unit_supply_cost
        })

    # 创建DataFrame并保存到Excel
    df = pd.DataFrame(results_table)
    df.columns = ['碳捕集成本(元/t)', '火电运行成本(万元)', '碳捕集成本(万元)',
                  '风电运维成本(万元)', '弃风电量(MWh)', '弃风损失(万元)',
                  '失负荷电量(MWh)', '失负荷损失(万元)',
                  '总发电成本(万元)', '单位供电成本(元/kWh)']

    excel_path = '第三问成本计算结果.xlsx'
    df.to_excel(excel_path, index=False)
    print(f"成本计算结果已保存到: {excel_path}")


    # 输出汇总结果
    print("\n风电装机600MW替代机组2场景下系统相关指标统计：")
    print(df)


if __name__ == "__main__":
    main()