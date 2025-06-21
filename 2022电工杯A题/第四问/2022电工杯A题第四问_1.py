from matplotlib.font_manager import FontProperties
import math
import pandas as pd

# 设置中文字体
font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf", size=12)


# ===================== 数据准备 =====================
def load_units_data():
    """加载机组参数数据（包含碳排放强度）"""
    return [
        {'name': '机组1', 'P_max': 600, 'P_min': 180, 'a': 0.226, 'b': 30.42, 'c': 786.80, 'emission': 0.72},
        {'name': '机组2', 'P_max': 300, 'P_min': 90, 'a': 0.588, 'b': 65.12, 'c': 451.32, 'emission': 0.75}
    ]


def load_demand_data():
    """从Excel文件中读取负荷数据"""
    df = pd.read_excel('C:/Users/HP/Desktop/问题一数据.xlsx', sheet_name='Sheet1')
    load_demand = df['负荷功率(p.u.)'].values * 900
    return load_demand.tolist()


def load_wind_data():
    """从Excel文件中读取风电数据"""
    df = pd.read_excel('C:/Users/HP/Desktop/问题二数据.xlsx', sheet_name='Sheet1')
    wind_power = df['风电功率(p.u.)'].values * 300
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
    """
    计算风电运维成本和弃风损失
    :param wind_power: 风电原始出力列表 (MW)
    :param heavy_loads: 弃风量列表 (MW)
    :return: (风电运维成本, 弃风损失) 单位：元
    """
    # 风电单位运维成本 (元/kWh)
    wind_om_cost_per_kwh = 0.045

    # 弃风损失单价 (元/kWh)
    wind_heavy_load_cost_per_kwh = 0.3
    total_wind_om_cost = 0.0
    total_heavy_load_cost = 0.0

    # 遍历每个时间点（15分钟间隔）
    for i in range(len(wind_power)):
        # 风电运维成本 = 实际发电量 × 单位运维成本
        wind_generation_kwh = wind_power[i] * 0.25 * 1000  # MW × 0.25h × 1000 = kWh
        wind_om_cost = wind_generation_kwh * wind_om_cost_per_kwh
        total_wind_om_cost += wind_om_cost

        # 弃风损失 = 弃风电量 × 弃风损失单价
        heavy_load_kwh = heavy_loads[i] * 0.25 * 1000  # MW × 0.25h × 1000 = kWh
        heavy_load_cost = heavy_load_kwh * wind_heavy_load_cost_per_kwh
        total_heavy_load_cost += heavy_load_cost

    return total_wind_om_cost, total_heavy_load_cost


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


# ===================== 格式化数值为三位有效数字 =====================
def format_value(value):
    if value == 0:
        return "0.00"

    magnitude = math.floor(math.log10(abs(value)))

    if magnitude >= 3 or magnitude <= -3:
        return f"{value:.3e}"

    decimals = max(0, 2 - magnitude)
    return f"{value:.{int(decimals)}f}"


# ===================== 主程序 =====================
def main():
    # 加载数据
    units = load_units_data()
    load_demand = load_demand_data()
    wind_power = load_wind_data()

    # 准备时间轴
    time_points = [f"{i // 4:02d}:{15 * (i % 4):02d}" for i in range(96)]

    # 初始化结果存储
    P_results = [[] for _ in range(len(units))]
    heavy_loads = []  # 弃风存储
    power_balance = []  # 功率平衡值存储

    # 执行经济调度
    for i, load_val in enumerate(load_demand):
        # 计算等效负荷（总负荷减去风电）
        equivalent_load = load_val - wind_power[i]

        # 计算弃风量
        min_thermal_output = units[0]['P_min'] + units[1]['P_min']
        if equivalent_load < min_thermal_output:
            heavy_load = min_thermal_output - equivalent_load
        else:
            heavy_load = 0

        heavy_loads.append(heavy_load)

        # 对火电机组进行经济调度
        P = economic_dispatch(equivalent_load, units)

        # 保存结果
        for j in range(len(units)):
            P_results[j].append(P[j])

        # 计算总发电功率（火电+实际风电）
        total_generation = sum(P) + wind_power[i]
        balance = total_generation - load_val
        power_balance.append(balance)

    # 计算总弃风量
    total_heavy_load = sum(heavy_loads)
    print(f"\n总弃风量: {total_heavy_load * 0.25:.2f} MWh")
    print(f"总弃风率: {total_heavy_load / sum(wind_power) * 100:.2f}%")

    # ===================== 新增成本计算部分 =====================
    # 碳捕集单价列表（元/吨）
    carbon_prices = [0, 60, 80, 100]

    # 计算总负荷电量（MWh）
    total_load_energy = sum(load_demand) * 0.25

    # 计算总弃风电量（MWh）
    total_heavy_load_energy = total_heavy_load * 0.25

    # 输出结果标题
    print("\n成本计算结果如下\n")

    # 准备结果表格
    results_table = []

    # 计算不同碳捕集单价下的成本
    for carbon_price in carbon_prices:
        # 计算火电成本
        operation_cost, carbon_cost = calculate_thermal_cost(units, P_results, carbon_price)

        # 计算风电成本和弃风损失
        wind_om_cost, wind_heavy_load_cost = calculate_wind_cost(wind_power, heavy_loads)

        # 总发电成本（元）= 火电成本 + 风电运维成本 + 弃风损失
        total_generation_cost = operation_cost + carbon_cost + wind_om_cost + wind_heavy_load_cost

        # 单位供电成本（元/kWh）
        unit_supply_cost = total_generation_cost / (total_load_energy * 1000)

        # 转换为万元（用于输出）
        operation_cost_wan = operation_cost / 10000
        carbon_cost_wan = carbon_cost / 10000
        wind_om_cost_wan = wind_om_cost / 10000
        wind_heavy_load_cost_wan = wind_heavy_load_cost / 10000
        total_generation_cost_wan = total_generation_cost / 10000

        # 保存结果用于表格输出
        results_table.append({
            'carbon_price': carbon_price,
            'operation_cost': operation_cost_wan,
            'carbon_cost': carbon_cost_wan,
            'wind_om_cost': wind_om_cost_wan,
            'heavy_load_energy': total_heavy_load_energy,
            'wind_heavy_load_cost': wind_heavy_load_cost_wan,
            'total_cost': total_generation_cost_wan,
            'unit_cost': unit_supply_cost
        })

        # 输出结果
        print(f"当碳捕集单价为 {carbon_price} 元/t 时：")
        print(f"  火电运行成本 = {format_value(operation_cost_wan)} 万元")
        print(f"  碳捕集成本 = {format_value(carbon_cost_wan)} 万元")
        print(f"  风电运维成本 = {format_value(wind_om_cost_wan)} 万元")
        print(f"  弃风损失 = {format_value(wind_heavy_load_cost_wan)} 万元")
        print(f"  总发电成本 = {format_value(total_generation_cost_wan)} 万元")
        print(f"  单位供电成本 = {format_value(unit_supply_cost)} 元/kWh")
        print(f"  总弃风电量 = {total_heavy_load_energy:.2f} MWh")
        print()

    # 保存结果到Excel
    df = pd.DataFrame(results_table)
    df.columns = ['碳捕集成本(元/t)', '火电运行成本(万元)', '碳捕集成本(万元)',
                  '风电运维成本(万元)', '弃风电量(MWh)', '弃风损失(万元)',
                  '总发电成本(万元)','单位供电成本(元/kWh)']
    excel_path = '第二问成本计算结果.xlsx'
    df.to_excel(excel_path, index=False)
    print(f"成本计算结果已保存到: {excel_path}")


if __name__ == "__main__":
    main()