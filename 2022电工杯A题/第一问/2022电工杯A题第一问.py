import numpy as np
import matplotlib.pyplot as plt
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
        {'name': '机组2', 'P_max': 300, 'P_min': 90, 'a': 0.588, 'b': 65.12, 'c': 451.32, 'emission': 0.75},
        {'name': '机组3', 'P_max': 150, 'P_min': 45, 'a': 0.785, 'b': 139.6, 'c': 1049.50, 'emission': 0.79}
    ]


def load_demand_data():
    """从Excel文件中读取负荷数据"""
    # 读取Excel文件
    df = pd.read_excel('C:/Users/HP/Desktop/问题一数据.xlsx', sheet_name='Sheet1')
    # 提取负荷功率(p.u.)列并转换为实际功率（MW）
    load_demand = df['负荷功率(p.u.)'].values * 900
    return load_demand.tolist()


# ===================== 火电成本计算函数 =====================
def calculate_thermal_cost(units, P_results, carbon_price):
    """
    计算火电总成本（包括运行成本和碳捕集成本）
    :param units: 机组列表
    :param P_results: 各机组出力结果（列表的列表）
    :param carbon_price: 碳捕集单价（元/吨）
    :return: (总运行成本, 总碳捕集成本) 单位：元
    """
    total_fuel_cost = 0.0  # 煤耗成本（元）
    total_om_cost = 0.0  # 运行维护成本（元） om——Operation and Maintenance（运行和维护）
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


# ===================== 核心算法 =====================
def calculate_lambda(load, units):
    """
    计算初始λ值（不考虑约束）
    煤耗微增率方程：dF/dP = 2aP + b = λ
    因此 P_i = (λ - b_i) / (2a_i)
    约束：ΣP_i = load
    推导λ的表达式：
    load = Σ[(λ - b_i)/(2a_i)] = λ * Σ(1/(2a_i)) - Σ(b_i/(2a_i))
    => λ = (load + Σ(b_i/(2a_i))) / Σ(1/(2a_i))
    """
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
    """格式化数值为三位有效数字"""
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

    # 准备时间轴
    time_points = [f"{i // 4:02d}:{15 * (i % 4):02d}" for i in range(96)]

    # 初始化结果存储
    P_results = [[] for _ in range(len(units))]

    # 执行经济调度
    for i, load in enumerate(load_demand):
        P = economic_dispatch(load, units)
        for j in range(len(units)):
            P_results[j].append(P[j])

    # 可视化结果
    plt.figure(figsize=(12, 6))
    plt.plot(load_demand, 'k-', linewidth=2, label='系统日总负荷')

    colors = ['r-', 'g-', 'b-']
    for j, u in enumerate(units):
        plt.plot(P_results[j], colors[j],
                 label=f"{u['name']} ({u['P_min']}-{u['P_max']}MW)")

    plt.title('机组日发电计划曲线', fontproperties=font, fontsize=16)
    plt.ylabel('出力 (MW)', fontproperties=font)
    plt.xlabel('时间', fontproperties=font)
    plt.xticks(range(0, 96, 4), time_points[::4], rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')
    plt.tight_layout()
    plt.savefig('第一问相关曲线图.png', dpi=300)
    plt.show()

    # 计算总负荷电量（MWh）
    total_load_energy = sum(load_demand) * 0.25  # 每个时间点0.25小时

    # 碳捕集单价列表（元/吨）
    carbon_prices = [0, 60, 80, 100]

    # 输出结果标题
    print("\n结果如下\n")

    # 计算不同碳捕集单价下的成本
    for carbon_price in carbon_prices:
        # 计算火电成本
        operation_cost, carbon_cost = calculate_thermal_cost(units, P_results, carbon_price)

        # 总发电成本（元）= 火电运行成本 + 碳捕集成本
        total_generation_cost = operation_cost + carbon_cost

        # 单位供电成本（元/kWh）
        unit_supply_cost = total_generation_cost / (total_load_energy * 1000)

        # 转换为万元（用于输出）
        operation_cost_wan = operation_cost / 10000
        carbon_cost_wan = carbon_cost / 10000
        total_generation_cost_wan = total_generation_cost / 10000

        # 输出结果
        print(f"当碳捕集单价为 {carbon_price} 元/t 时：")
        print(f"  火电运行成本 = {format_value(operation_cost_wan)} 万元")
        print(f"  碳捕集成本 = {format_value(carbon_cost_wan)} 万元")
        print(f"  总发电成本 = {format_value(total_generation_cost_wan)} 万元")
        print(f"  单位供电成本 = {format_value(unit_supply_cost)} 元/kWh")
        print()

        # 保存机组出力数据为Excel文件
        # 创建DataFrame
        df = pd.DataFrame({
            '时间': time_points,
            '系统负荷(MW)': load_demand,
            '机组1出力(MW)': P_results[0],
            '机组2出力(MW)': P_results[1],
            '机组3出力(MW)': P_results[2],
            '机组1+机组2+机组3出力(MW)': [sum(x) for x in zip(P_results[0], P_results[1], P_results[2])]
        })

        # 保存到Excel
        excel_path = '机组出力数据.xlsx'
        df.to_excel(excel_path, index=False)
        print(f"机组出力数据已保存到: {excel_path}")


if __name__ == "__main__":
    main()