import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import math
import pandas as pd
import random

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


# ===================== PSO算法实现 =====================
def pso_economic_dispatch(load, units, max_iter=100, num_particles=50, w=0.5, c1=1.5, c2=1.5):
    """
    使用PSO算法求解经济调度问题
    :param load: 当前时刻的负荷需求(MW)
    :param units: 机组列表
    :param max_iter: 最大迭代次数
    :param num_particles: 粒子数量
    :param w: 惯性权重
    :param c1: 个体学习因子
    :param c2: 群体学习因子
    :return: 各机组最优出力(MW)
    """
    num_units = len(units)

    # 获取机组出力上下限
    bounds = np.array([[u['P_min'], u['P_max']] for u in units])

    # 初始化粒子群
    particles_p = np.zeros((num_particles, num_units))
    particles_v = np.zeros((num_particles, num_units))

    # 初始化位置和速度
    for i in range(num_particles):
        for j in range(num_units):
            particles_p[i, j] = random.uniform(bounds[j, 0], bounds[j, 1])
            particles_v[i, j] = random.uniform(-1, 1) * (bounds[j, 1] - bounds[j, 0]) / 10.0

    # 初始化个体最优位置和适应度
    pbest_p = particles_p.copy()
    pbest_fitness = np.full(num_particles, float('inf'))

    # 初始化全局最优
    gbest_p = np.zeros(num_units)
    gbest_fitness = float('inf')

    # 煤价（元/kg）
    coal_price = 700 / 1000  # 700元/吨 = 0.7元/kg

    # 迭代优化
    for _ in range(max_iter):
        for i in range(num_particles):
            # 计算当前粒子的总出力
            total_power = np.sum(particles_p[i])

            # 计算惩罚项（功率不平衡惩罚）
            imbalance_penalty = 10000 * abs(total_power - load) ** 2

            # 计算约束惩罚（出力越限惩罚）
            constraint_penalty = 0
            for j in range(num_units):
                if particles_p[i, j] < bounds[j, 0]:
                    constraint_penalty += 10000 * (bounds[j, 0] - particles_p[i, j]) ** 2
                elif particles_p[i, j] > bounds[j, 1]:
                    constraint_penalty += 10000 * (particles_p[i, j] - bounds[j, 1]) ** 2

            # 计算总煤耗成本（目标函数）
            total_cost = 0
            for j, unit in enumerate(units):
                # 计算煤耗量（kg/h）
                F_hourly = unit['a'] * particles_p[i, j] ** 2 + unit['b'] * particles_p[i, j] + unit['c']
                # 15分钟的煤耗成本（元）
                fuel_cost = F_hourly * 0.25 * coal_price
                # 总运行成本 = 1.5 * 煤耗成本（包括运行维护成本）
                total_cost += 1.5 * fuel_cost

            # 总适应度 = 总成本 + 惩罚项
            fitness = total_cost + imbalance_penalty + constraint_penalty

            # 更新个体最优
            if fitness < pbest_fitness[i]:
                pbest_fitness[i] = fitness
                pbest_p[i] = particles_p[i].copy()

            # 更新全局最优
            if fitness < gbest_fitness:
                gbest_fitness = fitness
                gbest_p = particles_p[i].copy()

        # 更新粒子速度和位置
        for i in range(num_particles):
            for j in range(num_units):
                # 更新速度
                r1 = random.random()
                r2 = random.random()
                cognitive = c1 * r1 * (pbest_p[i, j] - particles_p[i, j])
                social = c2 * r2 * (gbest_p[j] - particles_p[i, j])
                particles_v[i, j] = w * particles_v[i, j] + cognitive + social

                # 更新位置
                particles_p[i, j] += particles_v[i, j]

                # 边界处理
                particles_p[i, j] = max(bounds[j, 0], min(particles_p[i, j], bounds[j, 1]))

    return gbest_p.tolist()


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

    # 使用PSO算法执行经济调度
    for i, load in enumerate(load_demand):
        P = pso_economic_dispatch(load, units)
        for j in range(len(units)):
            P_results[j].append(P[j])

    print("所有时段调度完成!")

    # 可视化结果
    plt.figure(figsize=(12, 6))
    plt.plot(load_demand, 'k-', linewidth=2, label='系统日总负荷')

    colors = ['r-', 'g-', 'b-']
    for j, u in enumerate(units):
        plt.plot(P_results[j], colors[j],
                 label=f"{u['name']} ({u['P_min']}-{u['P_max']}MW)")

    plt.title('机组日发电计划曲线 (PSO优化)', fontproperties=font, fontsize=16)
    plt.ylabel('出力 (MW)', fontproperties=font)
    plt.xlabel('时间', fontproperties=font)
    plt.xticks(range(0, 96, 4), time_points[::4], rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')
    plt.tight_layout()
    plt.savefig('第一问相关曲线图_PSO.png', dpi=300)
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
    excel_path = '机组出力数据_PSO.xlsx'
    df.to_excel(excel_path, index=False)
    print(f"机组出力数据已保存到: {excel_path}")


if __name__ == "__main__":
    main()