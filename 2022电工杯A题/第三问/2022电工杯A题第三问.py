import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import math
import pandas as pd
import os

# 设置中文字体
font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf", size=12)


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

    # 准备时间轴
    time_points = [f"{i // 4:02d}:{15 * (i % 4):02d}" for i in range(96)]

    # 初始化结果存储
    P_results = [[] for _ in range(len(units))]  # 火电机组结果
    heavy_loads = [0.0] * len(wind_power)  # 弃风存储，初始化为0
    light_loads = [0.0] * len(wind_power)  # 失负荷存储，初始话为0
    power_balance = []  # 功率平衡值存储

    # 执行经济调度
    for i, load_val in enumerate(load_demand):
        # 计算等效负荷（总负荷减去风电）
        equivalent_load = load_val - wind_power[i]

        # 计算弃风量
        # 当等效负荷小于剩余两机组最小出力时，出现弃风
        # 当等效负荷大于剩余两机组最大出力时，出现失负荷
        min_thermal_output = units[0]['P_min'] + units[1]['P_min']  # 机组1和机组3的最小出力总和
        max_thermal_output = units[0]['P_max'] + units[1]['P_max']  # 机组1和机组3的最大出力总和
        if equivalent_load < min_thermal_output:
            # 出现弃风，因为已达火电下限，此处记为重载
            heavy_load = min_thermal_output - equivalent_load
            equivalent_load = min_thermal_output
            heavy_loads[i] = heavy_load
        elif equivalent_load > max_thermal_output:
            # 出现失负荷，因为已达火电上限，此处记为轻载
            light_load = equivalent_load - max_thermal_output
            equivalent_load = max_thermal_output
            light_loads[i] = light_load

        # 对火电机组进行经济调度
        P = economic_dispatch(equivalent_load, units)

        # 保存结果
        for j in range(len(units)):
            P_results[j].append(P[j])

        # 计算总发电功率（火电+风电）
        total_generation = sum(P) + wind_power[i]

        # 计算功率平衡（总发电 - 负荷）
        balance = total_generation - load_val
        power_balance.append(balance)

        # 输出有弃风的时间点
        if heavy_loads[i] > 0:
            print(f"在 {time_points[i]} 时，弃风量为 {heavy_loads[i]:.2f} MW")
        # 输出有失负荷的时间点
        if light_loads[i] > 0:
            print(f"在 {time_points[i]} 时，失负荷 {light_loads[i]:.2f} MW")

    # 可视化结果
    plt.figure(figsize=(14, 10))

    # 子图1: 发电计划曲线
    plt.subplot(2, 1, 1)
    plt.plot(load_demand, 'k-', linewidth=2, label='系统日总负荷')

    colors = ['r-', 'g-', 'b-']
    for j, u in enumerate(units):
        plt.plot(P_results[j], colors[j],
                 label=f"{u['name']} ({u['P_min']}-{u['P_max']}MW)")

    # 添加风电曲线
    plt.plot(wind_power, 'c-', label='风电出力 (0-600MW)')

    # 添加火电+风电总出力曲线
    total_power = [sum(x) for x in zip(P_results[0], P_results[1], wind_power)]
    plt.plot(total_power, 'm--', label='火电+风电总出力')

    plt.title('1号和3号机组与600MW风电日发电计划曲线', fontproperties=font, fontsize=16)
    plt.ylabel('出力 (MW)', fontproperties=font)
    plt.xticks(range(0, 96, 4), time_points[::4], rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')

    # 子图2: 功率平衡曲线
    plt.subplot(2, 1, 2)
    plt.plot(power_balance, 'b-', label='系统功率平衡 (总发电-负荷)')

    # 标记弃风时间点
    g = 0 # 目的是让下面的图例只出现一次
    for i, c in enumerate(heavy_loads):
        if c > 0:
            g+=1
            if g >1:
                plt.plot(i, power_balance[i], 'ro', markersize=4)
            else:
                plt.plot(i, power_balance[i], 'ro', markersize=4, label='出现弃风的时间点')

    # 标记失负荷时间点
    g = 0 # 目的是让下面的图例只出现一次
    for j, u in enumerate(light_loads):
        if u > 0:
            g += 1
            if g > 1:
                plt.plot(j, power_balance[j], 'ks', markersize=4)
            else:
                plt.plot(j, power_balance[j], 'ks', markersize=4, label='出现失负荷的时间点')

    plt.title('系统功率平衡曲线', fontproperties=font, fontsize=16)
    plt.ylabel('功率平衡 (MW)', fontproperties=font)
    plt.xlabel('时间', fontproperties=font)
    plt.xticks(range(0, 96, 4), time_points[::4], rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')

    plt.tight_layout()
    plt.savefig('第三问相关曲线图.png', dpi=300)
    plt.show()

    # 计算总弃风量
    total_heavy_load = sum(heavy_loads)
    print(f"\n总弃风量: {total_heavy_load * 0.25:.2f} MWh")
    print(f"总弃风率: {total_heavy_load / sum(wind_power) * 100:.2f}%")

    # 计算总失负荷量
    total_light_load = sum(light_loads)
    print(f"\n总失负荷: {total_light_load * 0.25:.2f} MWh")
    print(f"总失负荷率: {total_light_load / sum(load_demand) * 100:.2f}%")

    # 保存机组出力数据为Excel文件
    # 创建DataFrame
    df = pd.DataFrame({
        '时间': time_points,
        '系统负荷(MW)': load_demand,
        '机组1出力(MW)': P_results[0],
        '机组3出力(MW)': P_results[1],
        '风电出力(MW)': wind_power,
        '弃风量(MW)': heavy_loads,
        '失负荷(MW)': light_loads,
        '火电+风电总出力(MW)': total_power,
        '功率平衡(MW)': power_balance
    })

    # 保存到Excel
    excel_path = '1号3号机组与600MW风电出力与功率平衡数据.xlsx'
    df.to_excel(excel_path, index=False)
    print(f"机组出力数据已保存到: {excel_path}")


if __name__ == "__main__":
    main()