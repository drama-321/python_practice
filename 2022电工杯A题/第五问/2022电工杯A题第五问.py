import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd

# 设置中文字体
font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf", size=12)


# ===================== 数据准备 =====================
def load_units_data():
    """加载机组参数数据（只保留1号机组）"""
    return [
        {'name': '机组1', 'P_max': 600, 'P_min': 180, 'a': 0.226, 'b': 30.42, 'c': 786.80, 'emission': 0.72}
    ]


def load_demand_data():
    """从Excel文件中读取负荷数据"""
    df = pd.read_excel('C:/Users/HP/Desktop/问题一数据.xlsx', sheet_name='Sheet1')
    load_demand = df['负荷功率(p.u.)'].values * 900
    return load_demand.tolist()


def load_wind_data():
    """从Excel文件中读取风电数据（900MW）"""
    df = pd.read_excel('C:/Users/HP/Desktop/问题五数据.xlsx', sheet_name='Sheet1')
    wind_power = df['风电功率(p.u.)'].values * 900  # 900MW风电
    return wind_power.tolist()


# ===================== 成本计算函数 =====================
def calculate_thermal_cost(units, thermal_power, carbon_price):
    total_fuel_cost = 0.0  # 煤耗成本（元）
    total_om_cost = 0.0  # 运行维护成本（元）
    total_carbon_cost = 0.0  # 碳捕集成本（元）

    # 煤价（元/kg）
    coal_price = 700 / 1000  # 700元/吨 = 0.7元/kg
    unit = units[0]  # 唯一机组

    # 遍历每个时间点（15分钟间隔）
    for P in thermal_power:
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


def calculate_light_load_cost(light_loads):
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


def calculate_energy_storage_cost(P_cap, E_cap, discharge_energy):
    """
    计算储能成本
    :param P_cap: 功率容量 (MW) 是指单位时间内充放电的最大功率，对应表格中单位功率成本
    :param E_cap: 能量容量 (MWh) 是指能够储存的最大容量，对应表格中单位能量成本
    :param discharge_energy: 总放电能量 (MWh)
    :return: (日均投资成本, 运维成本) 单位：元
    """
    # 储能成本参数（附表3）
    power_cost_per_kw = 3000  # 单位功率成本 (元/kW)
    energy_cost_per_kwh = 3000  # 单位能量成本 (元/kWh)
    om_cost_per_kwh = 0.05  # 单位电量运维成本 (元/kWh)
    lifetime_years = 10  # 运行年限

    # 计算总投资成本
    total_investment = (P_cap * 1000 * power_cost_per_kw) + (E_cap * 1000 * energy_cost_per_kwh)

    # 日均投资成本 = 总投资成本 / (运行年限 * 365)
    daily_investment_cost = total_investment / (lifetime_years * 365)

    # 运维成本 = 放电电量 × 单位运维成本 （运维是指在放电过程中）
    om_cost = discharge_energy * 1000 * om_cost_per_kwh

    return daily_investment_cost, om_cost


# ===================== 储能配置计算 =====================
def calculate_min_energy_storage(heavy_loads, light_loads, efficiency=0.9):
    """
    计算最小储能容量配置
    :param heavy_loads: 弃风量列表 (MW)
    :param light_loads: 失负荷量列表 (MW)
    :param efficiency: 充放电效率
    :return: (功率容量, 能量容量, 总充电量, 总放电量) 单位：MW, MWh, MWh, MWh
    """
    # 确定功率容量 (MW)
    max_heavy = max(heavy_loads) if heavy_loads else 0
    max_light = max(light_loads) if light_loads else 0
    P_cap = max(max_heavy, max_light)

    # 计算能量容量 (MWh)
    current_energy = 0  # 当前储能 (MWh)
    max_energy = 0  # 最大储能 (MWh)
    total_charge = 0  # 总充电量 (MWh)
    total_discharge = 0  # 总放电量 (MWh)

    # 模拟储能运行
    for i in range(len(heavy_loads)):
        # 充电阶段：利用弃风充电
        if heavy_loads[i] > 0:
            # 实际存储能量 = 充电功率 × 时间 × 效率
            charge_energy = heavy_loads[i] * 0.25 * efficiency
            # 更新储能
            current_energy += charge_energy
            total_charge += charge_energy
            # 更新最大储能
            if current_energy > max_energy:
                max_energy = current_energy

        # 放电阶段：弥补失负荷
        if light_loads[i] > 0:
            # 实际可放电量 = min(需求功率 × 时间, 当前储能)
            discharge_energy = min(light_loads[i] * 0.25, current_energy)
            # 更新储能
            current_energy -= discharge_energy
            total_discharge += discharge_energy

    # 能量容量 = 最大储能
    E_cap = max_energy

    return P_cap, E_cap, total_charge, total_discharge


# ===================== 主程序 =====================
def main():
    # 加载数据
    units = load_units_data()  # 只保留机组1
    load_demand = load_demand_data()
    wind_power = load_wind_data()  # 900MW风电数据

    # 获取机组参数
    unit = units[0]  # 唯一机组
    P_min = unit['P_min']
    P_max = unit['P_max']

    # 准备时间轴
    time_points = [f"{i // 4:02d}:{15 * (i % 4):02d}" for i in range(96)]

    # 初始化结果存储
    thermal_power = []  # 火电机组出力
    heavy_loads = [0.0] * len(wind_power)  # 弃风存储
    light_loads = [0.0] * len(wind_power)  # 失负荷存储
    power_balance = []  # 功率平衡值存储

    # 执行调度
    for i, load_val in enumerate(load_demand):
        # 计算等效负荷（总负荷减去风电）
        equivalent_load = load_val - wind_power[i]

        # 判断弃风和失负荷
        if equivalent_load < P_min:
            # 出现弃风，火电按最小出力运行
            heavy_loads[i] = P_min - equivalent_load
            equivalent_load = P_min
        elif equivalent_load > P_max:
            # 出现失负荷，火电按最大出力运行
            light_loads[i] = equivalent_load - P_max
            equivalent_load = P_max

        # 保存火电出力
        thermal_power.append(equivalent_load)

        # 计算总发电功率（火电+风电）
        total_generation = equivalent_load + wind_power[i]

        # 计算功率平衡（总发电 - 负荷）
        balance = total_generation - load_val
        power_balance.append(balance)

    # 计算总弃风量和总失负荷量
    total_heavy_load = sum(heavy_loads) * 0.25  # 转换为MWh
    total_light_load = sum(light_loads) * 0.25  # 转换为MWh

    # 计算最小储能配置
    P_cap, E_cap, total_charge, total_discharge = calculate_min_energy_storage(heavy_loads, light_loads)

    # ===================== 成本计算 =====================
    carbon_price = 60  # 单位碳捕集成本 (元/t)

    # 1. 火电成本计算
    operation_cost, carbon_cost = calculate_thermal_cost(units, thermal_power, carbon_price)

    # 2. 风电成本计算
    wind_om_cost, wind_heavy_load_cost = calculate_wind_cost(wind_power, heavy_loads)

    # 3. 失负荷成本计算
    light_load_cost = calculate_light_load_cost(light_loads)

    # 4. 储能成本计算
    daily_investment_cost, storage_om_cost = calculate_energy_storage_cost(
        P_cap, E_cap, total_discharge
    )

    # 5. 总发电成本
    total_generation_cost = (
            operation_cost + carbon_cost +
            wind_om_cost + wind_heavy_load_cost +
            light_load_cost +
            daily_investment_cost + storage_om_cost
    )

    # 6. 系统总负荷电量 (MWh)
    total_load_energy = sum(load_demand) * 0.25

    # 7. 单位供电成本 (元/kWh)
    unit_supply_cost = total_generation_cost / (total_load_energy * 1000)

    # ===================== 输出结果 =====================
    print("\n============== 第五问计算结果 ==============")
    print(f"总弃风量: {total_heavy_load:.2f} MWh")
    print(f"总失负荷量: {total_light_load:.2f} MWh")
    print(f"功率容量: {P_cap:.2f} MW, 能量容量: {E_cap:.2f} MWh")
    print(f"储能设备总充电量: {total_charge:.2f} MWh, 储能设备总放电量: {total_discharge:.2f} MWh")
    print("\n============== 成本计算结果 ==============")
    print(f"火电运行成本: {operation_cost / 10000:.2f} 万元")
    print(f"碳捕集成本: {carbon_cost / 10000:.2f} 万元")
    print(f"风电运维成本: {wind_om_cost / 10000:.2f} 万元")
    print(f"弃风损失: {wind_heavy_load_cost / 10000:.2f} 万元")
    print(f"失负荷损失: {light_load_cost / 10000:.2f} 万元")
    print(f"储能日均投资成本: {daily_investment_cost / 10000:.2f} 万元")
    print(f"储能运维成本: {storage_om_cost / 10000:.2f} 万元")
    print(f"总发电成本: {total_generation_cost / 10000:.2f} 万元")
    print(f"单位供电成本: {unit_supply_cost:.4f} 元/kWh")

    # ===================== 可视化结果 =====================
    plt.figure(figsize=(14, 10))

    # 子图1: 发电计划曲线
    plt.subplot(2, 1, 1)
    plt.plot(load_demand, 'k-', linewidth=2, label='系统日总负荷')
    plt.plot(thermal_power, 'r-', label=f"机组1 ({P_min}-{P_max}MW)")
    plt.plot(wind_power, 'c-', label='风电出力 (0-900MW)')

    # 添加火电+风电总出力曲线
    total_power = [thermal_power[i] + wind_power[i] for i in range(len(wind_power))]
    plt.plot(total_power, 'm--', label='火电+风电总出力')

    plt.title('1号机组与900MW风电日发电计划曲线', fontproperties=font, fontsize=16)
    plt.ylabel('出力 (MW)', fontproperties=font)
    plt.xticks(range(0, 96, 4), time_points[::4], rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')

    # 子图2: 功率平衡曲线
    plt.subplot(2, 1, 2)
    plt.plot(power_balance, 'b-', label='系统功率平衡 (总发电-负荷)')

    # 标记弃风时间点
    g = 0  # 目的是让下面的图例只出现一次
    for i, c in enumerate(heavy_loads):
        if c > 0:
            g += 1
            if g > 1:
                plt.plot(i, power_balance[i], 'ro', markersize=4)
            else:
                plt.plot(i, power_balance[i], 'ro', markersize=4, label='出现弃风的时间点')

    # 标记失负荷时间点
    g = 0  # 目的是让下面的图例只出现一次
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
    plt.savefig('第五问相关曲线图.png', dpi=300)
    plt.show()


if __name__ == "__main__":
    main()