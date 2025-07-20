import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 定义所有园区的装机容量
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

# 固定储能参数
storage_params = {
    'soc_min': 10,  # SOC下限 (%)
    'soc_max': 90,  # SOC上限 (%)
    'efficiency': 0.95,  # 充放电效率
    'power_cost': 800,  # 功率单价 (元/kW)
    'energy_cost': 1800,  # 能量单价 (元/kWh)
    'lifetime': 10  # 运行寿命 (年)
}

# 读取数据
load_data = pd.read_excel('C:/Users/HP/Desktop/附件1：各园区典型日负荷数据.xlsx')
renewable_data = pd.read_excel('C:/Users/HP/Desktop/附件2：各园区典型日风光发电数据.xlsx', skiprows=2, header=None,
                               names=['时间', 'A_pv', 'B_wind', 'C_pv', 'C_wind'])

# 数据处理
load_data['时间（h）'] = load_data['时间（h）'].astype(str)
renewable_data['时间'] = renewable_data['时间'].astype(str)
date = pd.merge(load_data, renewable_data, left_on='时间（h）', right_on='时间', how='inner')
date['A_wind'] = 0.0
date['B_pv'] = 0.0

# 计算实际发电功率
for area in ['A', 'B', 'C']:
    date[f'{area}_pv_power'] = capacities[area]['pv'] * date[f'{area}_pv']
    date[f'{area}_wind_power'] = capacities[area]['wind'] * date[f'{area}_wind']
    date[f'{area}_total_power'] = date[f'{area}_pv_power'] + date[f'{area}_wind_power']


# ======================= PSO优化部分 =======================
def simulate_storage(area, power, capacity):
    """模拟给定储能配置下的运行情况"""
    # 初始化储能状态
    storage_soc = 90.0  # 初始SOC 90%

    # 初始化结果字典
    simulation = {
        'renew_used': np.zeros(len(date)),
        'renew_to_storage': np.zeros(len(date)),
        'storage_discharge': np.zeros(len(date)),
        'curtail_pv': np.zeros(len(date)),
        'curtail_wind': np.zeros(len(date)),
        'grid_purchase': np.zeros(len(date)),
        'soc': np.zeros(len(date)),
        'pv_used': np.zeros(len(date)),  # 记录光伏利用量
        'wind_used': np.zeros(len(date))  # 记录风电利用量
    }

    # 无效配置检查
    if power == 0 and capacity > 0:
        return float('inf'), simulation

    # 计算储能投资成本
    storage_investment = power * storage_params['power_cost'] + capacity * storage_params['energy_cost']
    storage_daily_cost = storage_investment / (storage_params['lifetime'] * 365)

    # 模拟24小时运行
    for i, row in date.iterrows():
        load = row[f'园区{area}负荷(kW)']
        pv_gen = row[f'{area}_pv_power']
        wind_gen = row[f'{area}_wind_power']

        # 当前储能状态 (kWh)
        soc_kwh = storage_soc / 100 * capacity if capacity > 0 else 0

        # 1. 优先使用光伏发电
        pv_used = min(pv_gen, load)
        load_after_pv = load - pv_used

        # 2. 然后使用风电
        wind_used = min(wind_gen, load_after_pv)
        load_remain = load_after_pv - wind_used

        renew_used = pv_used + wind_used
        simulation['renew_used'][i] = renew_used
        simulation['pv_used'][i] = pv_used
        simulation['wind_used'][i] = wind_used

        # 3. 计算剩余可再生能源
        pv_surplus = pv_gen - pv_used
        wind_surplus = wind_gen - wind_used
        total_surplus = pv_surplus + wind_surplus

        # 4. 剩余可再生能源处理
        if total_surplus > 0:
            # 计算最大充电功率
            max_charge_kw = min(
                power,
                (storage_params['soc_max'] / 100 * capacity - soc_kwh) / storage_params['efficiency']
            ) if capacity > 0 else 0

            charge_kw = min(total_surplus, max_charge_kw)
            charge_actual = charge_kw * storage_params['efficiency']

            # 更新SOC
            soc_kwh += charge_actual
            simulation['renew_to_storage'][i] = charge_kw

            # 弃电分配（优先弃风电）
            curtail_total = total_surplus - charge_kw
            wind_curtail = min(wind_surplus, curtail_total)
            pv_curtail = curtail_total - wind_curtail
            simulation['curtail_pv'][i] = pv_curtail
            simulation['curtail_wind'][i] = wind_curtail
        else:
            simulation['renew_to_storage'][i] = 0
            simulation['curtail_pv'][i] = pv_surplus
            simulation['curtail_wind'][i] = wind_surplus

        # 5. 储能补充
        if load_remain > 0 and capacity > 0:
            # 计算最大放电功率
            max_discharge_kw = min(
                power,
                (soc_kwh - storage_params['soc_min'] / 100 * capacity) * storage_params['efficiency']
            )
            discharge_kw = min(load_remain, max_discharge_kw)
            discharge_actual = discharge_kw / storage_params['efficiency']

            # 更新SOC
            soc_kwh -= discharge_actual
            load_remain -= discharge_kw
            simulation['storage_discharge'][i] = discharge_kw
        else:
            simulation['storage_discharge'][i] = 0

        # 6. 电网补充
        simulation['grid_purchase'][i] = max(0, load_remain)

        # 7. 记录SOC
        storage_soc = soc_kwh / capacity * 100 if capacity > 0 else 0
        simulation['soc'][i] = storage_soc

    # 计算成本
    pv_used_total = simulation['pv_used'].sum()
    wind_used_total = simulation['wind_used'].sum()

    renew_cost = pv_used_total * electricity_prices['pv'] + wind_used_total * electricity_prices['wind']
    grid_cost = simulation['grid_purchase'].sum() * electricity_prices['grid']
    total_cost = renew_cost + grid_cost + storage_daily_cost

    return total_cost, simulation


def pso_optimize_storage(area, num_particles=20, max_iter=50, w=0.8, c1=1.5, c2=1.5):
    """使用PSO算法优化储能配置"""
    # 定义搜索边界
    power_bounds = [0, 200]  # 功率范围 (kW)
    capacity_bounds = [0, 400]  # 容量范围 (kWh)

    # 初始化粒子群
    particles_p = np.zeros((num_particles, 2))  # 位置
    particles_v = np.zeros((num_particles, 2))  # 速度

    # 随机初始化位置和速度
    particles_p[:, 0] = np.random.uniform(power_bounds[0], power_bounds[1], num_particles)
    particles_p[:, 1] = np.random.uniform(capacity_bounds[0], capacity_bounds[1], num_particles)
    particles_v = np.random.uniform(-1, 1, (num_particles, 2))

    # 初始化个体最优
    pbest_p = particles_p.copy()  # 个体最优位置
    pbest_cost = np.full(num_particles, float('inf'))  # 个体最优成本

    # 初始化全局最优
    gbest_p = np.zeros(2)  # 全局最优位置
    gbest_cost = float('inf')  # 全局最优成本
    gbest_simulation = None  # 全局最优模拟结果

    # 评估初始粒子群
    for i in range(num_particles):
        power = particles_p[i, 0]
        capacity = particles_p[i, 1]
        cost, _ = simulate_storage(area, power, capacity)

        pbest_cost[i] = cost
        if cost < gbest_cost:
            gbest_cost = cost
            gbest_p = particles_p[i].copy()

    # PSO主循环
    convergence = []
    for _ in range(max_iter):
        for i in range(num_particles):
            # 更新速度和位置
            r1, r2 = np.random.rand(2)
            particles_v[i] = (w * particles_v[i] +
                                c1 * r1 * (pbest_p[i] - particles_p[i]) +
                                c2 * r2 * (gbest_p - particles_p[i]))

            particles_p[i] += particles_v[i]

            # 应用边界约束
            particles_p[i, 0] = np.clip(particles_p[i, 0], power_bounds[0], power_bounds[1])
            particles_p[i, 1] = np.clip(particles_p[i, 1], capacity_bounds[0], capacity_bounds[1])

            # 评估新位置
            power = particles_p[i, 0]
            capacity = particles_p[i, 1]
            cost, simulation = simulate_storage(area, power, capacity)

            # 更新个体最优
            if cost < pbest_cost[i]:
                pbest_cost[i] = cost
                pbest_p[i] = particles_p[i].copy()

                # 更新全局最优
                if cost < gbest_cost:
                    gbest_cost = cost
                    gbest_p = particles_p[i].copy()
                    gbest_simulation = simulation

        convergence.append(gbest_cost)

    # 提取最优配置结果
    power_opt, capacity_opt = gbest_p
    power_opt = max(0, round(power_opt))  # 取整处理
    capacity_opt = max(0, round(capacity_opt))

    # 重新计算最优配置的详细结果
    total_load = date[f'园区{area}负荷(kW)'].sum()
    pv_used_total = gbest_simulation['pv_used'].sum()
    wind_used_total = gbest_simulation['wind_used'].sum()
    curtail_total = gbest_simulation['curtail_pv'].sum() + gbest_simulation['curtail_wind'].sum()
    grid_purchase_total = gbest_simulation['grid_purchase'].sum()

    # 计算成本
    renew_cost = pv_used_total * electricity_prices['pv'] + wind_used_total * electricity_prices['wind']
    grid_cost = grid_purchase_total * electricity_prices['grid']

    # 储能投资成本
    storage_investment = power_opt * storage_params['power_cost'] + capacity_opt * storage_params['energy_cost']
    storage_daily_cost = storage_investment / (storage_params['lifetime'] * 365)

    # 整理结果
    results = {
        '最优功率(kW)': power_opt,
        '最优容量(kWh)': capacity_opt,
        '总负荷电量(kWh)': total_load,
        '弃电量(kWh)': curtail_total,
        '网购电量(kWh)': grid_purchase_total,
        '储能充入电量(kWh)': gbest_simulation['renew_to_storage'].sum(),
        '储能供电量(kWh)': gbest_simulation['storage_discharge'].sum(),
        '光伏利用量(kWh)': pv_used_total,
        '风电利用量(kWh)': wind_used_total,
        '总供电成本(元)': gbest_cost,
        '可再生能源成本(元)': renew_cost,
        '网购电成本(元)': grid_cost,
        '储能投资成本(元)': storage_investment,
        '储能日分摊成本(元)': storage_daily_cost
    }

    return (power_opt, capacity_opt), results, gbest_simulation


# ======================= 主程序 =======================
# 优化每个园区的储能配置并存储结果
optimal_configs = {}
simulation_data = {}
for area in ['A', 'B', 'C']:
    best_config, best_results, sim_data = pso_optimize_storage(area)
    optimal_configs[area] = best_results
    simulation_data[area] = sim_data
    print(
        f"园区{area}优化完成: {best_config[0]}kW/{best_config[1]}kWh, 成本: {best_results['总供电成本(元)']:.2f}元/天")

# 打印优化结果
for area, config in optimal_configs.items():
    print(f"\n园区{area}最优配置: {config['最优功率(kW)']}kW/{config['最优容量(kWh)']}kWh")
    print(f"总供电成本: {config['总供电成本(元)']:.2f}元/天")
    print(f"可再生能源成本: {config['可再生能源成本(元)']:.2f}元/天")
    print(f"网购电成本: {config['网购电成本(元)']:.2f}元/天")
    print(f"光伏利用量: {config['光伏利用量(kWh)']:.2f}kWh")
    print(f"风电利用量: {config['风电利用量(kWh)']:.2f}kWh")
    print(f"网购电量: {config['网购电量(kWh)']:.2f}kWh")
    print(f"弃电量: {config['弃电量(kWh)']:.2f}kWh")

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

    # 获取最优配置
    opt_power = optimal_configs[area]['最优功率(kW)']
    opt_capacity = optimal_configs[area]['最优容量(kWh)']

    # 获取模拟数据
    sim = simulation_data[area]

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
        ax1.plot(hours, sim['curtail_pv'], 'r--', label='弃光', linewidth=1.8, alpha=0.7)

    # 绘制弃风曲线
    if wind_cap > 0:
        ax1.plot(hours, sim['curtail_wind'], 'm--', label='弃风', linewidth=1.8, alpha=0.7)

    # 设置图表属性
    ax1.set_title(f'园区{area} - 最优储能配置: {opt_power}kW/{opt_capacity}kWh', fontsize=18)
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
    ax2.plot(hours, sim['soc'], 'b-', label='储能SOC', linewidth=2.5)
    ax2.set_ylabel('SOC (%)', fontsize=14)
    ax2.set_ylim(0, 100)

    # 绘制储能充放电曲线（双Y轴）
    ax3 = ax2.twinx()
    ax3.bar(hours, sim['renew_to_storage'], color='g', alpha=0.5, label='储能充电')
    ax3.bar(hours, -sim['storage_discharge'], color='r', alpha=0.5, label='储能放电')
    ax3.set_ylabel('充放电功率 (kW)', fontsize=14)

    # 设置图例
    lines, labels = ax2.get_legend_handles_labels()
    bars, bar_labels = ax3.get_legend_handles_labels()
    ax3.legend(lines + bars, labels + bar_labels, loc='upper right', fontsize=12)

    ax2.grid(True, linestyle='--', alpha=0.6)
    ax2.set_xlabel('时间 (小时)', fontsize=14)

    # 添加装机容量说明
    plt.figtext(0.5, 0.01,
                f"注: 光伏装机容量={pv_cap}kW, 风电装机容量={wind_cap}kW, 储能配置={opt_power}kW/{opt_capacity}kWh",
                ha='center', fontsize=10, alpha=0.7)

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])

    # 保存图表
    plt.savefig(f'园区{area}_最优储能配置.png', dpi=300, bbox_inches='tight')
    plt.show()


# 为所有园区绘制曲线图
for area in ['A', 'B', 'C']:
    plot_area_energy(area)