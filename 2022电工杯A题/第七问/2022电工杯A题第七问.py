import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd
import matplotlib.dates as mdates
from datetime import datetime, timedelta

# 设置中文字体
font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf", size=12)


# ===================== 数据准备 =====================
def load_units_data():
    """加载机组参数数据（只保留1号机组）"""
    return [
        {'name': '机组1', 'P_max': 600, 'P_min': 180, 'a': 0.226, 'b': 30.42, 'c': 786.80, 'emission': 0.72}
    ]


def load_demand_data():
    """从Excel文件中读取15天负荷数据"""
    df = pd.read_excel('C:/Users/HP/Desktop/附件2.xlsx', sheet_name='Sheet1')
    load_demand = df['负荷功率(MW)'].values
    return load_demand.tolist()


def load_wind_data():
    """从Excel文件中读取15天风电数据（1200MW装机）"""
    df = pd.read_excel('C:/Users/HP/Desktop/附件2.xlsx', sheet_name='Sheet1')
    wind_power = df['风电功率(MW)'].values
    return wind_power.tolist()


# ===================== 主程序 =====================
def main():
    # 加载数据
    units = load_units_data()  # 只保留机组1
    load_demand = load_demand_data()
    wind_power = load_wind_data()  # 1200MW风电数据

    # 获取机组参数
    unit = units[0]
    P_min = unit['P_min']
    P_max = unit['P_max']

    # 创建时间轴（15天，每天96个15分钟间隔）
    start_date = datetime(2020, 7, 1)
    time_points = [start_date + timedelta(minutes=15 * i) for i in range(len(load_demand))]

    # 初始化结果存储
    thermal_power = []  # 火电机组出力
    power_balance = []  # 功率平衡值存储

    # 执行调度
    for i in range(len(load_demand)):
        load_val = load_demand[i]
        wind_val = wind_power[i]

        # 计算等效负荷（总负荷减去风电）
        equivalent_load = load_val - wind_val

        # 判断并调整火电出力
        if equivalent_load < P_min:
            equivalent_load = P_min
        elif equivalent_load > P_max:
            equivalent_load = P_max

        # 保存火电出力
        thermal_power.append(equivalent_load)

        # 计算总发电功率（火电+风电）
        total_generation = equivalent_load + wind_val

        # 计算功率平衡（总发电 - 负荷）
        balance = total_generation - load_val
        power_balance.append(balance)

    # 可视化结果 - 15天数据
    plt.figure(figsize=(18, 12))

    # 子图1: 发电计划曲线（15天）
    plt.subplot(2, 1, 1)
    plt.plot(time_points, load_demand, 'k-', linewidth=1.5, label='系统总负荷')
    plt.plot(time_points, thermal_power, 'r-', linewidth=1.5, label=f"机组1 ({P_min}-{P_max}MW)")
    plt.plot(time_points, wind_power, 'c-', linewidth=1.5, label='风电出力')

    # 添加总出力曲线
    total_power = [thermal_power[i] + wind_power[i] for i in range(len(thermal_power))]
    plt.plot(time_points, total_power, 'm--', linewidth=1.5, label='总出力')

    # 设置日期格式
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.gcf().autofmt_xdate()

    plt.title('15天发电计划曲线', fontproperties=font, fontsize=16)
    plt.ylabel('出力 (MW)', fontproperties=font)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')

    # 子图2: 功率平衡曲线（15天）
    plt.subplot(2, 1, 2)
    plt.plot(time_points, power_balance, 'b-', linewidth=1.5, label='功率平衡')

    # 设置日期格式
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.gcf().autofmt_xdate()

    plt.title('15天功率平衡曲线', fontproperties=font, fontsize=16)
    plt.ylabel('功率平衡 (MW)', fontproperties=font)
    plt.xlabel('日期', fontproperties=font)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(prop=font, loc='best')

    plt.tight_layout()
    plt.savefig('第七问相关曲线图.png', dpi=300)
    plt.show()


if __name__ == "__main__":
    main()