clear; close all; clc;

% 定义所有园区的装机容量
capacities = struct(...
    'A', struct('pv', 750, 'wind', 0), ...  % 园区A只有光伏
    'B', struct('pv', 0, 'wind', 1000), ... % 园区B只有风电
    'C', struct('pv', 600, 'wind', 500));   % 园区C有光伏和风电

% 电价参数
electricity_prices = struct(...
    'pv', 0.4, ...   % 光伏购电成本 (元/kWh)
    'wind', 0.5, ... % 风电购电成本 (元/kWh)
    'grid', 1.0);    % 主网购电价格 (元/kWh)

% 固定储能参数
storage_params = struct(...
    'soc_min', 10, ...       % SOC下限 (%)
    'soc_max', 90, ...       % SOC上限 (%)
    'efficiency', 0.95, ...  % 充放电效率
    'power_cost', 800, ...   % 功率单价 (元/kW)
    'energy_cost', 1800, ... % 能量单价 (元/kWh)
    'lifetime', 10);         % 运行寿命 (年)

% 读取数据 - 修复缺失值问题
load_data = readtable('C:/Users/HP/Desktop/附件1：各园区典型日负荷数据.xlsx', 'VariableNamingRule', 'preserve');
renewable_data = readtable('C:/Users/HP/Desktop/附件2：各园区典型日风光发电数据.xlsx', ...
    'Range', 'A3:E26', 'VariableNamingRule', 'preserve', 'ReadVariableNames', false);

% 重命名列
load_data.Properties.VariableNames = {'Time', 'A_load', 'B_load', 'C_load'};
renewable_data.Properties.VariableNames = {'Time', 'A_pv', 'B_wind', 'C_pv', 'C_wind'};

% 确保时间列格式一致
load_data.Time = string(load_data.Time);
renewable_data.Time = string(renewable_data.Time);

% 合并数据
date = innerjoin(load_data, renewable_data, 'Keys', 'Time');

% 添加缺失的列
date.A_wind = zeros(height(date), 1);
date.B_pv = zeros(height(date), 1);

% 计算实际发电功率
areas = {'A', 'B', 'C'};
for i = 1:length(areas)
    area = areas{i};
    pv_col = sprintf('%s_pv', area);
    wind_col = sprintf('%s_wind', area);
    
    date.(sprintf('%s_pv_power', area)) = capacities.(area).pv * date.(pv_col);
    date.(sprintf('%s_wind_power', area)) = capacities.(area).wind * date.(wind_col);
    date.(sprintf('%s_total_power', area)) = date.(sprintf('%s_pv_power', area)) + ...
                                            date.(sprintf('%s_wind_power', area));
end

% ======================= PSO优化部分 =======================
function [total_cost, simulation] = simulate_storage(area, power, capacity, date, electricity_prices, storage_params)
    % 模拟给定储能配置下的运行情况
    n = height(date);
    
    % 初始化储能状态
    storage_soc = 90.0;  % 初始SOC 90%
    
    % 初始化结果结构
    simulation = struct(...
        'renew_used', zeros(n, 1), ...
        'renew_to_storage', zeros(n, 1), ...
        'storage_discharge', zeros(n, 1), ...
        'curtail_pv', zeros(n, 1), ...
        'curtail_wind', zeros(n, 1), ...
        'grid_purchase', zeros(n, 1), ...
        'soc', zeros(n, 1), ...
        'pv_used', zeros(n, 1), ...  % 记录光伏利用量
        'wind_used', zeros(n, 1));   % 记录风电利用量
    
    % 无效配置检查
    if power == 0 && capacity > 0
        total_cost = inf;
        return;
    end
    
    % 计算储能投资成本
    storage_investment = power * storage_params.power_cost + capacity * storage_params.energy_cost;
    storage_daily_cost = storage_investment / (storage_params.lifetime * 365);
    
    % 模拟24小时运行
    for i = 1:n
        load_col = sprintf('%s_load', area);
        load = date.(load_col)(i);
        
        pv_power_col = sprintf('%s_pv_power', area);
        pv_gen = date.(pv_power_col)(i);
        
        wind_power_col = sprintf('%s_wind_power', area);
        wind_gen = date.(wind_power_col)(i);
        
        % 当前储能状态 (kWh)
        soc_kwh = storage_soc / 100 * capacity;
        
        % 1. 优先使用光伏发电
        pv_used = min(pv_gen, load);
        load_after_pv = load - pv_used;
        
        % 2. 然后使用风电
        wind_used = min(wind_gen, load_after_pv);
        load_remain = load_after_pv - wind_used;
        
        renew_used = pv_used + wind_used;
        simulation.renew_used(i) = renew_used;
        simulation.pv_used(i) = pv_used;
        simulation.wind_used(i) = wind_used;
        
        % 3. 计算剩余可再生能源
        pv_surplus = pv_gen - pv_used;
        wind_surplus = wind_gen - wind_used;
        total_surplus = pv_surplus + wind_surplus;
        
        % 4. 剩余可再生能源处理
        if total_surplus > 0
            % 计算最大充电功率
            if capacity > 0
                max_charge_kw = min(power, ...
                    (storage_params.soc_max/100 * capacity - soc_kwh) / storage_params.efficiency);
            else
                max_charge_kw = 0;
            end
            
            charge_kw = min(total_surplus, max_charge_kw);
            charge_actual = charge_kw * storage_params.efficiency;
            
            % 更新SOC
            soc_kwh = soc_kwh + charge_actual;
            simulation.renew_to_storage(i) = charge_kw;
            
            % 弃电分配（优先弃风电）
            curtail_total = total_surplus - charge_kw;
            wind_curtail = min(wind_surplus, curtail_total);
            pv_curtail = curtail_total - wind_curtail;
            simulation.curtail_pv(i) = pv_curtail;
            simulation.curtail_wind(i) = wind_curtail;
        else
            simulation.renew_to_storage(i) = 0;
            simulation.curtail_pv(i) = pv_surplus;
            simulation.curtail_wind(i) = wind_surplus;
        end
        
        % 5. 储能补充
        if load_remain > 0 && capacity > 0
            % 计算最大放电功率
            max_discharge_kw = min(power, ...
                (soc_kwh - storage_params.soc_min/100 * capacity) * storage_params.efficiency);
            discharge_kw = min(load_remain, max_discharge_kw);
            discharge_actual = discharge_kw / storage_params.efficiency;
            
            % 更新SOC
            soc_kwh = soc_kwh - discharge_actual;
            load_remain = load_remain - discharge_kw;
            simulation.storage_discharge(i) = discharge_kw;
        else
            simulation.storage_discharge(i) = 0;
        end
        
        % 6. 电网补充
        simulation.grid_purchase(i) = max(0, load_remain);
        
        % 7. 记录SOC
        if capacity > 0
            storage_soc = soc_kwh / capacity * 100;
        else
            storage_soc = 0;
        end
        simulation.soc(i) = storage_soc;
    end
    
    % 计算成本
    pv_used_total = sum(simulation.pv_used);
    wind_used_total = sum(simulation.wind_used);
    
    renew_cost = pv_used_total * electricity_prices.pv + wind_used_total * electricity_prices.wind;
    grid_cost = sum(simulation.grid_purchase) * electricity_prices.grid;
    total_cost = renew_cost + grid_cost + storage_daily_cost;
end

function [best_config, best_results, sim_data] = pso_optimize_storage(area, date, capacities, electricity_prices, storage_params)
    % PSO参数
    num_particles = 50;
    max_iter = 100;
    w = 0.8;
    c1 = 1.5;
    c2 = 1.5;
    
    % 定义搜索边界
    power_bounds = [0, 200];    % 功率范围 (kW)
    capacity_bounds = [0, 400]; % 容量范围 (kWh)
    
    % 初始化粒子群
    particles_p = zeros(num_particles, 2); % 位置
    particles_v = zeros(num_particles, 2); % 速度
    
    % 随机初始化位置和速度
    particles_p(:, 1) = power_bounds(1) + (power_bounds(2) - power_bounds(1)) * rand(num_particles, 1);
    particles_p(:, 2) = capacity_bounds(1) + (capacity_bounds(2) - capacity_bounds(1)) * rand(num_particles, 1);
    particles_v = -1 + 2 * rand(num_particles, 2);
    
    % 初始化个体最优
    pbest_p = particles_p; % 个体最优位置
    pbest_cost = inf(num_particles, 1); % 个体最优成本
    
    % 初始化全局最优
    gbest_p = zeros(1, 2); % 全局最优位置
    gbest_cost = inf;      % 全局最优成本
    gbest_simulation = []; % 全局最优模拟结果
    
    % 评估初始粒子群
    for i = 1:num_particles
        power = particles_p(i, 1);
        capacity = particles_p(i, 2);
        
        [cost, sim] = simulate_storage(area, power, capacity, date, electricity_prices, storage_params);
        
        pbest_cost(i) = cost;
        if cost < gbest_cost
            gbest_cost = cost;
            gbest_p = particles_p(i, :);
        end
    end
    
    % PSO主循环
    convergence = zeros(max_iter, 1);
    for iter = 1:max_iter
        for i = 1:num_particles
            % 更新速度和位置
            r1 = rand();
            r2 = rand();
            
            particles_v(i, :) = w * particles_v(i, :) + ...
                c1 * r1 * (pbest_p(i, :) - particles_p(i, :)) + ...
                c2 * r2 * (gbest_p - particles_p(i, :));
            
            particles_p(i, :) = particles_p(i, :) + particles_v(i, :);
            
            % 应用边界约束
            particles_p(i, 1) = min(max(particles_p(i, 1), power_bounds(1)), power_bounds(2));
            particles_p(i, 2) = min(max(particles_p(i, 2), capacity_bounds(1)), capacity_bounds(2));
            
            % 评估新位置
            power = particles_p(i, 1);
            capacity = particles_p(i, 2);
            
            [cost, sim] = simulate_storage(area, power, capacity, date, electricity_prices, storage_params);
            
            % 更新个体最优
            if cost < pbest_cost(i)
                pbest_cost(i) = cost;
                pbest_p(i, :) = particles_p(i, :);
                
                % 更新全局最优
                if cost < gbest_cost
                    gbest_cost = cost;
                    gbest_p = particles_p(i, :);
                    gbest_simulation = sim;
                end
            end
        end
        convergence(iter) = gbest_cost;
    end
    
    % 提取最优配置结果
    power_opt = max(0, round(gbest_p(1))); % 取整处理
    capacity_opt = max(0, round(gbest_p(2)));
    
    % 重新计算最优配置的详细结果
    load_col = sprintf('%s_load', area);
    total_load = sum(date.(load_col));
    
    pv_used_total = sum(gbest_simulation.pv_used);
    wind_used_total = sum(gbest_simulation.wind_used);
    curtail_total = sum(gbest_simulation.curtail_pv) + sum(gbest_simulation.curtail_wind);
    grid_purchase_total = sum(gbest_simulation.grid_purchase);
    
    % 计算成本
    renew_cost = pv_used_total * electricity_prices.pv + wind_used_total * electricity_prices.wind;
    grid_cost = grid_purchase_total * electricity_prices.grid;
    
    % 储能投资成本
    storage_investment = power_opt * storage_params.power_cost + capacity_opt * storage_params.energy_cost;
    storage_daily_cost = storage_investment / (storage_params.lifetime * 365);
    
    % 整理结果
    best_results = struct(...
        'optimal_power', power_opt, ...
        'optimal_capacity', capacity_opt, ...
        'total_load', total_load, ...
        'curtailment', curtail_total, ...
        'grid_purchase', grid_purchase_total, ...
        'storage_charge', sum(gbest_simulation.renew_to_storage), ...
        'storage_discharge', sum(gbest_simulation.storage_discharge), ...
        'pv_used', pv_used_total, ...
        'wind_used', wind_used_total, ...
        'total_cost', gbest_cost, ...
        'renew_cost', renew_cost, ...
        'grid_cost', grid_cost, ...
        'storage_investment', storage_investment, ...
        'storage_daily_cost', storage_daily_cost);
    
    best_config = [power_opt, capacity_opt];
    sim_data = gbest_simulation;
end

function plot_area_energy(area, date, capacities, optimal_configs, sim_data)
    % 绘制园区能源曲线图
    hours = 1:height(date);
    
    % 获取当前园区的装机容量
    pv_cap = capacities.(area).pv;
    wind_cap = capacities.(area).wind;
    
    % 获取最优配置
    opt_power = optimal_configs.(area).optimal_power;
    opt_capacity = optimal_configs.(area).optimal_capacity;
    
    % 获取模拟数据
    sim = sim_data.(area);
    
    % 创建图形
    figure('Position', [100, 100, 1000, 800]);
    
    % 第一个子图：发电和负荷曲线
    subplot(2, 1, 1);
    hold on;
    
    % 绘制负荷曲线
    load_col = sprintf('%s_load', area);
    plot(hours, date.(load_col), 'k-', 'LineWidth', 2.5, 'DisplayName', '负荷');
    
    % 绘制光伏曲线
    if pv_cap > 0
        pv_power_col = sprintf('%s_pv_power', area);
        plot(hours, date.(pv_power_col), 'Color', [1, 0.8, 0], 'LineWidth', 2, 'DisplayName', '光伏发电');
    end
    
    % 绘制风电曲线
    if wind_cap > 0
        wind_power_col = sprintf('%s_wind_power', area);
        plot(hours, date.(wind_power_col), 'Color', [0.25, 0.41, 0.88], 'LineWidth', 2, 'DisplayName', '风电发电');
    end
    
    % 绘制总发电曲线
    total_power_col = sprintf('%s_total_power', area);
    plot(hours, date.(total_power_col), 'g--', 'LineWidth', 1.8, 'DisplayName', '总发电');
    
    % 绘制弃光曲线
    if pv_cap > 0
        plot(hours, sim.curtail_pv, 'r--', 'LineWidth', 1.8, 'DisplayName', '弃光');
    end
    
    % 绘制弃风曲线
    if wind_cap > 0
        plot(hours, sim.curtail_wind, 'm--', 'LineWidth', 1.8, 'DisplayName', '弃风');
    end
    
    % 设置图表属性
    title(sprintf('园区%s - 最优储能配置: %dkW/%dkWh', area, opt_power, opt_capacity), 'FontSize', 16);
    ylabel('功率 (kW)', 'FontSize', 12);
    legend('Location', 'best', 'FontSize', 10);
    grid on;
    set(gca, 'XTick', hours, 'XTickLabel', date.Time, 'FontSize', 10);
    xtickangle(45);
    xlim([1, length(hours)]);
    
    % 第二个子图：储能状态
    subplot(2, 1, 2);
    hold on;
    
    % 绘制储能SOC曲线
    yyaxis left;
    plot(hours, sim.soc, 'b-', 'LineWidth', 2.5, 'DisplayName', '储能SOC');
    ylabel('SOC (%)', 'FontSize', 12);
    ylim([0, 100]);
    
    % 绘制储能充放电曲线
    yyaxis right;
    bar(hours, sim.renew_to_storage, 'FaceColor', 'g', 'FaceAlpha', 0.5, 'DisplayName', '储能充电');
    bar(hours, -sim.storage_discharge, 'FaceColor', 'r', 'FaceAlpha', 0.5, 'DisplayName', '储能放电');
    ylabel('充放电功率 (kW)', 'FontSize', 12);
    
    % 设置图例
    legend({'储能SOC', '储能充电', '储能放电'}, 'Location', 'best', 'FontSize', 10);
    
    % 设置公共属性
    grid on;
    set(gca, 'XTick', hours, 'XTickLabel', date.Time, 'FontSize', 10);
    xtickangle(45);
    xlabel('时间 (小时)', 'FontSize', 12);
    xlim([1, length(hours)]);
    
    % 添加注释
    annotation('textbox', [0.1, 0.01, 0.8, 0.03], 'String', ...
        sprintf('注: 光伏装机容量=%dkW, 风电装机容量=%dkW, 储能配置=%dkW/%dkWh', ...
        pv_cap, wind_cap, opt_power, opt_capacity), ...
        'HorizontalAlignment', 'center', 'FontSize', 9, 'EdgeColor', 'none');
    
    % 保存图表
    saveas(gcf, sprintf('园区%s_最优储能配置.png', area));
end

% ======================= 主程序 =======================
% 优化每个园区的储能配置并存储结果
optimal_configs = struct();
simulation_data = struct();
areas = {'A', 'B', 'C'};

for i = 1:length(areas)
    area = areas{i};
    [best_config, best_results, sim_data] = pso_optimize_storage(...
        area, date, capacities, electricity_prices, storage_params);
    
    optimal_configs.(area) = best_results;
    simulation_data.(area) = sim_data;
    
    fprintf('园区%s优化完成: %dkW/%dkWh, 成本: %.2f元/天\n', ...
        area, best_config(1), best_config(2), best_results.total_cost);
end

% 打印优化结果
for i = 1:length(areas)
    area = areas{i};
    config = optimal_configs.(area);
    
    fprintf('\n园区%s最优配置: %dkW/%dkWh\n', area, config.optimal_power, config.optimal_capacity);
    fprintf('总供电成本: %.2f元/天\n', config.total_cost);
    fprintf('可再生能源成本: %.2f元/天\n', config.renew_cost);
    fprintf('网购电成本: %.2f元/天\n', config.grid_cost);
    fprintf('光伏利用量: %.2fkWh\n', config.pv_used);
    fprintf('风电利用量: %.2fkWh\n', config.wind_used);
    fprintf('网购电量: %.2fkWh\n', config.grid_purchase);
    fprintf('弃电量: %.2fkWh\n', config.curtailment);
end

% 为所有园区绘制曲线图
for i = 1:length(areas)
    area = areas{i};
    plot_area_energy(area, date, capacities, optimal_configs, simulation_data);
end