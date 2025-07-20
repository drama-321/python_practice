% clear;
% close all;
% clc;
%%
function PSO_2022_A_1()
    % 主函数 - 电工杯A题第一问PSO算法实现
    
    % 加载机组数据
    units = load_units_data();
    
    % 加载负荷数据
    load_demand = load_demand_data();
    
    % 准备时间轴
    time_points = cell(96, 1);
    for i = 0:95
        hour = floor(i/4);
        minute = 15 * mod(i,4);
        time_points{i+1} = sprintf('%02d:%02d', hour, minute);
    end
    
    % 初始化结果存储
    num_units = length(units);
    P_results = cell(num_units, 1);
    for i = 1:num_units
        P_results{i} = zeros(1, length(load_demand));
    end
    
    % 使用PSO算法执行经济调度
    for i = 1:length(load_demand)
        P = pso_economic_dispatch(load_demand(i), units);
        for j = 1:num_units
            P_results{j}(i) = P(j);
        end
    end
    
    disp('所有时段调度完成!');
    
    % 可视化结果
    figure('Position', [100, 100, 1200, 600]);
    plot(load_demand, 'k-', 'LineWidth', 2);
    hold on;
    
    colors = ['r-', 'g-', 'b-'];
    for j = 1:num_units
        plot(P_results{j}, colors(j), 'LineWidth', 1.5, ...
            'DisplayName', sprintf('%s (%d-%dMW)', units(j).name, units(j).P_min, units(j).P_max));
    end
    
    title('机组日发电计划曲线 (PSO优化)', 'FontSize', 16);
    ylabel('出力 (MW)');
    xlabel('时间');
    xticks(1:4:96);
    xticklabels(time_points(1:4:96));
    xtickangle(45);
    grid on;
    legend('show', 'Location', 'best');
    hold off;
    saveas(gcf, '第一问相关曲线图_PSO.png');
    
    % 计算总负荷电量（MWh）
    total_load_energy = sum(load_demand) * 0.25;  % 每个时间点0.25小时
    
    % 碳捕集单价列表（元/吨）
    carbon_prices = [0, 60, 80, 100];
    
    % 输出结果标题
    fprintf('\n结果如下\n\n');
    
    % 计算不同碳捕集单价下的成本
    for carbon_price = carbon_prices
        [operation_cost, carbon_cost] = calculate_thermal_cost(units, P_results, carbon_price);
        
        % 总发电成本（元）= 火电运行成本 + 碳捕集成本
        total_generation_cost = operation_cost + carbon_cost;
        
        % 单位供电成本（元/kWh）
        unit_supply_cost = total_generation_cost / (total_load_energy * 1000);
        
        % 转换为万元（用于输出）
        operation_cost_wan = operation_cost / 10000;
        carbon_cost_wan = carbon_cost / 10000;
        total_generation_cost_wan = total_generation_cost / 10000;
        
        % 输出结果
        fprintf('当碳捕集单价为 %d 元/t 时：\n', carbon_price);
        fprintf('  火电运行成本 = %s 万元\n', format_value(operation_cost_wan));
        fprintf('  碳捕集成本 = %s 万元\n', format_value(carbon_cost_wan));
        fprintf('  总发电成本 = %s 万元\n', format_value(total_generation_cost_wan));
        fprintf('  单位供电成本 = %s 元/kWh\n\n', format_value(unit_supply_cost));
    end
    
    % 保存机组出力数据为Excel文件
    % 关键修改：确保所有变量都是96×1列向量
    combined_power = P_results{1} + P_results{2} + P_results{3};
    T = table(time_points, load_demand, P_results{1}', P_results{2}', P_results{3}', combined_power', ...
        'VariableNames', {'时间', '系统负荷_MW', '机组1出力_MW', '机组2出力_MW', '机组3出力_MW', '机组1_2_3出力_MW'});
    
    writetable(T, '机组出力数据_PSO.xlsx');
    disp('机组出力数据已保存到: 机组出力数据_PSO.xlsx');
end

function units = load_units_data()
    % 加载机组参数数据（包含碳排放强度）
    units = struct(...
        'name', {'机组1', '机组2', '机组3'}, ...
        'P_max', {600, 300, 150}, ...
        'P_min', {180, 90, 45}, ...
        'a', {0.226, 0.588, 0.785}, ...
        'b', {30.42, 65.12, 139.6}, ...
        'c', {786.80, 451.32, 1049.50}, ...
        'emission', {0.72, 0.75, 0.79});
end

function load_demand = load_demand_data()
    % 从Excel文件中读取负荷数据 - 添加'VariableNamingRule','preserve'避免警告
    data = readtable('C:/Users/HP/Desktop/问题一数据.xlsx', 'Sheet', 'Sheet1', 'VariableNamingRule','preserve');
    
    % 确保数据列存在
    if size(data, 2) < 2
        error('Excel文件必须包含至少两列数据');
    end
    
    % 提取第二列数据（负荷功率(p.u.)）并保持为列向量
    load_demand = data{:, 2} * 900;  % 转换为实际功率（MW）
end

function best_power = pso_economic_dispatch(load, units, max_iter, num_particles, w, c1, c2)
    % PSO算法求解经济调度问题
    % 参数默认值
    if nargin < 3, max_iter = 100; end
    if nargin < 4, num_particles = 50; end
    if nargin < 5, w = 0.5; end
    if nargin < 6, c1 = 1.5; end
    if nargin < 7, c2 = 1.5; end
    
    num_units = length(units);
    
    % 获取机组出力上下限
    bounds = zeros(num_units, 2);
    for i = 1:num_units
        bounds(i, :) = [units(i).P_min, units(i).P_max];
    end
    
    % 初始化粒子群
    particles_p = zeros(num_particles, num_units);
    particles_v = zeros(num_particles, num_units);
    
    % 随机初始化位置和速度
    for i = 1:num_particles
        for j = 1:num_units
            particles_p(i, j) = bounds(j, 1) + (bounds(j, 2) - bounds(j, 1)) * rand();
            particles_v(i, j) = (-1 + 2*rand()) * (bounds(j, 2) - bounds(j, 1)) / 10.0;
        end
    end
    
    % 初始化个体最优位置和适应度
    pbest_p = particles_p;
    pbest_fitness = inf(1, num_particles);
    
    % 初始化全局最优
    gbest_p = zeros(1, num_units);
    gbest_fitness = inf;
    
    % 煤价（元/kg）
    coal_price = 700 / 1000;  % 700元/吨 = 0.7元/kg
    
    % 迭代优化
    for iter = 1:max_iter
        for i = 1:num_particles
            % 计算当前粒子的总出力
            total_power = sum(particles_p(i, :));
            
            % 计算惩罚项（功率不平衡惩罚）
            imbalance_penalty = 10000 * abs(total_power - load)^2;
            
            % 计算约束惩罚（出力越限惩罚）
            constraint_penalty = 0;
            for j = 1:num_units
                if particles_p(i, j) < bounds(j, 1)
                    constraint_penalty = constraint_penalty + 10000 * (bounds(j, 1) - particles_p(i, j))^2;
                elseif particles_p(i, j) > bounds(j, 2)
                    constraint_penalty = constraint_penalty + 10000 * (particles_p(i, j) - bounds(j, 2))^2;
                end
            end
            
            % 计算总煤耗成本（目标函数）
            total_cost = 0;
            for j = 1:num_units
                % 计算煤耗量（kg/h）
                F_hourly = units(j).a * particles_p(i, j)^2 + units(j).b * particles_p(i, j) + units(j).c;
                % 15分钟的煤耗成本（元）
                fuel_cost = F_hourly * 0.25 * coal_price;
                % 总运行成本 = 1.5 * 煤耗成本（包括运行维护成本）
                total_cost = total_cost + 1.5 * fuel_cost;
            end
            
            % 总适应度 = 总成本 + 惩罚项
            fitness = total_cost + imbalance_penalty + constraint_penalty;
            
            % 更新个体最优
            if fitness < pbest_fitness(i)
                pbest_fitness(i) = fitness;
                pbest_p(i, :) = particles_p(i, :);
            end
            
            % 更新全局最优
            if fitness < gbest_fitness
                gbest_fitness = fitness;
                gbest_p = particles_p(i, :);
            end
        end
        
        % 更新粒子速度和位置
        for i = 1:num_particles
            for j = 1:num_units
                % 更新速度
                r1 = rand();
                r2 = rand();
                cognitive = c1 * r1 * (pbest_p(i, j) - particles_p(i, j));
                social = c2 * r2 * (gbest_p(j) - particles_p(i, j));
                particles_v(i, j) = w * particles_v(i, j) + cognitive + social;
                
                % 更新位置
                particles_p(i, j) = particles_p(i, j) + particles_v(i, j);
                
                % 边界处理
                particles_p(i, j) = max(bounds(j, 1), min(particles_p(i, j), bounds(j, 2)));
            end
        end
    end
    
    best_power = gbest_p;
end

function [total_operation_cost, total_carbon_cost] = calculate_thermal_cost(units, P_results, carbon_price)
    % 计算火电总成本（包括运行成本和碳捕集成本）
    total_fuel_cost = 0.0;  % 煤耗成本（元）
    total_om_cost = 0.0;    % 运行维护成本（元）
    total_carbon_cost = 0.0; % 碳捕集成本（元）
    
    % 煤价（元/kg）
    coal_price = 700 / 1000;  % 700元/吨 = 0.7元/kg
    
    % 遍历每个时间点（15分钟间隔）
    num_time_points = length(P_results{1});
    for t = 1:num_time_points
        % 遍历每台机组
        for i = 1:length(units)
            P = P_results{i}(t);  % 机组在t时刻的出力（MW）
            
            % 计算煤耗量（kg/h）
            F_hourly = units(i).a * P^2 + units(i).b * P + units(i).c;
            
            % 15分钟的煤耗量（kg）
            F_15min = F_hourly * 0.25;
            
            % 煤耗成本（元）
            fuel_cost = F_15min * coal_price;
            total_fuel_cost = total_fuel_cost + fuel_cost;
            
            % 运行维护成本（元）= 0.5 * 煤耗成本
            om_cost = 0.5 * fuel_cost;
            total_om_cost = total_om_cost + om_cost;
            
            % 计算发电量（MWh）
            generation_mwh = P * 0.25;  % MW × 0.25h = MWh
            
            % 计算碳排放量（kg）
            carbon_emission = generation_mwh * 1000 * units(i).emission;
            
            % 碳捕集成本（元）
            carbon_cost = carbon_emission * (carbon_price / 1000);  % 碳捕集单价元/吨 = 元/1000kg
            total_carbon_cost = total_carbon_cost + carbon_cost;
        end
    end
    
    % 总运行成本 = 煤耗成本 + 运行维护成本
    total_operation_cost = total_fuel_cost + total_om_cost;
end

function str = format_value(value)
    % 格式化数值为三位有效数字
    if value == 0
        str = "0.00";
        return;
    end
    
    magnitude = floor(log10(abs(value)));
    
    if magnitude >= 3 || magnitude <= -3
        str = sprintf('%.3e', value);
    else
        decimals = max(0, 2 - magnitude);
        % 处理小数位数
        if decimals > 0
            str = sprintf(['%.' num2str(decimals) 'f'], value);
        else
            str = sprintf('%.0f', value);
        end
    end
end