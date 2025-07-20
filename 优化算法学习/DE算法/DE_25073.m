clear;
close all;
clc;

% 定义Rastrigin函数
function val = rastrigin(x)
    A = 10; % 默认A为10
    n = length(x);
    val = A * n + sum(x.^2 - A * cos(2 * pi * x));
end

% 差分进化(DE)算法实现
function [global_best_pos, global_best_val, convergence_curve] = de_rastrigin(dimensions, population_size, max_iter, bounds, F, CR)
    % 初始化种群
    population = bounds(1) + (bounds(2)-bounds(1)) .* rand(population_size, dimensions);
    fitness = arrayfun(@(i) rastrigin(population(i,:)), 1:population_size)';
    
    % 全局最优初始化
    [global_best_val, global_best_idx] = min(fitness);
    global_best_pos = population(global_best_idx, :);
    
    % 存储历史最优值用于绘图
    convergence_curve = zeros(max_iter, 1);
    
    % 创建动态可视化图形
    fig = figure('Position', [100, 100, 1200, 600], 'Color', 'w');
    
    % 设置动态可视化更新间隔
    update_interval = ceil(max_iter / 50); % 最多更新50次
    if update_interval < 1
        update_interval = 1;
    end
    
    % DE主循环
    for i = 1:max_iter
        % 创建新一代种群
        new_population = population;
        new_fitness = fitness;
        
        for n = 1:population_size
            % 选择三个不同的随机个体（不包括当前个体）
            candidates = 1:population_size;
            candidates(n) = [];
            idx = candidates(randperm(length(candidates), 3));
            a = population(idx(1), :);
            b = population(idx(2), :);
            c = population(idx(3), :);
            
            % 变异操作：DE/rand/1策略
            mutant = a + F * (b - c);
            
            % 交叉操作：二项交叉
            trial = population(n, :);
            cross_points = rand(1, dimensions) < CR;
            
            % 确保至少有一个维度发生交叉
            if ~any(cross_points)
                cross_points(randi(dimensions)) = true;
            end
            
            trial(cross_points) = mutant(cross_points);
            
            % 边界处理
            trial = min(max(trial, bounds(1)), bounds(2));
            
            % 计算试验向量的适应度
            trial_fitness = rastrigin(trial);
            
            % 选择操作：贪婪选择
            if trial_fitness < fitness(n)
                new_population(n, :) = trial;
                new_fitness(n) = trial_fitness;
                
                % 更新全局最优
                if trial_fitness < global_best_val
                    global_best_val = trial_fitness;
                    global_best_pos = trial;
                end
            end
        end
        
        % 更新种群
        population = new_population;
        fitness = new_fitness;
        convergence_curve(i) = global_best_val;
        
        % 动态可视化（每隔update_interval次迭代更新一次）
        if mod(i, update_interval) == 0 || i == 1 || i == max_iter
            % 清除当前图形
            clf;
            
            % 绘制种群位置（仅显示前两个维度）
            subplot(1, 2, 1);
            hold on;
            
            % 生成Rastrigin函数的等高线图（仅前两个维度）
            [X, Y] = meshgrid(linspace(bounds(1), bounds(2), 100));
            Z = zeros(size(X));
            for a = 1:size(X,1)
                for b = 1:size(X,2)
                    point = [X(a,b), Y(a,b), zeros(1, dimensions-2)];
                    Z(a,b) = rastrigin(point);
                end
            end
            
            % 绘制等高线
            contourf(X, Y, Z, 50, 'LineStyle', 'none');
            colormap(jet(256));
            colorbar;
            title(sprintf('Population Positions (Iteration %d/%d)', i, max_iter));
            xlabel('Dimension 1');
            ylabel('Dimension 2');
            axis([bounds(1) bounds(2) bounds(1) bounds(2)]);
            
            % 绘制种群个体位置
            scatter(population(:,1), population(:,2), 40, 'filled', ...
                    'MarkerFaceColor', [0.3, 0.75, 0.93], ...
                    'MarkerEdgeColor', 'k');
                
            % 绘制全局最优位置
            scatter(global_best_pos(1), global_best_pos(2), 120, 'p', ...
                    'MarkerFaceColor', [0.93, 0.69, 0.13], ...
                    'MarkerEdgeColor', 'k', ...
                    'LineWidth', 1.5);
                
            % 添加图例
            legend({'Population', 'Global Best'}, ...
                   'Location', 'northwest');
            
            % 绘制收敛曲线
            subplot(1, 2, 2);
            hold on;
            plot(1:i, convergence_curve(1:i), 'b-', 'LineWidth', 2);
            
            % 标记当前最优值
            plot(i, convergence_curve(i), 'ro', 'MarkerSize', 8, 'LineWidth', 1.5);
            
            title('Convergence Curve');
            xlabel('Iteration');
            ylabel('Best Fitness Value');
            grid on;
            set(gca, 'YScale', 'log');
            xlim([0 max_iter]);
            ylim([1e-10 1000]);
            
            % 添加信息文本
            text(0.05*max_iter, 0.9*max(ylim), ...
                 sprintf('Current Best: %.4e', convergence_curve(i)), ...
                 'FontSize', 12, 'FontWeight', 'bold');
                 
            % 显示DE参数
            text(0.05*max_iter, 0.7*max(ylim), ...
                 sprintf('F: %.2f, CR: %.2f', F, CR), ...
                 'FontSize', 10);
            
            hold off;
            
            % 立即刷新图形
            drawnow;
        end
        
        % 打印进度
        if mod(i, 10) == 0
            fprintf('Iteration %4d/%d, Best Value: %.4e\n', i, max_iter, global_best_val);
        end
    end
    
    fprintf('\nOptimization Complete!\n');
    fprintf('Global Best Position: [');
    fprintf('%.4f ', global_best_pos);
    fprintf(']\n');
    fprintf('Global Best Value: %.4e\n', global_best_val);
    
    % 保存最终结果图像
    saveas(fig, 'de_optimization_result.png');
end

% 主程序
dimensions = 30;      % 问题维度
population_size = 100; % 种群大小
max_iter = 500;       % 最大迭代次数
bounds = [-5.12, 5.12]; % Rastrigin函数的标准边界
F = 0.5;              % 变异因子 (推荐0.5-1.0)
CR = 0.8;             % 交叉概率 (推荐0.8-1.0)

% 运行DE算法（带动态可视化）
[best_position, best_value, convergence] = de_rastrigin(...
    dimensions, population_size, max_iter, bounds, F, CR);