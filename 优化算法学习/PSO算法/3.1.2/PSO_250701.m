clear;
close all;
clc;
% 定义Rastrigin函数
function val = rastrigin(x)
    A = 10; % 默认A为10
    n = length(x);
    val = A * n + sum(x.^2 - A * cos(2 * pi * x));
end

% 改进的PSO算法实现（带动态可视化）
function [global_best_pos, global_best_val, convergence_curve] = pso_rastrigin(dimensions, num_particles, max_iter, bounds, w_start, w_end, c1, c2)
    % 初始化粒子群
    particles_pos = bounds(1) + (bounds(2)-bounds(1)) .* rand(num_particles, dimensions);
    particles_vel = -1 + 2 * rand(num_particles, dimensions);
    personal_best_pos = particles_pos;
    personal_best_val = arrayfun(@(i) rastrigin(particles_pos(i,:)), 1:num_particles)';
    
    % 全局最优初始化
    [global_best_val, global_best_idx] = min(personal_best_val);
    global_best_pos = personal_best_pos(global_best_idx, :);
    
    % 存储历史最优值用于绘图
    convergence_curve = zeros(max_iter, 1);
    
    % 创建动态可视化图形
    fig = figure('Position', [100, 100, 1200, 600], 'Color', 'w');
    
    % 设置动态可视化更新间隔
    update_interval = ceil(max_iter / 50); % 最多更新50次
    if update_interval < 1
        update_interval = 1;
    end
    
    % 速度限制因子
    velocity_limit = 0.2 * (bounds(2) - bounds(1));
    
    % PSO主循环
    for i = 1:max_iter
        % 动态调整惯性权重（线性衰减）
        w = w_start - (w_start - w_end) * (i / max_iter);
        
        for n = 1:num_particles
            % 更新粒子速度
            r1 = rand(1, dimensions);
            r2 = rand(1, dimensions);
            cognitive = c1 * r1 .* (personal_best_pos(n,:) - particles_pos(n,:));
            social = c2 * r2 .* (global_best_pos - particles_pos(n,:));
            particles_vel(n,:) = w * particles_vel(n,:) + cognitive + social;
            
            % 速度限制
            particles_vel(n,:) = sign(particles_vel(n,:)) .* min(abs(particles_vel(n,:)), velocity_limit);
            
            % 更新粒子位置
            particles_pos(n,:) = particles_pos(n,:) + particles_vel(n,:);
            
            % 边界处理
            particles_pos(n,:) = min(max(particles_pos(n,:), bounds(1)), bounds(2));
            
            % 计算适应度
            fitness = rastrigin(particles_pos(n,:));
            
            % 更新个体最优
            if fitness < personal_best_val(n)
                personal_best_pos(n,:) = particles_pos(n,:);
                personal_best_val(n) = fitness;
                
                % 更新全局最优
                if fitness < global_best_val
                    global_best_val = fitness;
                    global_best_pos = particles_pos(n,:);
                    global_best_idx = n;  % 更新全局最优索引
                end
            end
        end
        
        convergence_curve(i) = global_best_val;
        
        % 动态可视化（每隔update_interval次迭代更新一次）
        if mod(i, update_interval) == 0 || i == 1 || i == max_iter
            % 清除当前图形
            clf;
            
            % 绘制粒子位置（仅显示前两个维度）
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
            title(sprintf('Particle Positions (Iteration %d/%d)', i, max_iter));
            xlabel('Dimension 1');
            ylabel('Dimension 2');
            axis([bounds(1) bounds(2) bounds(1) bounds(2)]);
            
            % 绘制粒子位置
            scatter(particles_pos(:,1), particles_pos(:,2), 40, 'filled', ...
                    'MarkerFaceColor', [0.3, 0.75, 0.93], ...
                    'MarkerEdgeColor', 'k');
                
            % 绘制个体最优位置
            scatter(personal_best_pos(:,1), personal_best_pos(:,2), 40, 'filled', ...
                    'MarkerFaceColor', [0.85, 0.33, 0.1], ...
                    'MarkerEdgeColor', 'k');
                
            % 绘制全局最优位置
            scatter(global_best_pos(1), global_best_pos(2), 120, 'p', ...
                    'MarkerFaceColor', [0.93, 0.69, 0.13], ...
                    'MarkerEdgeColor', 'k', ...
                    'LineWidth', 1.5);
                
            % 添加图例
            legend({'Particles', 'Personal Best', 'Global Best'}, ...
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
                 
            % 显示动态参数
            text(0.05*max_iter, 0.7*max(ylim), ...
                 sprintf('w: %.2f', w), ...
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
    saveas(fig, 'pso_optimization_result.png');
end

% 主程序
dimensions = 10;   % 问题维度
num_particles = 100;
max_iter = 200;   % 增加迭代次数以展示更多动态过程
bounds = [-5.12, 5.12];  % Rastrigin函数的标准边界
w_start = 0.9;    % 初始惯性权重
w_end = 0.4;      % 最终惯性权重
c1 = 1;         % 认知系数
c2 = 1;         % 社会系数

% 运行改进的PSO算法（带动态可视化）
[best_position, best_value, convergence] = pso_rastrigin(...
    dimensions, num_particles, max_iter, bounds, w_start, w_end, c1, c2);