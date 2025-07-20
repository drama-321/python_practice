import numpy as np
import matplotlib.pyplot as plt


# 定义Rastrigin函数
def rastrigin(x, A=10):
    """计算Rastrigin函数值"""
    return A * len(x) + np.sum(x ** 2 - A * np.cos(2 * np.pi * x))


# PSO算法实现
def pso_rastrigin(dimensions, num_particles, max_iter, bounds, w=0.5, c1=1, c2=1):
    """
    dimensions ： 问题维度
    num_particles ： 粒子数量
    max_iter ： 最大迭代次数
    bounds ： 变量边界 [min, max]
    """
    # 初始化粒子群
    particles_p = np.random.uniform(bounds[0], bounds[1], (num_particles, dimensions))
    particles_v = np.random.uniform(-1, 1, (num_particles, dimensions))
    pbest_p = particles_p.copy()
    pbest_v = np.array([rastrigin(p) for p in particles_p])

    # 全局最优初始化
    gbest_idx = np.argmin(pbest_v)
    gbest_p = pbest_p[gbest_idx].copy()
    gbest_v = pbest_v[gbest_idx]

    # 存储历史最优值用于绘图
    convergence_curve = np.zeros(max_iter)

    # PSO主循环
    for i in range(max_iter):
        for n in range(num_particles):
            # 更新粒子速度
            r1, r2 = np.random.rand(2)
            cognitive = c1 * r1 * (pbest_p[n] - particles_p[n])
            social = c2 * r2 * (gbest_p - particles_p[n])
            particles_v[n] = w * particles_v[n] + cognitive + social

            # 更新粒子位置
            particles_p[n] += particles_v[n]

            # 边界处理
            particles_p[n] = np.clip(particles_p[n], bounds[0], bounds[1])

            # 计算适应度
            fitness = rastrigin(particles_p[n])

            # 更新个体最优
            if fitness < pbest_v[n]:
                pbest_p[n] = particles_p[n].copy()
                pbest_v[n] = fitness

                # 更新全局最优
                if fitness < gbest_v:
                    gbest_p = particles_p[n].copy()
                    gbest_v = fitness

        convergence_curve[i] = gbest_v

        # 打印进度
        if i % 10 == 0:
            print(f"Iteration {i:4d}, Best Value: {gbest_v:.4f}")

    print("\nOptimization Complete!")
    print(f"Global Best Position: {gbest_p}")
    print(f"Global Best Value: {gbest_v}")

    return gbest_p, gbest_v, convergence_curve


# 参数设置
dimensions = 10  # 问题维度
num_particles = 50
max_iter = 100
bounds = [-5.12, 5.12]  # 函数边界

# 运行PSO算法
best_position, best_value, convergence = pso_rastrigin(
    dimensions=dimensions,
    num_particles=num_particles,
    max_iter=max_iter,
    bounds=bounds
)

# 绘制收敛曲线
plt.figure(figsize=(10, 6))
plt.plot(convergence, 'b-', linewidth=2)
plt.title('PSO')
plt.xlabel('Iteration')
plt.ylabel('Best Fitness Value')
plt.grid(True)
plt.savefig('pso_convergence.png')
plt.show()