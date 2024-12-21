% 加载数据并进行优化

% 设置文件路径
file_path = 'D:\\jm\\2.xlsx';  % 替换为实际的路径

% 风机数量和时间步数
N_t = 100;  % 风机数量
T = 2000;   % 每台风机2000个时间步

% 预分配变量以存储从Excel中读取的数据
Time = zeros(N_t, T);
Pref = zeros(N_t, T);
WindSpeed = zeros(N_t, T);
Tshaft = zeros(N_t, T);
Ft = zeros(N_t, T);
Pout = zeros(N_t, T);
PitchAngle = zeros(N_t, T);
OmegeR = zeros(N_t, T);
OmegeF = zeros(N_t, T);

% 从Excel中逐个工作表读取数据，工作表名称 WT_1 到 WT_100
for i = 1:N_t
    sheet_name = sprintf('WT_%d', i);  % 工作表名称 WT_1 到 WT_100
    data = readtable(file_path, 'Sheet', sheet_name, 'Range', 'A2:I2001'); % 读取第2到2001行，前9列
    Time(i,:) = data{:, 1};
    Pref(i,:) = data{:, 2};
    WindSpeed(i,:) = data{:, 3};
    Tshaft(i,:) = data{:, 4};
    Ft(i,:) = data{:, 5};
    Pout(i,:) = data{:, 6};
    PitchAngle(i,:) = data{:, 7};
    OmegeR(i,:) = data{:, 8};
    OmegeF(i,:) = data{:, 9};
end

% 机械效率和系统响应系数等参数设置
eta = 0.95; % 机械效率
alpha = 0.1; % 系统响应系数
% k1 = 1; % 主轴疲劳损伤系数
% k2 = 2.5; % 塔架疲劳损伤系数
k1 = 1.15e-21; % 主轴疲劳损伤系数（k2 = 3 * k）
k2 = 6.13e-24; % 塔架疲劳损伤系数
Ct = 0.7; % 推力系数
rho = 1.225; % 空气密度 kg/m^3
A = 64; % 扫掠面积 m^2
P_max = 5; % 风机额定功率 MW
w_shaft = 0.3; % 主轴疲劳损伤权重
w_thrust = 0.7; % 塔架疲劳损伤权重
P_current = Pref; % 当前调度功率设为Pref
P_t = sum(Pref, 1); % 总功率要求，设为每个时间点所有风机的功率之和

% 初始化优化后的参考功率矩阵
P_ref_opt_all = Pref; % 初始参考功率设为当前调度功率

% 初始化总疲劳损伤
D_total = 0;

% 优化循环：按时间步进行优化
for t = 1:T
    fprintf('优化时间步 %d / %d\n', t, T);
    
    % 总功率需求
    P_t_current = P_t(t);
    
    % 定义优化变量：P_ref(:, t)
    % 目标函数：最小化当前时间步的总疲劳损伤
    objective_t = @(P_ref_t) sum(k1 * (eta * OmegeR(:,t) .* P_ref_t + ...
        alpha * (OmegeR(:,t) .* P_ref_t - OmegeR(:,max(t-1,1)) .* P_ref_opt_all(:,max(t-1,1)))).^2) * w_shaft + ...
        sum(k2 * (Ct * 0.5 * rho * A * WindSpeed(:,t).^2 .* (P_ref_t / P_max)).^2) * w_thrust;
    
    % 约束：
    % 1. sum(P_ref_t) == P_t_current
    % 2. |P_ref_t - P_current(:,t)| <= 1
    % 3. 0 <= P_ref_t <= P_max
    
    % 等式约束：sum(P_ref_t) == P_t_current
    A_eq = ones(1, N_t);
    b_eq = P_t_current;
    
    % 不等式约束：P_ref_t - P_current(:,t) <= 1
    %               P_current(:,t) - P_ref_t <= 1
    A_ineq = [eye(N_t); -eye(N_t)];
    b_ineq = [ones(N_t,1); ones(N_t,1)];
    
    % 下界和上界
    lb = max(0, P_current(:,t) - 1);
    ub = min(P_max, P_current(:,t) + 1);
    
    % 使用 fmincon 进行优化
    options = optimoptions('fmincon', 'Algorithm', 'sqp', 'Display', 'none', 'MaxIterations', 100, 'OptimalityTolerance', 1e-6);
    [P_ref_opt_t, D_t] = fmincon(objective_t, P_ref_opt_all(:,t), A_ineq, b_ineq, A_eq, b_eq, lb, ub, [], options);
    
    % 更新 P_ref
    P_ref_opt_all(:,t) = P_ref_opt_t;
    
    % 累加总疲劳损伤
    D_total_3 = D_total + D_t;
end

% 显示优化完成后的总疲劳损伤
fprintf('优化完成。\n');
fprintf('优化后的总疲劳损伤最小值：%f\n', D_total_3);

% 将优化后的 P_ref 保存到新的 Excel 文件
% 创建表格，列名为 WT_1 到 WT_100
TurbineNames = arrayfun(@(x) sprintf('WT_%d', x), 1:N_t, 'UniformOutput', false);
P_ref_table1 = array2table(P_ref_opt_all', 'VariableNames', TurbineNames);



% 添加时间步列，如果需要
% 假设 Time 是一致的，可以选择不添加。如果需要，可以取消注释以下代码：
% Time_vector = Time(1, :)'; % 假设所有风机的时间步相同
% P_ref_table = [table(Time_vector, 'VariableNames', {'Time'}) P_ref_table];

% % 定义保存的文件路径
% output_file = 'C:\your\correct\path\to\optimized_P_ref.xlsx';  % 替换为实际的保存路径
% 
% % 写入 Excel 文件
% writetable(P_ref_table, output_file);

% fprintf('优化后的 P_ref 已成功保存到 %s\n', output_file);
