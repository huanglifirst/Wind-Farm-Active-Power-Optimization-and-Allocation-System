% 加载数据并进行优化

% 设置文件路径

filename = 'WF1_1.xlsx';
%data = xlsread(filename);
data = readmatrix(filename);
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

% 从Excel中逐个工作表读取数据
for i = 1:N_t
    sheet_name = sprintf('WT%d', i);  % 工作表名称 WT1 到 WT100
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
k1 = 1; % 主轴疲劳损伤系数
k2 = 1; % 塔架疲劳损伤系数
Ct = 1; % 推力系数
rho = 1.225; % 空气密度 kg/m^3
A = 100; % 扫掠面积 m^2
P_max = 5; % 风机额定功率 MW
w_shaft = 1; % 主轴疲劳损伤权重
w_thrust = 1; % 塔架疲劳损伤权重
P_current = Pref; % 当前调度功率设为Pref
P_t = sum(Pref, 1); % 总功率要求，设为每个时间点所有风机的功率之和

% 初始功率参考值（随机初始化）
P_ref_init = rand(N_t, T) * P_max;

% 使用遗传算法进行优化
options = optimoptions('ga', 'Display', 'iter', 'PopulationSize', 200, 'MaxGenerations', 100);

% 调用遗传算法
[P_ref_opt, fval] = ga(@(P_ref) objective(P_ref, WindSpeed, OmegeR, N_t, T, eta, alpha, k1, k2, Ct, rho, A, w_shaft, w_thrust), ...
                        N_t * T, [], [], [], [], zeros(N_t * T, 1), P_max * ones(N_t * T, 1), ...
                        @(P_ref) constraints(P_ref, Pref, P_t, N_t, T, P_max), options);

% 显示优化结果
disp('优化后的疲劳损伤最小值：');
disp(fval);

% -------- 函数定义在文件的末尾 -------- %

% 优化目标函数
function D_total = objective(P_ref, V, omega, N_t, T, eta, alpha, k1, k2, Ct, rho, A, w_shaft, w_thrust)
    D_total = 0; % 初始化总疲劳损伤
    P_ref = reshape(P_ref, [N_t, T]); % 调整P_ref的形状
    for t = 1:T
        D_shaft_total = 0;
        D_thrust_total = 0;
        for i = 1:N_t
            % 主轴扭矩计算
            if t > 1
                d_term = (omega(i,t) * P_ref(i,t) - omega(i,t-1) * P_ref(i,t-1));
            else
                d_term = 0;
            end
            T_shaft = eta * omega(i,t) * P_ref(i,t) + alpha * d_term;

            % 主轴疲劳损伤
            D_shaft = k1 * T_shaft^2;
            D_shaft_total = D_shaft_total + D_shaft * w_shaft;

            % 塔架推力计算
            F_thrust = Ct * 0.5 * rho * A * V(i,t)^2 * (P_ref(i,t) / P_max);

            % 塔架疲劳损伤
            D_thrust = k2 * F_thrust^2;
            D_thrust_total = D_thrust_total + D_thrust * w_thrust;
        end
        % 每个时间步的总疲劳损伤
        D_total = D_total + D_shaft_total + D_thrust_total;
    end
end

% 约束函数
function [c, ceq] = constraints(P_ref, P_current, P_t, N_t, T, P_max)
    P_ref = reshape(P_ref, [N_t, T]); % 调整P_ref的形状
    ceq = zeros(T, 1); % 初始化等式约束
    for t = 1:T
        % 总功率平衡约束
        ceq(t) = sum(P_ref(:,t)) - P_t(t); % 每个时间步的功率平衡
    end
    
    % 不等式约束：|P_ref,i(t) - P_current,t(t)| <= 1 和 0 <= P_ref,i(t) <= 5
    c = zeros(N_t * T, 1); % 初始化不等式约束
    count = 1;
    for i = 1:N_t
        for t = 1:T
            % 限制 |P_ref,i(t) - P_current,t(t)| <= 1
            c(count) = abs(P_ref(i,t) - P_current(i,t)) - 1;
            count = count + 1;
        end
    end
end
