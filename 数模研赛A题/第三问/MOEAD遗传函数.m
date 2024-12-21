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
k1 = 1.15e-21; % 主轴疲劳损伤系数
k2 = 6.13e-24; % 塔架疲劳损伤系数
Ct = 0.7; % 推力系数
rho = 1.225; % 空气密度 kg/m^3
A = 64; % 扫掠面积 m^2
P_max = 5; % 风机额定功率 MW
w_shaft = 0.3; % 主轴疲劳损伤权重
w_thrust = 0.7; % 塔架疲劳损伤权重
P_current = Pref; % 当前调度功率设为Pref
P_t = sum(Pref, 1); % 总功率要求，设为每个时间点所有风机的功率之和

% 初始功率参考值（随机初始化）
P_ref_init = rand(N_t, T) * P_max;

% 使用遗传算法进行优化
options = optimoptions('ga', 'Display', 'iter', 'PopulationSize', 200, 'MaxGenerations', 100);

% 调用遗传算法
[P_ref_opt, fval] = ga(@(P_ref) objective(P_ref, WindSpeed, OmegeR, N_t, T, eta, alpha, k1, k2, Ct, rho, A, w_shaft, w_thrust, P_max), ...
                        N_t * T, [], [], [], [], zeros(N_t * T, 1), P_max * ones(N_t * T, 1), ...
                        @(P_ref) constraints(P_ref, Pref, P_t, N_t, T, P_max), options);

% 显示优化结果
disp('优化后的疲劳损伤最小值：');
disp(fval);

% -------- 函数定义在文件的末
