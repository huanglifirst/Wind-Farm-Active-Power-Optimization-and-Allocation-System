% 设置文件路径
file_path_ref = '平均功率.xlsx';  % 存储功率分配的Excel文件
file_path_wind = '2.xlsx'; % 存储风速和转速数据的Excel文件

% 风机数量和时间步数
N_t = 100;  % 风机数量
T = 2000;   % 时间步数（秒）

% 读取Excel中的功率分配数据，2000秒 * 100台风机
P_ref = readmatrix(file_path_ref);  % 读取功率参考值数据

% 检查功率数据的尺寸是否为2000*100
[rows, cols] = size(P_ref);
if rows ~= 2000 || cols ~= 100
    error('功率参考值数据的尺寸不符合预期：应为 2000行 * 100列');
end

% 读取每个风机的风速和转速数据从2.xlsx文件的各个工作表（WT_1 到 WT_100）
WindSpeed = zeros(T, N_t);  % 初始化存储风速数据的矩阵
OmegeR = zeros(T, N_t);     % 初始化存储转速数据的矩阵

for i = 1:N_t
    sheet_name = sprintf('WT_%d', i);  % 工作表名称 WT_1 到 WT_100
    % 读取每台风机的风速和转速数据，假设风速在第3列，转速在第8列
    wind_data = readtable(file_path_wind, 'Sheet', sheet_name, 'Range', 'A2:I2001'); 
    WindSpeed(:, i) = wind_data{:, 3};  % 存储风速数据
    OmegeR(:, i) = wind_data{:, 8};     % 存储转速数据
end

% 设置相关常数
eta = 0.95;  % 机械效率
alpha = 0.1; % 系统响应系数
k1 = 1.15e-40; % 主轴疲劳损伤系数（k2 = 3 * k）
k2 = 6.13e-40; % 塔架疲劳损伤系数
Ct = 0.7;      % 推力系数
rho = 1.225; % 空气密度 (kg/m^3)
A =64;     % 扫掠面积 (m^2)
P_max = 5;   % 风机的额定功率 (MW)
w_shaft = 0.3; % 主轴疲劳损伤权重
w_thrust = 0.7; % 塔架疲劳损伤权重

% 初始化存储主轴和塔架疲劳损伤的矩阵
D_shaft = zeros(2000, 100);  % 主轴疲劳损伤
D_thrust = zeros(2000, 100); % 塔架疲劳损伤
D_total = zeros(2000, 100);  % 总疲劳损伤

% 计算每秒每台风机的主轴疲劳损伤和塔架疲劳损伤
for t = 1:2000
    for i = 1:100
        % 计算主轴扭矩 T_shaft
        if t > 1
            d_term = OmegeR(t,i) * P_ref(t,i) - OmegeR(t-1,i) * P_ref(t-1,i);
        else
            d_term = 0;
        end
        T_shaft = eta * OmegeR(t,i) * P_ref(t,i) + alpha * d_term;

        % 计算主轴疲劳损伤 D_shaft
        D_shaft(t, i) = k1 * T_shaft^2;

        % 计算塔架推力 F_thrust
        F_thrust = Ct * 0.5 * rho * A * WindSpeed(t,i)^2 * (P_ref(t,i) / P_max);

        % 计算塔架疲劳损伤 D_thrust
        D_thrust(t, i) = k2 * F_thrust^2;

        % 计算总疲劳损伤 D_total
        D_total(t, i) = w_shaft * D_shaft(t, i) + w_thrust * D_thrust(t, i);
    end
end


% 保存主轴疲劳损伤、塔架疲劳损伤和总疲劳损伤到Excel文件
writematrix(D_shaft, 'D_shaft2_p.xlsx', 'Sheet', 1);
writematrix(D_thrust, 'D_thrust2_p.xlsx', 'Sheet', 1);
writematrix(D_total, 'D_total2_p.xlsx', 'Sheet', 1);

% 显示每秒每个风机的主轴疲劳损伤和塔架疲劳损伤
for t = 1:T
    fprintf('第 %d 秒的主轴疲劳损伤 (D_shaft)：\n', t);
    disp(D_shaft(t, :));  % 显示每个风机在第t秒的主轴疲劳损伤
    
    fprintf('第 %d 秒的塔架疲劳损伤 (D_thrust)：\n', t);
    disp(D_thrust(t, :));  % 显示每个风机在第t秒的塔架疲劳损伤
    
    % 如果不需要输出全部 2000 秒，可以在此添加跳出条件
    % 比如只显示前 10 秒的数据
    if t >= 10
        break;
    end
end