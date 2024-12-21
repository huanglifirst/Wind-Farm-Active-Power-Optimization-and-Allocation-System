% 读取 Excel 数据，假设文件名为 '1.xlsx'
filename = '1.xlsx';
data = xlsread(filename);

% 获取数据维度
[numRows, numCols] = size(data);

% 确保每列有100个数据点
if numRows ~= 100
    error('数据的行数应为100');
end

% 初始化存储结果的矩阵
stress_amplitude = zeros(50, numCols); % 用于存储应力幅值
S_corrected = zeros(50, numCols);      % 用于存储修正后的应力幅值
D_ti = zeros(50, numCols);             % 用于存储累积疲劳损伤量
D_total = zeros(1, numCols);           % 用于存储每列的总损伤量
L_N_values = zeros(1, numCols);        % 用于存储每列的 L_N 值
cumulative_damage = zeros(50, numCols); % 用于存储逐秒累积损伤值

% 常量定义
sigma_b = 50000000;  % 材料的极限强度
m = 10;              % 材料疲劳强度指数
N = 42565440.4361;   % 总应力循环次数

% 对每列单独操作
for col = 1:numCols
    % 获取当前设备的载荷数据
    load_data = data(:, col);
    
    % 第一步：对该列的100个数据分成50份，每份两个数据点进行相减求绝对值
    for t = 1:50
        stress_amplitude(t, col) = abs(load_data(2*t) - load_data(2*t-1));
    end
    
    % 第二步：对该列数据进行 Goodman 曲线修正，计算修正后的应力幅值 S_i
    for t = 1:50
        S_ai = stress_amplitude(t, col);
        S_mi = S_ai / 2;  % 计算应力均值
        S_corrected(t, col) = S_ai / (1 - (S_mi / sigma_b));  % 解方程求 S_i
    end
    
    % 第三步：对该列数据计算等效疲劳载荷 L_N
    L_N = (sum(S_corrected(:, col).^m) / N)^(1/m);
    L_N_values(col) = L_N;  % 保存 L_N 值

    % 第四步：计算每列的累积疲劳损伤量 D_ti
    for t = 1:50
        D_ti(t, col) = (S_corrected(t, col)^m) /(9.77 * 10^70);  % 累积疲劳损伤量
    end

    % 第四步：对该列数据计算疲劳寿命 D_F
   D_F = (9.77 * 10^70) ./ (S_corrected(:, col).^m);
    % 第五步：计算逐秒累积损伤值
cumulative_damage(:, col) = cumsum(D_ti(:, col));  % 逐秒累积
    % 总损伤值
    D_total(col) = sum(D_ti(:, col));  % 计算每列的总损伤值
end

% 保存结果
writematrix(stress_amplitude, 'stress_amplitude2.xlsx');
writematrix(S_corrected, 'S_corrected1.xlsx');
writematrix(D_ti, 'D_ti1.xlsx');
writematrix(D_total', 'D_total1.xlsx'); % 转置以保持列向量形式
writematrix(L_N_values', 'L_N_values1.xlsx'); % 转置以保持列向量形式
writematrix(cumulative_damage, 'cumulative_damage2.xlsx');
