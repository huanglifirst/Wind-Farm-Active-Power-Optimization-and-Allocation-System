clear all;
% 读取 Excel 数据，假设文件名为 '1.xlsx'
filename = '22.xlsx';
%data = xlsread(filename);
data = readmatrix(filename);
% 获取数据维度
[numRows, numCols] = size(data);

% 初始化存储结果的矩阵
stress_amplitude = zeros(numRows-1, numCols); % 用于存储应力幅值
S_corrected = zeros(numRows-1, numCols);     % 用于存储修正后的应力幅值
D_ti = zeros(numRows-1, numCols);            % 用于存储累积疲劳损伤量
D_total = zeros(1, numCols);                 % 用于存储每列的总损伤量
L_N_values = zeros(1, numCols);              % 用于存储每列的 L_N 值
D_cumulative = zeros(numRows-1, numCols);    % 用于存储逐秒累加的总损伤值

% 常量定义
sigma_b = 50000000;  % 材料的极限强度
m = 10;              % 材料疲劳强度指数
N = 42565440.4361;   % 总应力循环次数

% 对每列单独操作
for col = 1:numCols
    % 获取当前设备的载荷数据
    load_data = data(:, col);
    
    % 第一步：对该列数据计算应力幅值 S_ai
    for t = 2:numRows
        % 不重叠滑动窗口法计算应力幅值 S_ai = |L_t - L_{t-1}|
        stress_amplitude(t-1, col) = abs(load_data(t) - load_data(t-1));
    end
    
    % 第二步：对该列数据进行 Goodman 曲线修正，计算修正后的应力幅值 S_i
    for t = 1:numRows-1
        S_ai = stress_amplitude(t, col);
        S_mi = S_ai / 2;  % 计算应力均值
        S_corrected(t, col) = S_ai / (1 - (S_mi / sigma_b));  % 解方程求 S_i
    end
    
    % 第三步：对该列数据计算等效疲劳载荷 L_N
    L_N = (sum(S_corrected(:, col).^m) / N)^(1/m);
    L_N_values(1, col) = L_N;  % 将该列的 L_N 值保存到矩阵中
    
    % 第四步：对该列数据计算疲劳寿命 D_F
    D_F = (9.77 * 10^70) ./ (S_corrected(:, col).^m);
    
    % 第五步：对该列数据计算累积疲劳损伤量 D_ti
    for t = 1:numRows-1
        D_ti(t, col) = (S_corrected(t, col).^m) / (9.77 * 10^70);
        
        % 第七步：逐秒累加总损伤值
        if t == 1
            D_cumulative(t, col) = D_ti(t, col);  % 第一秒
        else
            D_cumulative(t, col) = D_cumulative(t-1, col) + D_ti(t, col);  % 逐秒累加
        end
    end
    
    % 第六步：计算该列数据的总损伤值 D_total
    D_total(1, col) = sum(D_ti(:, col));
end

% 将各列的结果分别保存到 Excel 文件
xlswrite('塔架差值.xlsx', stress_amplitude);
xlswrite('S_corrected5.xlsx', S_corrected);
xlswrite('塔架每秒.xlsx', D_ti);
xlswrite('塔架总损.xlsx', D_total);
xlswrite('等效疲劳载荷.xlsx', L_N_values);  % 保存每列的 L_N 值
xlswrite('累加.xlsx', D_cumulative);  % 保存逐秒累加的总损伤值

% 显示处理完成的信息
disp('每列数据的应力幅值、L_N、逐秒累加的总损伤值和疲劳损伤计算完成，结果已保存。');
