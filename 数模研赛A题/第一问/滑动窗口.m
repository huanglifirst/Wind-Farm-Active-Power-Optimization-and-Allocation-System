clear all;
% 读取 Excel 数据，假设文件名为 '1.xlsx'
filename = '1.xlsx';
data = readmatrix(filename);

% 获取数据维度
[numRows, numCols] = size(data);

% 初始化存储结果的矩阵
stress_amplitude = zeros(numRows-1, numCols); 
S_corrected = zeros(numRows-1, numCols);     
D_ti = zeros(numRows-1, numCols);            
D_total = zeros(1, numCols);                 
L_N_values = zeros(1, numCols);              
D_cumulative = zeros(numRows-1, numCols);    

% 常量定义
sigma_b = 50000000;  
m = 10;              
N = 42565440.4361;   

% 对每列单独操作
for col = 1:numCols
    load_data = data(:, col);
    
    % 计算应力幅值 S_ai
    for t = 2:numRows
        stress_amplitude(t-1, col) = abs(load_data(t) - load_data(t-1));
    end
    
    % Goodman 曲线修正
    for t = 1:numRows-1
        S_ai = stress_amplitude(t, col);
        S_mi = S_ai / 2;  
        S_corrected(t, col) = S_ai / (1 - (S_mi / sigma_b));  
    end
    
    % 计算等效疲劳载荷 L_N
    L_N = (sum(S_corrected(:, col).^m) / N)^(1/m);
    L_N_values(1, col) = L_N;  
    
    % 计算疲劳寿命 D_F
    D_F = (9.77 * 10^70) ./ (S_corrected(:, col).^m);
    
    % 计算累积疲劳损伤量 D_ti
    for t = 1:numRows-1
        D_ti(t, col) = (S_corrected(t, col).^m) / (9.77 * 10^70);
        
        % 逐秒累加总损伤值
        if t == 1
            D_cumulative(t, col) = D_ti(t, col);  
        else
            D_cumulative(t, col) = D_cumulative(t-1, col) + D_ti(t, col);  
        end
    end
    
    % 计算总损伤值 D_total
    D_total(1, col) = sum(D_ti(:, col));
end

% 保存结果到 Excel 文件
xlswrite('stress_amplitude1.xlsx', stress_amplitude);
xlswrite('S_corrected.xlsx', S_corrected);
xlswrite('D_ti3.xlsx', D_ti);
xlswrite('D_total.xlsx', D_total);
xlswrite('L_N_values2.xlsx', L_N_values);
xlswrite('D_cumulative2.xlsx', D_cumulative);

disp('每列数据的应力幅值、L_N、逐秒累加的总损伤值和疲劳损伤计算完成，结果已保存。');
