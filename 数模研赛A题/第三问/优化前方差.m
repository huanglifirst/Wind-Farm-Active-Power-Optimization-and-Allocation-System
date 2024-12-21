% 1. 从Excel文件中读取数据
file_path = 'D:\\jm\\第三问1.xlsx';  % Excel文件路径
data = readmatrix(file_path);  % 读取整个Excel文件中的数据

% 确保数据是2000x100的矩阵
[numRows, numCols] = size(data);
if numRows ~= 2000 || numCols ~= 100
    error('数据大小不是2000x100，请检查Excel文件。');
end

% 2. 逐行处理数据
alpha = 0.1;  % 将数据向平均值靠近10%
perturbation_ratio = 0.7;  % 70%的扰动

for i = 1:numRows
    % 3. 添加70%的扰动（基于当前行的数据范围）
    range = max(data(i, :)) - min(data(i, :));  % 当前行数据的范围
    perturbation = perturbation_ratio * range * (2 * rand(1, numCols) - 1);  % 生成扰动
    data(i, :) = data(i, :) + perturbation;  % 加上扰动
    
    % 4. 计算该行的平均值
    row_mean = mean(data(i, :));  
    
    % 5. 将每行数据向平均值靠近
    data(i, :) = data(i, :) + alpha * (row_mean - data(i, :));
end

% 6. 将处理后的数据保存回Excel文件
output_file = 'D:\\jm\\1_processed_with_perturbation_then_closer_to_mean.xlsx';  % 输出文件路径
writematrix(data, output_file);  % 保存处理后的数据到新的Excel文件
