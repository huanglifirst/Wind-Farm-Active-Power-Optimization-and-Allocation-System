% 1. 从Excel文件中读取数据
file_path = 'D:\\jm\\第三问1.xlsx';  % Excel文件路径
data = readmatrix(file_path);  % 读取整个Excel文件中的数据

% 确保数据是2000x100的矩阵
[numRows, numCols] = size(data);
if numRows ~= 2000 || numCols ~= 100
    error('数据大小不是2000x100，请检查Excel文件。');
end

% 2. 逐行处理数据
alpha = 0.2;  % 将数据向平均值靠近10%
for i = 1:numRows
    row_mean = mean(data(i, :));  % 计算该行的平均值
    % 3. 将每行数据向平均值靠近
    data(i, :) = data(i, :) + alpha * (row_mean - data(i, :));
end

% 4. 将处理后的数据保存回Excel文件
output_file = 'D:\\jm\\1_processed.xlsx';  % 输出文件路径
writematrix(data, output_file);  % 保存处理后的数据到新的Excel文件
