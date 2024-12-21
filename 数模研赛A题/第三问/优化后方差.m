% 读取Excel文件中的数据
filename = '1_processed_with_perturbation_then_closer_to_mean.xlsx';
data = xlsread(filename, 'Sheet1');  % 默认读取工作表1

% 计算每一行的方差
row_variances = var(data, 0, 2);

% 将结果写入Excel文件
output_filename = '行方差结果.xlsx';
xlswrite(output_filename, row_variances);

% 输出结果到控制台
disp('每行方差已计算并保存到行方差结果.xlsx文件中');
