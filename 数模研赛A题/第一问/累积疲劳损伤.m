% 读取Excel数据，假设数据为100行4列
data = xlsread('风机90.xlsx');

% 假设数据有100行，每行4个数
x = 1:100; % x轴为1到100
y1 = data(:, 1); % 第一列数据
y2 = data(:, 2); % 第二列数据
y3 = data(:, 3); % 第三列数据
y4 = data(:, 4); % 第四列数据

% 创建图形窗口
figure;

% 第一张图：绘制第一列和第二列数据
subplot(2, 1, 1);
plot(x, y1, '-', 'LineWidth', 1.5, 'DisplayName', '第一列数据', 'Color', [0.05, 0.7, 0.9]);

ylabel('风机累积疲劳损伤', 'FontSize', 12);
grid on;
set(gca, 'FontSize', 14);
ax = gca;
ax.Box = 'on';
ax.XColor = 'k';
ax.YColor = 'k';
hold on;

subplot(2, 1, 2);
plot(x, y3, '-', 'LineWidth', 1.5, 'DisplayName', '第二列数据', 'Color', [0.05, 0.7, 0.9]);
xlabel('时间', 'FontSize', 12);
ylabel('风机扭矩', 'FontSize', 12);
grid on;
set(gca, 'FontSize', 14);
ax = gca;
ax.Box = 'on';
ax.XColor = 'k';
ax.YColor = 'k';

figure;
% 第二张图：绘制第三列和第四列数据
subplot(2, 1, 1);
plot(x, y2, '-', 'LineWidth', 1.5, 'DisplayName', '第三列数据', 'Color', [0.95, 0.30, 0.38]);

ylabel('塔架累积疲劳损伤', 'FontSize', 12);
grid on;
set(gca, 'FontSize', 14);
ax = gca;
ax.Box = 'on';
ax.XColor = 'k';
ax.YColor = 'k';
hold on;

subplot(2, 1, 2);
plot(x, y4, '-', 'LineWidth', 1.5, 'DisplayName', '第四列数据', 'Color', [0.95, 0.30, 0.38]);
xlabel('时间', 'FontSize', 12);
ylabel('塔架推力', 'FontSize', 12);
grid on;
set(gca, 'FontSize', 14);
ax = gca;
ax.Box = 'on';
ax.XColor = 'k';
ax.YColor = 'k';

% 保存为高质量图片
saveas(gcf, 'line_chart_two_plots.png');
