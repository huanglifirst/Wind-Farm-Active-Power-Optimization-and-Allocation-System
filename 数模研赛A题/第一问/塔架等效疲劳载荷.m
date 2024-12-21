% 读取Excel数据
data = xlsread('塔架图.xlsx');

% 假设数据有2行，每行100个数
x = 1:100; % x轴为1到100
y1 = data(1, :); % 第一行数据
y2 = data(2, :); % 第二行数据

% 创建折线图
figure;

% 绘制第二行数据，使用自定义颜色
plot(x, y2, '-.', 'LineWidth', 1.5, 'MarkerSize', 0.5, 'DisplayName', '参考值', 'Color', [0.05, 0.7, 0.9]);
hold on;
% 绘制第一行数据，使用自定义颜色
plot(x, y1, '-', 'LineWidth', 1.5, 'MarkerSize', 0.5, 'DisplayName', '实时值', 'Color',[0.95, 0.30, 0.38]);

% 图形美化
xlabel('风机编号', 'FontSize', 12);
ylabel('等效疲劳载荷', 'FontSize', 12);
legend('show', 'Location', 'best');
grid on;
set(gca, 'FontSize', 14);
ax = gca;
ax.Box = 'on';
ax.XColor = 'k';
ax.YColor = 'k';

% 保存为高质量图片
saveas(gcf, 'line_chart_colored.png');
