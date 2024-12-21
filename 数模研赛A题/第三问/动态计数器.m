% 读取Excel数据
filename = '第三问1.xlsx';
data = xlsread(filename, 1, 'A1:CV2000');  % 读取 2000 行 * 100 列的数据

% 动图设置
gif_filename = 'output6.gif';  % GIF文件名
figure('Position', [100, 100, 800, 600]);

% 自定义曲线的颜色和大小
line_color = 'b';  % 设定曲线颜色 ('b'表示蓝色，可换成'r'红色等)
line_width = 2;    % 设定曲线粗细

for t = 1:50  % 遍历前50秒的数据
    % 清空图像
    clf;

    % 绘制当前时间的设备数据
    plot(1:100, data(t, :), '-o', 'Color', [0.05, 0.7, 0.9], 'MarkerSize', 6, 'LineWidth', 2);
    xlabel('风机编号');
    ylabel('功率优化分配');
    title(['风机功率随时间变化 - 第 ', num2str(t+1), ' 秒']);  % 显示时间从0秒开始
    
    % 设置y轴范围为2000000到5000000
    ylim([2000000 5000000]);  
    xlim([1 100]);  % x轴范围固定为1到100个设备

    % 绘制计时器的背景框
    rect_position = [80 2000000 + 5000000 * 0.015 19 5000000 * 0.1];  % [x, y, width, height]
    rectangle('Position', rect_position, 'FaceColor', [0.03, 0.7, 0.3], 'EdgeColor', 'black', 'LineWidth', 2);

    % 在框中显示计时器，计时器显示从0秒开始
    text(81, 2000000 + 5000000 * 0.065, ['计时器: ', num2str(t+1), ' 秒'], 'FontSize', 12, 'Color', [0.95, 0.03, 0.38], 'FontWeight', 'bold');

    % 获取当前图像帧
    frame = getframe(gcf);
    im = frame2im(frame);
    [imind, cm] = rgb2ind(im, 256);

    % 将帧写入GIF文件
    if t == 1
        imwrite(imind, cm, gif_filename, 'gif', 'Loopcount', inf, 'DelayTime', 0.1);  % 第一次写入GIF，设定循环次数为无限
    else
        imwrite(imind, cm, gif_filename, 'gif', 'WriteMode', 'append', 'DelayTime', 0.1);  % 追加写入GIF
    end
    
    % 如果是第50秒，保存当前图片并结束循环
    if t == 49
        saveas(gcf, 'frame_at_50s.png');  % 保存图片
        break;  % 停止循环
    end
end
