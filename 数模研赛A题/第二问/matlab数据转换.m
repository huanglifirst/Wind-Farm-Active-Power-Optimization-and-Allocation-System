WF_1_data = data_TS_WF.WF_1.WT;  % 提取WF_1的100个风机的数据
WF_2_data = data_TS_WF.WF_2.WT;  % 提取WF_2的100个风机的数据
% 提取WF_1的第一台风机的数据作为示例
wind_turbine_1 = WF_1_data(1);

% 提取time、inputs、outputs和states字段
time = wind_turbine_1.time;       % 时间
inputs = wind_turbine_1.inputs;   % 输入，包括功率调度指令和风速
outputs = wind_turbine_1.outputs; % 输出，包括主轴扭矩、塔架推力、实际输出功率
states = wind_turbine_1.status;   % 状态，包括桨距角、低速轴转速、高速轴转速

% 将数据转换为表格格式
turbine_1_table = table(time, inputs(:,1), inputs(:,2), ...
                        outputs(:,1), outputs(:,2), outputs(:,3), ...
                        states(:,1), states(:,2), states(:,3), ...
                        'VariableNames', {'Time', 'Pref', 'WindSpeed', 'Tshaft', 'Ft', 'Pout', 'PitchAngle', 'OmegeR', 'OmegeF'});
此类型的变量不支持使用点进行索引。
 
% 处理WF_1的100台风机的数据
for i = 1:100
    wind_turbine = WF_1_data(i);
    
    % 提取time、inputs、outputs、states字段
    time = wind_turbine.time;
    inputs = wind_turbine.inputs;
    outputs = wind_turbine.outputs;
    states = wind_turbine.status;
    
    % 将数据转换为表格
    turbine_table = table(time, inputs(:,1), inputs(:,2), ...
                          outputs(:,1), outputs(:,2), outputs(:,3), ...
                          states(:,1), states(:,2), states(:,3), ...
                          'VariableNames', {'Time', 'Pref', 'WindSpeed', 'Tshaft', 'Ft', 'Pout', 'PitchAngle', 'OmegeR', 'OmegeF'});
    
    % 保存到Excel文件，给每台风机单独一个工作表
    sheet_name = sprintf('WT_%d', i);  % 工作表名
    writetable(turbine_table, 'WF1_data.xlsx', 'Sheet', sheet_name);
end
此类型的变量不支持使用点进行索引。
 
for i = 1:100
    wind_turbine = WF_1_data{i};  % 使用 {} 访问元胞数组中的结构体
    
    % 提取time、inputs、outputs、states字段
    time = wind_turbine.time;       % 时间数据
    inputs = wind_turbine.inputs;   % 输入数据
    outputs = wind_turbine.outputs; % 输出数据
    states = wind_turbine.status;   % 状态数据
    
    % 将数据转换为表格
    turbine_table = table(time, inputs(:,1), inputs(:,2), ...
                          outputs(:,1), outputs(:,2), outputs(:,3), ...
                          states(:,1), states(:,2), states(:,3), ...
                          'VariableNames', {'Time', 'Pref', 'WindSpeed', 'Tshaft', 'Ft', 'Pout', 'PitchAngle', 'OmegeR', 'OmegeF'});
    
    % 保存到Excel文件，每个风机对应一个工作表
    sheet_name = sprintf('WT_%d', i);  % 工作表名
    writetable(turbine_table, 'WF1_data.xlsx', 'Sheet', sheet_name);
end

% 处理WF_2的100台风机的数据
for i = 1:100
    wind_turbine = WF_2_data{i};  % 使用 {} 访问元胞数组中的结构体
    
    % 提取time、inputs、outputs、states字段
    time = wind_turbine.time;
    inputs = wind_turbine.inputs;
    outputs = wind_turbine.outputs;
    states = wind_turbine.status;
    
    % 将数据转换为表格
    turbine_table = table(time, inputs(:,1), inputs(:,2), ...
                          outputs(:,1), outputs(:,2), outputs(:,3), ...
                          states(:,1), states(:,2), states(:,3), ...
                          'VariableNames', {'Time', 'Pref', 'WindSpeed', 'Tshaft', 'Ft', 'Pout', 'PitchAngle', 'OmegeR', 'OmegeF'});
    
    % 保存到Excel文件，每个风机对应一个工作表
    sheet_name = sprintf('WT_%d', i);  % 工作表名
    writetable(turbine_table, 'WF2_data.xlsx', 'Sheet', sheet_name);
end
无法识别的字段名称 "status"。
 
% 假设已经加载数据为 'data_TS_WF'
WF_1_data = data_TS_WF.WF_1.WT;  % 提取WF_1的100个风机的数据
WF_2_data = data_TS_WF.WF_2.WT;  % 提取WF_2的100个风机的数据

% 处理WF_1的100台风机的数据
for i = 1:100
    wind_turbine = WF_1_data{i};  % 使用 {} 访问元胞数组中的结构体
    
    % 提取time、inputs、outputs、states字段
    time = wind_turbine.time;       % 时间数据
    inputs = wind_turbine.inputs;   % 输入数据
    outputs = wind_turbine.outputs; % 输出数据
    states = wind_turbine.states;   % 改为 states
    
    % 将数据转换为表格
    turbine_table = table(time, inputs(:,1), inputs(:,2), ...
                          outputs(:,1), outputs(:,2), outputs(:,3), ...
                          states(:,1), states(:,2), states(:,3), ...
                          'VariableNames', {'Time', 'Pref', 'WindSpeed', 'Tshaft', 'Ft', 'Pout', 'PitchAngle', 'OmegeR', 'OmegeF'});
    
    % 保存到Excel文件，每个风机对应一个工作表
    sheet_name = sprintf('WT_%d', i);  % 工作表名
    writetable(turbine_table, 'WF1_data.xlsx', 'Sheet', sheet_name);
end

% 处理WF_2的100台风机的数据
for i = 1:100
    wind_turbine = WF_2_data{i};  % 使用 {} 访问元胞数组中的结构体
    
    % 提取time、inputs、outputs、states字段
    time = wind_turbine.time;
    inputs = wind_turbine.inputs;
    outputs = wind_turbine.outputs;
    states = wind_turbine.states;  % 改为 states
    
    % 将数据转换为表格
    turbine_table = table(time, inputs(:,1), inputs(:,2), ...
                          outputs(:,1), outputs(:,2), outputs(:,3), ...
                          states(:,1), states(:,2), states(:,3), ...
                          'VariableNames', {'Time', 'Pref', 'WindSpeed', 'Tshaft', 'Ft', 'Pout', 'PitchAngle', 'OmegeR', 'OmegeF'});
    
    % 保存到Excel文件，每个风机对应一个工作表
    sheet_name = sprintf('WT_%d', i);  % 工作表名
    writetable(turbine_table, 'WF2_data.xlsx', 'Sheet', sheet_name);
end

Exception "java.lang.ClassNotFoundException: com/intellij/openapi/editor/impl/EditorCopyPasteHelperImpl$CopyPasteOptionsTransferableData"while constructing DataFlavor for: application/x-java-serialized-object; class=com.intellij.openapi.editor.impl.EditorCopyPasteHelperImpl$CopyPasteOptionsTransferableData
Exception "java.lang.ClassNotFoundException: com/intellij/openapi/editor/impl/EditorCopyPasteHelperImpl$CopyPasteOptionsTransferableData"while constructing DataFlavor for: application/x-java-serialized-object; class=com.intellij.openapi.editor.impl.EditorCopyPasteHelperImpl$CopyPasteOptionsTransferableData
Exception "java.lang.ClassNotFoundException: com/intellij/codeInsight/editorActions/FoldingData"while constructing DataFlavor for: application/x-java-jvm-local-objectref; class=com.intellij.codeInsight.editorActions.FoldingData
Exception "java.lang.ClassNotFoundException: com/intellij/codeInsight/editorActions/FoldingData"while constructing DataFlavor for: application/x-java-jvm-local-objectref; class=com.intellij.codeInsight.editorActions.FoldingData
