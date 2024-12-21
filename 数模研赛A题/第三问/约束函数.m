
% 约束函数
function [c, ceq] = constraints(P_ref, P_current, P_t, N_t, T, P_max)
    P_ref = reshape(P_ref, [N_t, T]); % 调整P_ref的形状
    ceq = zeros(T, 1); % 初始化等式约束
    for t = 1:T
        % 总功率平衡约束
        ceq(t) = sum(P_ref(:,t)) - P_t(t); % 每个时间步的功率平衡
    end
    
    % 不等式约束：|P_ref,i(t) - P_current,t(t)| <= 1 和 0 <= P_ref,i(t) <= 5
    c = zeros(N_t * T, 1); % 初始化不等式约束
    count = 1;
    for i = 1:N_t
        for t = 1:T
            % 限制 |P_ref,i(t) - P_current,t(t)| <= 1
            c(count) = abs(P_ref(i,t) - P_current(i,t)) - 1;
            count = count + 1;
        end
    end
end