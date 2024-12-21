% 优化目标函数
function D_total = objective(P_ref, V, omega, N_t, T, eta, alpha, k1, k2, Ct, rho, A, w_shaft, w_thrust, P_max)
    D_total = 0; % 初始化总疲劳损伤
    P_ref = reshape(P_ref, [N_t, T]); % 调整P_ref的形状
    for t = 1:T
        D_shaft_total = 0;
        D_thrust_total = 0;
        for i = 1:N_t
            % 主轴扭矩计算
            if t > 1
                d_term = (omega(i,t) * P_ref(i,t) - omega(i,t-1) * P_ref(i,t-1));
            else
                d_term = 0;
            end
            T_shaft = eta * omega(i,t) * P_ref(i,t) + alpha * d_term;

            % 主轴疲劳损伤
            D_shaft = k1 * T_shaft^2;
            D_shaft_total = D_shaft_total + D_shaft * w_shaft;

            % 塔架推力计算
            F_thrust = Ct * 0.5 * rho * A * V(i,t)^2 * (P_ref(i,t) / P_max);

            % 塔架疲劳损伤
            D_thrust = k2 * F_thrust^2;
            D_thrust_total = D_thrust_total + D_thrust * w_thrust;
        end
        % 每个时间步的总疲劳损伤
        D_total = D_total + D_shaft_total + D_thrust_total;
    end
end
