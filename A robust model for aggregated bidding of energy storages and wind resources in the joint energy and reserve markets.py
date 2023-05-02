from __future__ import print_function
from cmath import inf
from docplex.mp.model import Model
from docplex.util.environment import get_environment
import win32com.client as win32
import pandas as pd
import os

### 엑셀 불러오기
excel = win32.Dispatch("Excel.Application")
wb1 = excel.Workbooks.Open(os.getcwd()+"\\robust model_data.xlsx")  
Price_DA = wb1.Sheets("Price_DA")              # Day-ahead prices
Price_RS = wb1.Sheets("Price_RS")              # Reserve price
Price_UR = wb1.Sheets("Price_UR")              # Up regulation prices
Price_DR = wb1.Sheets("Price_DR")              # Down regulation prices
Ramp_rate_BESS = wb1.Sheets("Ramp_rate_BESS")  # Ramp-rate of BES
Expected_P_UR = wb1.Sheets("Expected_P_UR")    # Expected deployed power in up regulation services
Expected_P_DR = wb1.Sheets("Expected_P_DR")    # Expected deployed power in down regulation services
Expected_P_RT_WPR = wb1.Sheets("Expected_P_RT_WPR")  # Expected wind power realization

### 파라미터 설정
time_dim = 24     # 시간 개수 (t)
min_dim = 12      # 5분 x 12 = 1시간 (j)
BESS_dim = 2      # BESS 개수 (s)
WPR_dim = 1       # 풍력발전기 개수 (w)
Marginal_cost_CH = 1    # Marginal cost of BES in charging modes
Marginal_cost_DCH = 1   # Marginal cost of BES in discharging modes
Marginal_cost_WPR = 3   # Marginal cost of WPR
Ramp_rate_WPR = 3       # Ramp-rate of WPR
E_min_BESS = 0    # Minimum energy of BES
E_max_BESS = 30   # Maximum energy of BES

### 최적화 파트
def build_optimization_model(name='Robust_Optimization_Model'):
    mdl = Model(name=name)   # Model - Cplex에 입력할 Model 이름 입력 및 Model 생성
    mdl.parameters.mip.tolerances.mipgap = 0.0001;   # 최적화 계산 오차 설정

    time = [t for t in range(1,time_dim+1)]    # (t)의 one dimension
    time_min = [(t,j) for t in range(1,time_dim + 1) for j in range(1,min_dim+1)]   # (t,j)의 two dimension
    time_n_BESS = [(t,j,s) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1)]   # (t,j,s)의 three dimension
    time_n_WPR = [(t,j,w) for t in range(1,time_dim + 1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1)]     # (t,j,w)의 three dimension

    ### Continous Variable 지정 (연속 변수, 실수 변수)
    P_DA_S = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-S")   # Selling bids in the day-ahead market
    P_DA_B = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-DA-B")   # Buying bids in the day-ahead market
    P_RS = mdl.continuous_var_dict(time, lb=0, ub=inf, name="P-RS")       # Reserve bid 
    
    P_UR = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-UR")   # Deployed power in the up-regulation services
    P_DR = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="P-DR")   # Deployed power in the down-regulation services

    P_DA_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-CH")     # Day-ahead scheduling of BES in charging modes
    P_DA_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DA-DCH")   # Day-ahead scheduling of BES in discharging modes
    P_DA_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DA-WPR")    # Day-ahead scheduling of WPR

    P_UR_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-UR-CH")     # Deployed up regulation power of BES in charging mode
    P_UR_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-UR-DCH")   # Deployed up regulation power of BES in discharging mode
    P_UR_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-UR-WPR")    # Deployed up regulation power of WPR   

    P_DR_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DR-CH")      # Deployed down regulation power of BES in charging mode
    P_DR_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-DR-DCH")    # Deployed down regulation power of BES in discharging mode
    P_DR_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-DR-WPR")     # Deployed down regulation power of WPR  

    P_RS_CH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-CH")      # Reserve scheduling of BES in charging modes
    P_RS_DCH = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=inf, name="P-RS-DCH")    # Reserve scheduling of BES in discharging modes
    P_RS_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RS-WPR")     # Reserve scheduling of WPR

    # P_SP_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-SP-WPR")            # Spilled power of WPR (difference between the realization of wind power and the scheduled power of WPR)
    Energy_BESS = mdl.continuous_var_dict(time_n_BESS, lb=0, ub=E_max_BESS, name="Energy-BESS")  # Energy level of BES 
    P_RT_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="P-RT-WPR")                # Realization of wind power in real-time

    # AV_WPR = mdl.continuous_var_dict(time_n_WPR, lb=0, ub=inf, name="AV-WPR")                    # Auxiliary variables for linearization
    AV_RO = mdl.continuous_var_dict(time_min, lb=0, ub=inf, name="AV-RO")                          # Auxiliary variable of RO
    
    # Auxiliary variable of RO
    B_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="B-t")          # Income of owner
    C_t = mdl.continuous_var_dict(time, lb=0, ub=inf, name="C-t")          # Cost of owner

    ### Binary Variable 지정 (이진 변수)
    D_Char = mdl.binary_var_dict(time_n_BESS, name="D-Char")      # Charging binary variables of BES (알파)
    D_Dchar = mdl.binary_var_dict(time_n_BESS, name="D-DChar")    # Discharging binary variables of BES (베타)
    D_WPR = mdl.binary_var_dict(time_n_WPR, name="D-WPR")         # Commitment status binary variable of WPR
    
    ### Objective function - 식(1) / 식(65)
    # mdl.maximize(mdl.sum(Price_DA.Cells(t+1,2).Value * P_DA_S[t] - Price_DA.Cells(t+1,2).Value * P_DA_B[t] + Price_RS.Cells(t+1,2).Value * P_RS[t]  
    #                       - Marginal_cost_DCH * P_DA_DCH[(t,j,s)] - Marginal_cost_CH * P_DA_CH[(t,j,s)] - Marginal_cost_WPR * P_DA_WPR[(t,j,w)] + AV_RO[(t,j)]
    #                       for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)))
    
    mdl.maximize(mdl.sum(Price_DA.Cells(t+1,2).Value * (P_DA_DCH[(t,j,s)] - P_DA_CH[(t,j,s)] + P_DA_WPR[(t,j,w)]) + Price_RS.Cells(t+1,2).Value * (P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] + P_RS_WPR[(t,j,w)])
                         - Marginal_cost_DCH * P_DA_DCH[(t,j,s)] - Marginal_cost_CH * P_DA_CH[(t,j,s)] - Marginal_cost_WPR * P_DA_WPR[(t,j,w)]
                         + AV_RO[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1))) 

    # Robust Optizimation을 위한 변수 (BESS + WPR) - 식(65)
    mdl.add_constraints(AV_RO[(t,j)] <= mdl.sum(Price_UR.Cells(t+1,j+1).Value * (P_UR_DCH[(t,j,s)] + P_UR_CH[(t,j,s)] + P_UR_WPR[(t,j,w)]) 
                                                 + Price_DR.Cells(t+1,j+1).Value * (P_DR_DCH[(t,j,s)] + P_DR_CH[(t,j,s)] + P_DR_WPR[(t,j,w)])
                                                 - Marginal_cost_DCH * P_UR_DCH[(t,j,s)] - Marginal_cost_CH * P_DR_CH[(t,j,s)] - Marginal_cost_WPR * P_UR_WPR[(t,j,w)] 
                                                 for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1))
    
    ### B_t - 식(2)
    mdl.add_constraints(B_t[t] == mdl.sum((Price_DA.Cells(t+1,2).Value * P_DA_S[t] - Price_DA.Cells(t+1,2).Value * P_DA_B[t] + Price_RS.Cells(t+1,2).Value * P_RS[t])
                                          + (Price_UR.Cells(t+1,j+1).Value * P_UR[(t,j)] + Price_DR.Cells(t+1,j+1).Value * P_DR[(t,j)]) for j in range(1,min_dim+1)) for t in range(1,time_dim+1))  # Income of owner

     ### C_t - 식(3)
    mdl.add_constraints(C_t[t] == mdl.sum(Marginal_cost_DCH * (P_DA_DCH[(t,j,s)] + P_UR_DCH[(t,j,s)] - P_DR_DCH[(t,j,s)]) + Marginal_cost_CH * (P_DA_CH[(t,j,s)] + P_DR_CH[(t,j,s)] - P_UR_CH[(t,j,s)]) 
                                          + Marginal_cost_WPR * (P_DA_WPR[(t,j,w)] + P_UR_WPR[(t,j,w)] - P_DR_WPR[(t,j,w)]) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)) for t in range(1,time_dim+1))  # Cost of owner
    
    ### Equality constraints - 식(4) ~ 식(6) + 식(12) ~ 식(14)
    mdl.add_constraints(P_DA_DCH[(t,j,s)] == P_DA_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(4)

    mdl.add_constraints(P_DA_CH[(t,j,s)] == P_DA_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(5)

    mdl.add_constraints(P_DA_WPR[(t,j,w)] == P_DA_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(6

    mdl.add_constraints(P_RS_CH[(t,j,s)] == P_RS_CH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(12)   

    mdl.add_constraints(P_RS_DCH[(t,j,s)] == P_RS_DCH[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(13)                                           

    mdl.add_constraints(P_RS_WPR[(t,j,w)] == P_RS_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(14)

    ### Constraints of day-ahead energy / reserve bids / real-time deployed power in the up and down regulation services - 식(7) ~ 식(11)          
    mdl.add_constraints(P_DA_S[t] == mdl.sum(P_DA_DCH[(t,j,s)] + P_DA_WPR[(t,j,w)] for j in range(1,min_dim+1) for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)) for t in range(1,time_dim+1))  # 식(7)

    mdl.add_constraints(P_DA_B[t] == mdl.sum(P_DA_CH[(t,j,s)] for j in range(1,min_dim+1) for s in range(1,BESS_dim+1)) for t in range(1,time_dim+1))  # 식(8)
    
    mdl.add_constraints(P_RS[t] == mdl.sum(P_RS_CH[(t,j,s)] + P_RS_DCH[(t,j,s)] + P_RS_WPR[(t,j,w)] for j in range(1,min_dim+1) for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)) for t in range(1,time_dim+1))  # 식(9)
    
    mdl.add_constraints(P_UR[(t,j)] == mdl.sum(P_UR_DCH[(t,j,s)] + P_UR_CH[(t,j,s)] + P_UR_WPR[(t,j,w)] for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(10)

    mdl.add_constraints(P_DR[(t,j)] == mdl.sum(P_DR_DCH[(t,j,s)] + P_DR_CH[(t,j,s)] + P_DR_WPR[(t,j,w)] for s in range(1,BESS_dim+1) for w in range(1,WPR_dim+1)) for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(11)

    ### 식(15) ~ 식(16)
    mdl.add_constraints(P_UR[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(15)

    mdl.add_constraints(P_DR[(t,j)] <= P_RS[t] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식(16)
    
    ### Constarints of stored energy of BES - 식(17) ~ 식(19)
    for t in range(1,time_dim+1):
        for j in range(1,min_dim+1):
            for s in range(1,BESS_dim+1):
                if j == 1:
                    if t == 1:
                        mdl.add_constraint(Energy_BESS[(t,j,s)] == 15)
                        
                    else:
                        mdl.add_constraint(Energy_BESS[(t,j,s)] == Energy_BESS[(t-1,12,s)])
                
                elif t == 24 and j == 12:
                    mdl.add_constraint(Energy_BESS[(t,j,s)] == 15)
                
                else:
                    mdl.add_constraint(Energy_BESS[(t,j,s)] == Energy_BESS[(t,j-1,s)] + P_DA_CH[(t,j,s)] - P_DA_DCH[(t,j,s)] + P_DR_CH[(t,j,s)] - P_UR_DCH[(t,j,s)])    
                        
    ### Constarints of capacity - 식(20) ~ 식(38)
    # Power capacity of BES in day-ahead planning - 식(20) ~ 식(25) + 식(30) ~ 식(31)
    for t in range(1,time_dim+1):
        for j in range(1,min_dim+1):
            for s in range(1,BESS_dim+1):
                if s == 1:
                    mdl.add_constraint(P_DA_CH[(t,j,s)] <= 5 * D_Char[(t,j,s)])                         # 식(20)
                    mdl.add_constraint(P_DA_DCH[(t,j,s)] <= 5 * D_Dchar[(t,j,s)])                       # 식(24)
                    mdl.add_constraint(P_RS_CH[(t,j,s)] <= 5 * D_Char[(t,j,s)] - P_DA_CH[(t,j,s)])     # 식(21)
                    mdl.add_constraint(P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)] <= 5 * D_Char[(t,j,s)])     # 식(22)
                    mdl.add_constraint(P_RS_DCH[(t,j,s)] <= 5 * D_Dchar[(t,j,s)] - P_DA_DCH[(t,j,s)])  # 식(25)
                    mdl.add_constraint(P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)] <= 5 * D_Dchar[(t,j,s)])  # 식(30)
                
                else:
                    mdl.add_constraint(P_DA_CH[(t,j,s)] <= 3 * D_Char[(t,j,s)])                         # 식(20)
                    mdl.add_constraint(P_DA_DCH[(t,j,s)] <= 3 * D_Dchar[(t,j,s)])                       # 식(24)
                    mdl.add_constraint(P_RS_CH[(t,j,s)] <= 3 * D_Char[(t,j,s)] - P_DA_CH[(t,j,s)])     # 식(21)
                    mdl.add_constraint(P_DA_CH[(t,j,s)] + P_RS_CH[(t,j,s)] <= 3 * D_Char[(t,j,s)])     # 식(22)
                    mdl.add_constraint(P_RS_DCH[(t,j,s)] <= 3 * D_Dchar[(t,j,s)] - P_DA_DCH[(t,j,s)])  # 식(25)
                    mdl.add_constraint(P_DA_DCH[(t,j,s)] + P_RS_DCH[(t,j,s)] <= 3 * D_Dchar[(t,j,s)])  # 식(30)
                
                mdl.add_constraint(0 * D_Char[(t,j,s)] <= P_DA_CH[(t,j,s)])                             # 식(20)
                mdl.add_constraint(0 * D_Dchar[(t,j,s)] <= P_DA_DCH[(t,j,s)])                           # 식(24)
                mdl.add_constraint(0 <= P_RS_CH[(t,j,s)])                                              # 식(21)
                mdl.add_constraint(0 * D_Char[(t,j,s)] <= P_DA_CH[(t,j,s)] - P_RS_CH[(t,j,s)])         # 식(23)
                mdl.add_constraint(0 <= P_RS_DCH[(t,j,s)])                                             # 식(25)
                mdl.add_constraint(0 * D_Dchar[(t,j,s)] <= P_DA_DCH[(t,j,s)] - P_RS_DCH[(t,j,s)])      # 식(31)

    # Deployed power of BES in regulation service - 식(26) ~ 식(31)
    mdl.add_constraints(0 <= P_UR_CH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(26)

    mdl.add_constraints(P_UR_CH[(t,j,s)] <= P_RS_CH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(26)

    mdl.add_constraints(0 <= P_DR_CH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(27)

    mdl.add_constraints(P_DR_CH[(t,j,s)] <= P_RS_CH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(27)

    mdl.add_constraints(0 <= P_UR_DCH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(28)

    mdl.add_constraints(P_UR_DCH[(t,j,s)] <= P_RS_DCH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(28)

    mdl.add_constraints(0 <= P_DR_DCH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(29)

    mdl.add_constraints(P_DR_DCH[(t,j,s)] <= P_RS_DCH[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))   # 식(29)

    # Energy capacity in the real-time - 식(32)
    mdl.add_constraints(E_min_BESS * (D_Char[(t,j,s)] + D_Dchar[(t,j,s)]) <= Energy_BESS[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(32)

    mdl.add_constraints(Energy_BESS[(t,j,s)] <= E_max_BESS * (D_Char[(t,j,s)] + D_Dchar[(t,j,s)]) for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(32)
    
    # Capacity of WPR in the dayahead planning - 식(33) ~ 식(36)
    # mdl.add_constraints(0 <= P_DA_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(33)
    
    # mdl.add_constraints(P_DA_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] + (1 - D_WPR[(t,j,w)]) * 10000000000 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 
    
    # mdl.add_constraints(P_DA_WPR[(t,j,w)] <= 10000000000 * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 

    mdl.add_constraints(P_RT_WPR[(t,j,w)] - (1 - D_WPR[(t,j,w)]) * 10000000000 <= P_DA_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(33)
    
    mdl.add_constraints(P_DA_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] + (1 - D_WPR[(t,j,w)]) * 10000000000 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 
    
    mdl.add_constraints(P_DA_WPR[(t,j,w)] <= 10000000000 * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 

    mdl.add_constraints(-1 * 10000000000 * D_WPR[(t,j,w)] <= P_DA_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(33) 

    # mdl.add_constraints(0 <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(34)

    # mdl.add_constraints(P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] - P_DA_WPR[(t,j,w)] + (1 - D_WPR[(t,j,w)]) * 10000000000 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(34)

    # mdl.add_constraints(P_RS_WPR[(t,j,w)] <= 10000000000 * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(34)

    mdl.add_constraints(P_RT_WPR[(t,j,w)] - P_DA_WPR[(t,j,w)] - (1 - D_WPR[(t,j,w)]) * 10000000000 <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(34)

    mdl.add_constraints(P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] - P_DA_WPR[(t,j,w)] + (1 - D_WPR[(t,j,w)]) * 10000000000 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(34)

    mdl.add_constraints(P_RS_WPR[(t,j,w)] <= 10000000000 * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(34)

    mdl.add_constraints(-1 * 10000000000 * D_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식(34)
    
    mdl.add_constraints(P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] + (1 - D_WPR[(t,j,w)]) * 10000000000 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(35)
    
    mdl.add_constraints(P_DA_WPR[(t,j,w)] + P_RS_WPR[(t,j,w)] <= 10000000000 * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(35)    

    mdl.add_constraints(0 <= P_DA_WPR[(t,j,w)] - P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(36)

    # Deployed power of WPR in the regulation service - 식(37) ~ 식(38)
    mdl.add_constraints(0 <= P_UR_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(37)

    mdl.add_constraints(P_UR_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(37)

    mdl.add_constraints(0 <= P_DR_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(38)

    mdl.add_constraints(P_DR_WPR[(t,j,w)] <= P_RS_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(38)   

    ### Constarints of binary decision Variables - 식(39) ~ 식(42)
    # Commitment status of WPRs, and BESs in the charging and discharging modes in the dayahead planning
    mdl.add_constraints(D_WPR[(t,j,w)] == D_WPR[(t,J,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for w in range(1,WPR_dim+1))       # 식(39)

    mdl.add_constraints(D_Char[(t,j,s)] == D_Char[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))    # 식(40)

    mdl.add_constraints(D_Dchar[(t,j,s)] == D_Dchar[(t,J,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for J in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(41)

    mdl.add_constraints(0 <= D_Char[(t,j,s)] + D_Dchar[(t,j,s)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(42)

    mdl.add_constraints(D_Char[(t,j,s)] + D_Dchar[(t,j,s)] <= 1 for t in range(1,time_dim+1) for j in range(1,min_dim+1) for s in range(1,BESS_dim+1))  # 식(42)

    ### Constarints of ramp-rate - 식(43) ~ 식(57)
    for t in range(1,time_dim+1):
        for j in range(1,min_dim+1):
            for s in range(1,BESS_dim+1):
                if j == 1:
                    if t == 1:
                        mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_CH[(t,j,s)])
                        mdl.add_constraint(P_DA_CH[(t,j,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_DCH[(t,j,s)])
                        mdl.add_constraint(P_DA_DCH[(t,j,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(P_RS_CH[(t,j,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(P_RS_DCH[(t,j,s)] <= Ramp_rate_BESS.Cells(s,3).Value)    
                        
                    else:
                        mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,12,s)])
                        mdl.add_constraint(P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,12,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,12,s)])
                        mdl.add_constraint(P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,12,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,12,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,12,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,12,s)] + P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,12,s)])
                        mdl.add_constraint(P_DA_CH[(t,j,s)] - P_DA_CH[(t-1,12,s)] + P_RS_CH[(t,j,s)] + P_RS_CH[(t-1,12,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                        mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,12,s)] + P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,12,s)])
                        mdl.add_constraint(P_DA_DCH[(t,j,s)] - P_DA_DCH[(t-1,12,s)] + P_RS_DCH[(t,j,s)] + P_RS_DCH[(t-1,12,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                
                else:
                    mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)])   
                    mdl.add_constraint(P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] <= Ramp_rate_BESS.Cells(s,3).Value)  
                    mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)])    
                    mdl.add_constraint(P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                    mdl.add_constraint(P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                    mdl.add_constraint(P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
                    mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] + P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)])
                    mdl.add_constraint(P_DA_CH[(t,j,s)] - P_DA_CH[(t,j-1,s)] + P_RS_CH[(t,j,s)] + P_RS_CH[(t,j-1,s)] <= Ramp_rate_BESS.Cells(s,3).Value)         
                    mdl.add_constraint(-1 * Ramp_rate_BESS.Cells(s,3).Value <= P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] + P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)])
                    mdl.add_constraint(P_DA_DCH[(t,j,s)] - P_DA_DCH[(t,j-1,s)] + P_RS_DCH[(t,j,s)] + P_RS_DCH[(t,j-1,s)] <= Ramp_rate_BESS.Cells(s,3).Value)
    
    for t in range(1,time_dim+1):
        for j in range(1,min_dim+1):
            for w in range(1,WPR_dim+1):    
                if j == 1:
                    if t == 1: 
                        mdl.add_constraint(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)])
                        mdl.add_constraint(P_DA_WPR[(t,j,w)] <= Ramp_rate_WPR)
                        mdl.add_constraint(P_RS_WPR[(t,j,w)] <= Ramp_rate_WPR)                    
                    else:
                        mdl.add_constraint(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,12,w)])
                        mdl.add_constraint(P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,12,w)] <= Ramp_rate_WPR)
                        mdl.add_constraint(P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,12,w)] <= Ramp_rate_WPR)
                        mdl.add_constraint(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,12,w)] + P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,12,w)])
                        mdl.add_constraint(P_DA_WPR[(t,j,w)] - P_DA_WPR[(t-1,12,w)] + P_RS_WPR[(t,j,w)] + P_RS_WPR[(t-1,12,w)] <= Ramp_rate_WPR) 
                else:
                    mdl.add_constraint(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)])
                    mdl.add_constraint(P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] <= Ramp_rate_WPR)
                    mdl.add_constraint(P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)] <= Ramp_rate_WPR)
                    mdl.add_constraint(-1 * Ramp_rate_WPR <= P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] + P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)])
                    mdl.add_constraint(P_DA_WPR[(t,j,w)] - P_DA_WPR[(t,j-1,w)] + P_RS_WPR[(t,j,w)] + P_RS_WPR[(t,j-1,w)] <= Ramp_rate_WPR)
                      
    ### Constarints of spillage power - 식(58) ~ 식(59)
    # mdl.add_constraints(P_SP_WPR[(t,j,w)] == P_RT_WPR[(t,j,w)] - (P_DA_WPR[(t,j,w)] + P_UR_WPR[(t,j,w)] - P_DR_WPR[(t,j,w)]) 
    #                     for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))                  # 식(58)

    # mdl.add_constraints(0 <= P_SP_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59) - (A.1)

    #mdl.add_constraints(P_SP_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59)

    # mdl.add_constraints(P_SP_WPR[(t,j,w)] <= AV_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59) - (A.1)
                        
    # mdl.add_constraints(0 <= AV_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59) - (A.2)

    # mdl.add_constraints(AV_WPR[(t,j,w)] <= 10000000000 * D_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59) - (A.3) 
                        
    # mdl.add_constraints(AV_WPR[(t,j,w)] <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59) - (A.4) 

    # mdl.add_constraints(P_RT_WPR[(t,j,w)] - 10000000000 * (1 - D_WPR[(t,j,w)]) <= AV_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))   # 식(59) - (A.5) 

    ### Constraints of uncertain parameters-  식(61) ~ 식(63)
    mdl.add_constraints(0.9 * Expected_P_UR.Cells(t+1,j+1).Value <= P_UR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (61) / 변동구간 +-10%

    mdl.add_constraints(P_UR[(t,j)] <= 1.1 * Expected_P_UR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1))   # 식 (61) / 변동구간 +-10%

    mdl.add_constraints(0.9 * Expected_P_DR.Cells(t+1,j+1).Value <= P_DR[(t,j)] for t in range(1,time_dim+1) for j in range(1,min_dim+1))  # 식 (62) / 변동구간 +-10%

    mdl.add_constraints(P_DR[(t,j)] <= 1.1 * Expected_P_DR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1))   # 식 (62) / 변동구간 +-10%

    mdl.add_constraints(0.9 * Expected_P_RT_WPR.Cells(t+1,j+1).Value <= P_RT_WPR[(t,j,w)] for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식 (63) / 변동구간 +-10%

    mdl.add_constraints(P_RT_WPR[(t,j,w)] <= 1.1 * Expected_P_RT_WPR.Cells(t+1,j+1).Value for t in range(1,time_dim+1) for j in range(1,min_dim+1) for w in range(1,WPR_dim+1))  # 식 (63) / 변동구간 +-10%

    return mdl

### Write Result
def result_optimization_model(model, DataFrame):
    mdl = model
    frame = DataFrame
    
    wb_result = excel.Workbooks.Open(os.getcwd()+"\\robust model_result.xlsx")
    ws1 = wb_result.Worksheets("Optimization Result")
    
    ### Sheet 1  
    # Total Revenue
    ws1.Cells(1,2).Value = "Optimization Result"
    ws1.Cells(2,1).Value = "Total Revenue [$]"
    ws1.Cells(2,2).Value = float(mdl.objective_value)
    
    # B_t
    ws1.Cells(4,1).Value = "Income of owner [$]"
    ws1.Cells(4,2).Value = frame.loc[frame['var']=="B-t"]['index2'].sum()
    
    # C_t
    ws1.Cells(5,1).Value = "Cost of owner [$]"
    ws1.Cells(5,2).Value = frame.loc[frame['var']=="C-t"]['index2'].sum()    
     
    # AV_RO
    ws1.Cells(6,1).Value = "AV_RO"
    ws1.Cells(6,2).Value = frame.loc[frame['var']=="AV-RO"]['index3'].sum()
          
    print("Optimization Result Calculation Done!")   
    
    wb_result.Save()
    excel.Quit()
    
### Main Program    
if __name__ == '__main__':
    mdl = build_optimization_model() # 최적화 모델 생성
    mdl.print_information() # 모델로부터 나온 정보를 출력
    s = mdl.solve(log_output=True) # 모델 풀기
    
    if s: # 해가 존재하는 경우                
        obj = mdl.objective_value
        mdl.get_solve_details()
        print("* Total cost=%g" % obj)
        print("*Gap tolerance = ", mdl.parameters.mip.tolerances.mipgap.get())
        
        data = [v.name.split('_') + [s.get_value(v)] for v in mdl.iter_variables()] # 변수 데이터 저장
        frame = pd.DataFrame(data, columns=['var', 'index1', 'index2', 'index3', 'value']) # 변수 중 시간 성분만 있는 경우 'index2'에 값이 저장됨
        frame.to_excel(os.getcwd()+"\\variable_result.xlsx")
        
        result_optimization_model(mdl, frame)  # 결과 출력부        
        
        # Save the CPLEX solution as "solution.json" program output
        with get_environment().get_output_stream("solution.json") as fp: #json 형태로 solution 저장
            mdl.solution.export(fp, "json")
        
    else: # 해가 존재하지 않는 경우
        print("* model has no solution")
        excel.Quit() 
    