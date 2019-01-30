# -*- coding: utf-8 -*-
"""
Created on Tue Jan 29 11:17:30 2019

@author: BLala
"""

def cash_flow_validity_fx(cash_flows_eff,newest,startDate, lst_fund, bf=0.005):

    import pandas as pd
    import numpy as np
        
#str(dirtoimport_file+newest)
#newest
#newest=str(dirtoimport_file+'UFMPosCash20190128.xls')
#startDate=startDate.replace(day=28)

    cash_xls = pd.read_excel(newest,sheet_name='Cash', 
                             converters={'Settle Date': pd.to_datetime, 'Trade date':pd.to_datetime,
                                                'Portfolio':str, 'Type':str, 
                                                'Security name':str,
                                                'Security Code':str,
                                                'Quantity':float,
                                                ' +/-':str,
                                                'Amount': float},
                            )
    
    cash_xls.columns = [col.strip()  for col in cash_xls.columns]
    
    
    
    cash_xlssub = (cash_xls.copy())[((cash_xls.Type.isin(['CSFLOW','CSHOUT','CSHINJ','CSHWTHD']))&(cash_xls.Portfolio.isin(lst_fund))
                                    &(pd.to_datetime(cash_xls['Settle Date'])==pd.to_datetime(startDate.date()))
                                    &(pd.to_datetime(cash_xls['Settle Date'])==pd.to_datetime(startDate.date())))]
    cash_xlssub.loc[:,'SysFlow']=np.where(cash_xlssub['+/-']=='-',-1.0*cash_xlssub.Amount.values, 
                                   np.where(cash_xlssub['+/-']=='+', cash_xlssub.Amount.values, cash_xlssub.Amount.values))
    
    cash_xlssub=cash_xlssub[['Trade date', 'Portfolio','Type', 'SysFlow']]
    cash_xlssub.columns=['Trade_date','Port_code','Type','SysFlow']
    cash_xlssub_agg=cash_xlssub.groupby(['Port_code','Type']).agg({'SysFlow':'sum'})
           
    cash_flows_eff['Type']=np.where(cash_flows_eff['Inflow']<0,'CSHWTHD', 
                                   np.where(cash_flows_eff['Inflow']>0, 'CSHINJ', ''))
    
    cash_flows_eff_agg=cash_flows_eff.groupby(['Port_code','Type']).agg({'Inflow':'sum'})
    
    csh_tab=pd.concat([cash_xlssub_agg,cash_flows_eff_agg], axis=1)
    csh_tab=csh_tab.reset_index().fillna(0)
    
    
    def cash_flw_flg(man_flow, sys_flow, buffer=bf,x=0):
        if man_flow==0:
            if sys_flow==0:
                return ['No flow',0,0][x]
            elif sys_flow!=0:
                return ['Valid flow, use system flow',sys_flow,0][x]
            else:
                return ['Weird',-999,0][x]
        elif man_flow!=0:
            if sys_flow==0:
                return ['Valid flow, use man flow',man_flow,man_flow][x]
            elif sys_flow!=0:
                if (abs(sys_flow/man_flow -1)<=buffer):
                    return ['Duplicated flow, use system',sys_flow,0][x]
                else:
                    return ['Valid flow, use man flow',man_flow,man_flow][x]
            else:
                return ['Weird',-999,0][x]
        else:
            return ['No flow',0,0,0][x]
                
        
    csh_tab['ValidCashFlag']=csh_tab.apply(lambda r: (cash_flw_flg(r.Inflow,r.SysFlow,0.005,0)),axis=1)
    csh_tab['ActFlow']=csh_tab.apply(lambda r: (cash_flw_flg(r.Inflow,r.SysFlow,0.005,1)),axis=1)       
    csh_tab['Inflow_use']= csh_tab.apply(lambda r: (cash_flw_flg(r.Inflow,r.SysFlow,0.005,2)),axis=1)       
    
    csh_tab_agg=csh_tab.groupby(['Port_code']).agg({'Inflow_use':'sum','ActFlow':'sum'})
    csh_tab_agg=csh_tab_agg.reset_index()
    return[csh_tab,csh_tab_agg]
    
cash_flow_validity_fx()