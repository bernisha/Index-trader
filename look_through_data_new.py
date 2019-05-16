# -*- coding: utf-8 -*-
"""
Created on Thu Mar 14 11:37:15 2019

@author: BLala
"""

# Map Benchmarks


def look_through_fx(dfprt1,n_comb,fund_xls,dic_om_index,cash_flows_eff, lst_fund):

    import pandas as pd
    import numpy as np
    from write_excel import assetClassB as assetClass  
    from tkinter import filedialog
    from tkinter import Tk

    dfprt1.loc[:,'BmkCode']=dfprt1.Port_code.map(lambda x:dic_om_index[x][3])
    #dfprt1.loc[:,'TypeFund']=dfprt1.Port_code.map(lambda x:dic_om_index[x][2])
    
    dfprt_eq=dfprt1[dfprt1.Sec_type.isin(['EQ : ORDINARY', 'EQ : RIGHTS',  'FUTURE : EQUITY', 'FUND : LOC EQ' ,"EQ : PROPERTY","EQ : PREFERED","FUND : LOC EQ S",
                                          'EQ : ORDINARY SHARE','EQ : STANDARD RIGHTS ISSUE','FUTURE : EQUITY','FUND : LOCAL EQUITY','EQ : PROPERTY','EQ : PREFERED SHARE','FUND : LOCAL EQUITY SMALL CAP'])]
    
    # Add Futures and cash inforamtion
    
    fut_dtaX=n_comb[['Port_code','tot_fut','Close_price','Future_Code_y']]
    fut_dta = fut_dtaX.copy()
    fut_dta.loc[:,'EffExposure']=fut_dta[['tot_fut']].values*np.nan_to_num(fut_dta[['Close_price']].values)*10
    fut_dta.loc[:,'MarketValue']=0
    fut_dta.loc[:,'Quantity']= fut_dta['tot_fut']
   # fut_dta.loc[:,'Sec_type']='FUTRE STCK INDX'
    fut_dta.loc[:,'Sec_type']='FUTURE : EQUITY INDEX'
    fut_dta.loc[:,'Sec_code']=fut_dta['Future_Code_y'].values
    fut_dta.loc[:,'Sec_name']=fut_dta['Future_Code_y'].values
    
    fut_dta.loc[:,'Trade_date']=dfprt1['Trade_date'].head(len(fut_dta)).values
    
    fut_dta=fut_dta[['Close_price', 'EffExposure','MarketValue','Port_code','Quantity','Sec_code', 'Sec_name','Sec_type','Trade_date']]
    fut_dta=fut_dta[~(fut_dta.Sec_code=='NoFuture')]
    
    del fut_dtaX
    
    cash_dtaX=n_comb[['Port_code','cash_bpm']]
    cash_dta=cash_dtaX.copy()
    cash_dta.loc[:,'Close_price'] = 1
    cash_dta.loc[:,'EffExposure']= cash_dta['cash_bpm']
    cash_dta.loc[:,'MarketValue'] = cash_dta['cash_bpm']
    cash_dta.loc[:,'Quantity'] = cash_dta['cash_bpm']
    cash_dta.loc[:,'Sec_type'] = 'CURRENCY'
    cash_dta.loc[:,'Sec_code'] = 'ZAR'
    cash_dta.loc[:,'Trade_date']=dfprt1['Trade_date'].head(len(cash_dta)).values
    cash_dta.loc[:,'Sec_name'] = 'VAL'
    
    cash_dta=cash_dta[['Close_price', 'EffExposure','MarketValue','Port_code','Quantity','Sec_code', 'Sec_name','Sec_type','Trade_date']]
    
    
    del cash_dtaX
    
    tot_dta = fut_dta.append(cash_dta, sort=False)
    tot_dta.loc[:,'TypeFund']=tot_dta.Port_code.map(lambda x:dic_om_index[x][2])
    
    
    tot_dta['AssetType1']=tot_dta.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund, cash_flows_eff)).split(",")[0],axis=1)
    tot_dta['AssetType3']=tot_dta.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[2],axis=1)
    tot_dta['AssetType5']=tot_dta.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[4],axis=1)
    tot_dta.loc[:,'BmkCode']=tot_dta.Port_code.map(lambda x:dic_om_index[x][3])
          
    tot_dta= tot_dta[['AssetType1', 'AssetType3', 'AssetType5', 'Close_price', 'EffExposure',
           'MarketValue', 'Port_code', 'Quantity', 'Sec_code', 'Sec_type',
           'Trade_date', 'BmkCode']]
    
    # Append to equity holdings, cash and futures
    
    dfprt_all = dfprt_eq.append(tot_dta, sort=False)
    
    # Pul in exceptions
    
    # Pull in n_comb (data - cash_bpm, tot_fut)
    
    lst_bmks=dfprt1['BmkCode'].unique().tolist()
    
    bmk_dta=fund_xls[fund_xls.Port_code.isin(lst_bmks)][['Port_code','Sec_code','Close_price','Quantity','Market_price']]
    bmk_dta.columns = ['BmkCode','Sec_code','Bmk_price','Bmk_quantity','BmkValue']
    
    all_bmk_data = pd.DataFrame()
    for key, val in dic_om_index.items():
        if key in lst_fund:
            print(key)
            print(val[3])
            df=bmk_dta[bmk_dta.BmkCode==val[3]]
            if df.empty:
                print("No Data")
            else:
                df1=df.copy()
                df1.loc[:,'Port_code']=key
                df1.loc[:,'bmk_wgt']=df['BmkValue']/np.nansum(df['BmkValue'])
                all_bmk_data= all_bmk_data.append([df1])
                del df1 
                del df
    # Composites Data
    
    mapp_comp=pd.read_csv('C:\\IndexTrader\\required_inputs\\comp_mappings.csv')
    mapp_comp['InstName'] = np.where(mapp_comp.Type.isin(['FUTURE : EQUITY','FUTRE STCK INDX','FUTURE : EQUITY INDEX']), 
                                     mapp_comp.InstName+cash_flows_eff.fut_sufx.unique()[0], mapp_comp.InstName)
    dic_om_composites=mapp_comp.set_index(['InstName']).T.to_dict('list')
    
    all_dta=pd.DataFrame()
    
    for key,val in dic_om_composites.items():
        key_1= key
        val_1 = val[1]
        val_2 = val[0]
        print(val_1)
        if val_2 in ['FUND : LOC EQ','FUTRE STCK INDX','FUTURE : EQUITY INDEX']:
            comp_dta=fund_xls[fund_xls.Port_code == val_1][['Port_code','Sec_code','Close_price','Quantity','Market_price']]
            comp_dta['Comp_Code']= key_1
            comp_dta['Comp_wght']=comp_dta['Market_price']/comp_dta['Market_price'].sum()
        else:
            comp_dta=fund_xls[fund_xls.Sec_code == key_1][['Port_code','Sec_code','Close_price','Quantity','Market_price']]
            comp_dta['Comp_Code']= key_1
            comp_dta['Port_code']=val_1
            comp_dta['Close_price'] = (fund_xls[fund_xls.Sec_code==val_1]['Close_price']).unique()[0]
            comp_dta=comp_dta.drop_duplicates(['Port_code','Sec_code'])
            comp_dta['Comp_wght']=1
        
        all_dta = all_dta.append([comp_dta])
        print(key_1)
    
    
    all_dta.columns = ['BComCode','BSec_code','Comp_price','comp_quantity','CompValue','Comp_Code','Comp_wgt']
    
    dfprt_comp = pd.merge(dfprt_all,all_dta,left_on=['Sec_code'], right_on=['Comp_Code'], how = 'outer')
    dfprt_comp  = dfprt_comp[~dfprt_comp.Sec_code.isnull()]
    
    dfprt_comp.loc[:,'Cmb_code'] =   np.where(dfprt_comp.Sec_type.isin(['FUTURE : EQUITY','EQ : RIGHTS','EQ : STANDARD RIGHTS ISSUE']), dfprt_comp.BComCode,
                                        np.where(dfprt_comp.BSec_code.isnull(), dfprt_comp.Sec_code.values, 
                                                 dfprt_comp.BSec_code.values))
    
    dfprt_comp.loc[:,'Cmb_effexp']=np.where(dfprt_comp.Comp_wgt.isnull(), dfprt_comp.EffExposure.values, dfprt_comp.Comp_wgt.values*dfprt_comp.EffExposure.values)
    dfprt_comp.loc[:,'Cmb_effexp']=np.where(dfprt_comp.Sec_type.isin(['EQ : RIGHTS','EQ : STANDARD RIGHTS ISSUE']), dfprt_comp.Quantity.values*dfprt_comp.Comp_price.values, dfprt_comp.Cmb_effexp.values)
    
    #dfprt_comp[dfprt_comp.Sec_code.isin(['RBP','RBPN','ALSIJ19','DSWIXCCH','DRIEQCCH','OMUSJ19','NXDSJ19'])].to_csv('c:\\data\\right_x.csv')
    
    dfprt_comp_agg=dfprt_comp.groupby(['Trade_date','Port_code','BmkCode','Cmb_code']).agg({'Cmb_effexp':'sum'})
    dfprt_comp_fund=dfprt_comp.groupby(['Trade_date','Port_code','BmkCode']).agg({'Cmb_effexp':'sum'})
    dfprt_comp_fund.columns = ['Fund_EffExp'] 
    #dfprt_comp_fund.reindex(dfprt_comp_agg.index, level=1).ffill()
    
    
    dfprt_comp_fund=dfprt_comp_fund.reset_index()
    dfprt_comp_agg=dfprt_comp_agg.reset_index()
    #dfprt_comp_fund.set_index(['Trade_date','Port_code','BmkCode','Cmb_code'],inplace = True) 
    
    dfprt_comp_agg=dfprt_comp_agg.merge(dfprt_comp_fund, on =['Trade_date','Port_code','BmkCode'], how = 'left')
    dfprt_comp_agg.loc[:,'Cmp_wgt'] = dfprt_comp_agg.Cmb_effexp.values/dfprt_comp_agg.Fund_EffExp.values
    dfprt_comp_agg=dfprt_comp_agg.groupby(['Trade_date','Port_code','BmkCode','Cmb_code']).agg({'Cmb_effexp':'sum','Cmp_wgt':'sum','Fund_EffExp':'max'})
    
    
    dfprt_comp_agg_R = dfprt_comp_agg.reset_index()
    dfprt_comp_agg_R.columns = ['Trade_date', 'Port_code', 'BmkCode', 'Sec_code', 'fnd_val','fnd_wgt', 'tot_fnd_val']
    dfprt_comp_agg_R_B=dfprt_comp_agg_R.merge(all_bmk_data, on = ['Port_code','BmkCode', 'Sec_code'], how='outer' )
    dfprt_comp_agg_R_B.loc[:,'fnd_wgt']=dfprt_comp_agg_R_B.fnd_wgt.fillna(0)
    dfprt_comp_agg_R_B.loc[:,'bmk_wgt']=dfprt_comp_agg_R_B.bmk_wgt.fillna(0)
    dfprt_comp_agg_R_B.loc[:,'fnd_val']=dfprt_comp_agg_R_B.fnd_val.fillna(0)
    dfprt_comp_agg_R_B.loc[:,'BmkValue']=dfprt_comp_agg_R_B.BmkValue.fillna(0)
    dfprt_comp.loc[:,'a_Sec_code']= np.where((dfprt_comp.Sec_code!=dfprt_comp.Cmb_code)&(dfprt_comp.Sec_type=='FUTURE : EQUITY'), dfprt_comp.Cmb_code,dfprt_comp.Sec_code) 
    dfprt_comp.loc[:,'a_Price']= np.where((dfprt_comp.Sec_code!=dfprt_comp.Cmb_code)&(dfprt_comp.Sec_type=='FUTURE : EQUITY'), dfprt_comp.Comp_price,dfprt_comp.Close_price) 
     
    dfprt_comp_agg_R_B_q=dfprt_comp_agg_R_B.merge(dfprt_comp[['Port_code','Quantity','a_Sec_code','Cmb_code']], left_on=['Port_code','Sec_code'],right_on=['Port_code','a_Sec_code'], how='left' )
    dfprt_comp_agg_R_B_q=dfprt_comp_agg_R_B_q.merge(dfprt_comp[['Port_code','a_Sec_code','a_Price']], left_on=['Port_code','Sec_code'],right_on=['Port_code','a_Sec_code'], how='left' ) 
    dfprt_comp_agg_R_B_q=dfprt_comp_agg_R_B_q[['Trade_date','Port_code','Sec_code', 'BmkCode','fnd_val','fnd_wgt','tot_fnd_val','Bmk_price','Bmk_quantity','BmkValue','bmk_wgt','Quantity','a_Price']]
    dfprt_comp_agg_R_B_q.loc[:,'Quantity']=dfprt_comp_agg_R_B_q.fillna(0)
    dfprt_comp_agg_R_B_q.loc[:,'U_Price']=np.where(dfprt_comp_agg_R_B_q.a_Price.isnull(), np.where(dfprt_comp_agg_R_B_q.Bmk_price.isnull(),np.nan,dfprt_comp_agg_R_B_q.Bmk_price.values), dfprt_comp_agg_R_B_q.a_Price.values)
    
    dfprt_comp_agg_R_B_q.loc[:,'act_bet']=dfprt_comp_agg_R_B_q.fnd_wgt-dfprt_comp_agg_R_B_q.bmk_wgt

    root = Tk()
    root.filename =  filedialog.askopenfilename(initialdir = '\\\\za.investment.int\\dfs\\dbshared\\DFM\\Benchmarks\\BlockList\\',
                                                title = "Please import Block list",filetypes = [("all files","*.*")])
  #  print (root.filename)
    root.withdraw()
    
    hd=pd.read_excel(root.filename)
    hd=hd.drop(['_type'], axis=1)
    
    excep_xls = (hd[hd.iloc[:,1]=='P'].iloc[:,0]).tolist()
    excl_xls = (hd[hd.iloc[:,1]!='P'].iloc[:]).set_index(['!ID']).T.to_dict('list')
    x_list=(dfprt_comp_agg_R_B_q[~dfprt_comp_agg_R_B_q.Sec_code.isin(list(excl_xls.keys()))].drop_duplicates(['Sec_code']))[['Sec_code','fnd_wgt']].set_index(['Sec_code']).T.to_dict('list')
    zxclusion = {**excl_xls , **x_list}        
    
    return [excep_xls,excl_xls,dfprt_comp_agg_R_B_q]
    
 


fund= 'ALSCPF'
buffer=0.0005
min_trd_thrs=0.0005
#sub_fnd=dfprt_comp_agg_R_B_q[dfprt_comp_agg_R_B_q.Port_code==fnd]
tgt_eff_cash=n_comb[n_comb.Port_code==fund]['fin_teff_cash']
"""
' 1=Buy
' 2=Sell
' 3=Two_way
"""

tdr_typ=1 #

def trade_fx( n_comb, dfprt_comp_agg_R_B_q,excl_xls, zxclusion , min_trd_thrs=0.0005, buffer=0.0005, fnd='ALSCPF', trade_type=3, excep=True,min_hold= 0.00001):
    
    import numpy as np
    import pandas as pd
    from write_excel import round_down
    import math
    
    
    if trade_type==1:
        rank_tab=(dfprt_comp_agg_R_B_q[dfprt_comp_agg_R_B_q.Port_code==fnd]).sort_values(['act_bet'], ascending = True)
    elif trade_type==2:
        rank_tab=(dfprt_comp_agg_R_B_q[dfprt_comp_agg_R_B_q.Port_code==fnd]).sort_values(['act_bet'], ascending = False)
    elif trade_type==3:
        rank_tab=(dfprt_comp_agg_R_B_q[dfprt_comp_agg_R_B_q.Port_code==fnd]).loc[np.array((dfprt_comp_agg_R_B_q[dfprt_comp_agg_R_B_q.Port_code==fnd]).act_bet.abs().sort_values(ascending=False).index),]
       # pos_bets=(rank_tab[((rank_tab.act_bet>0)&(rank_tab.Sec_code!='ZAR'))]['act_bet']).sum()
       # neg_bets=(rank_tab[((rank_tab.act_bet<0)&(rank_tab.Sec_code!='ZAR'))]['act_bet']).sum()
    
     
    
    ZAR_amt=rank_tab[rank_tab.Sec_code=='ZAR']['fnd_wgt'].values
    tgt_eff_cash=n_comb[n_comb.Port_code==fnd]['fin_teff_cash']
    eq_trd=np.round(n_comb[n_comb.Port_code==fnd]['eq_trade'],2).values
    
    ex_ZAR=rank_tab.copy()
    ex_ZAR=ex_ZAR[ex_ZAR.Sec_code!='ZAR']
    
    if min_trd_thrs > ex_ZAR.act_bet.abs().max():
        buffer = max(buffer,min_trd_thrs-round_down(ex_ZAR.act_bet.abs().max(), 4) )
        print("buffer:"+str(buffer))
    
    # Exclude the illiquid stocks at this point
    if excep:
        ex_ZAR=ex_ZAR[~ex_ZAR.Sec_code.isin(excep_xls)]
        
    ex_ZAR.loc[:,'use_bet']=np.where(ex_ZAR['act_bet']<0,(ex_ZAR['act_bet']-buffer).values, (ex_ZAR['act_bet']+buffer).values)
    
    if trade_type in [2,3]:
        ex_ZAR.loc[:,'use_bet']=np.where(ex_ZAR['use_bet']>0, 
                                         np.where(ex_ZAR['use_bet']>ex_ZAR['fnd_wgt'],ex_ZAR['fnd_wgt'].values, ex_ZAR['use_bet'].values),ex_ZAR['use_bet'].values)
        ex_ZAR.loc[:,'re_flag']=np.where(ex_ZAR['use_bet']>0, np.where(ex_ZAR['use_bet']>ex_ZAR['fnd_wgt'], 1 , 0) ,0)
    else:
        ex_ZAR.loc[:,'re_flag']=0
    
    if trade_type ==3:
        ex_ZAR.loc[:,'pos_bet_sells']=np.where((ex_ZAR['use_bet']>0)&(ex_ZAR['use_bet'].abs() >= min_trd_thrs),ex_ZAR['use_bet'].values,0)
        ex_ZAR.loc[:,'neg_bet_buys']=np.where((ex_ZAR['use_bet']<0)&(ex_ZAR['use_bet'].abs() >= min_trd_thrs),ex_ZAR['use_bet'].values,0)
        ex_ZAR.loc[:,'pos_bet_sells']=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())),0,ex_ZAR.pos_bet_sells.values)
        ex_ZAR.loc[:,'neg_bet_buys']=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())),0,ex_ZAR.neg_bet_buys.values)
#        if excep:
#            ex_ZAR.loc[:,'pos_bet_sells']= np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())), 
#                                              np.where(ex_ZAR.Sec_code.map(lambda x: zxclusion[x][0])<=ex_ZAR.fnd_wgt,ex_ZAR.fnd_wgt.values-ex_ZAR.Sec_code.map(lambda x:zxclusion[x][0]),ex_ZAR.pos_bet_sells.values),
#                                              ex_ZAR.pos_bet_sells.values)
#            
#            ex_ZAR.loc[:,'neg_bet_buys']= np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())), 
#                                              np.where(ex_ZAR.Sec_code.map(lambda x: zxclusion[x][0])>ex_ZAR.fnd_wgt,ex_ZAR.fnd_wgt.values-ex_ZAR.Sec_code.map(lambda x:zxclusion[x][0]),ex_ZAR.neg_bet_buys.values),
#                                              ex_ZAR.neg_bet_buys.values)
        if excep:
            pos_bet_excl=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())), 
                                              np.where(ex_ZAR.Sec_code.map(lambda x: zxclusion[x][0])<=ex_ZAR.fnd_wgt,
                                                       ex_ZAR.fnd_wgt.values-ex_ZAR.Sec_code.map(lambda x:zxclusion[x][0]),0),0).sum()
             
            neg_bet_excl=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())), 
                                              np.where(ex_ZAR.Sec_code.map(lambda x: zxclusion[x][0])>ex_ZAR.fnd_wgt,
                                                       ex_ZAR.fnd_wgt.values-ex_ZAR.Sec_code.map(lambda x:zxclusion[x][0]),0),0).sum()
            #net_sell= np.where(abs(pos_bet_excl)>abs(neg_bet_excl), pos_bet_excl-neg_bet_excl
            
        
        if (ZAR_amt>tgt_eff_cash.values)&(eq_trd!=0):
            to_bs=max(abs((ZAR_amt-tgt_eff_cash).values+abs(ex_ZAR['pos_bet_sells'].sum()))+pos_bet_excl, abs(ex_ZAR['neg_bet_buys'].sum()))
            
            if ((ZAR_amt-tgt_eff_cash).values+abs(ex_ZAR['pos_bet_sells'].sum())) > abs(ex_ZAR['neg_bet_buys'].sum()):
                pop_neg_bets=True
                print('A')
            else:
                pop_neg_bets=False
                print('B')
                
        elif (ZAR_amt < tgt_eff_cash.values)&(eq_trd!=0):
            to_bs=max(abs((ZAR_amt-tgt_eff_cash).values+abs(ex_ZAR['neg_bet_buys'].sum()))+abs(neg_bet_excl), abs(ex_ZAR['pos_bet_sells'].sum()))
            if abs((ZAR_amt-tgt_eff_cash).values+abs(ex_ZAR['neg_bet_buys'].sum())) > abs(ex_ZAR['pos_bet_sells'].sum()):
                pop_neg_bets=False
                print('C')
            else:
                pop_neg_bets=True
                print('D')
        else:
            to_bs=max(abs(ex_ZAR['neg_bet_buys'].sum()), abs(ex_ZAR['pos_bet_sells'].sum()))
            if (abs(ex_ZAR['pos_bet_sells'].sum()) > abs(ex_ZAR['neg_bet_buys'].sum())):
                pop_neg_bets=True
                print('E')
            else:
                pop_neg_bets=False
                print('F')
    else:
        pass
    
    if excep:
        if trade_type==1:
            ex_ZAR.loc[:,'re_flag']=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())),1, 0)
        else:
            ex_ZAR.loc[:,'re_flag']=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())),1, ex_ZAR.re_flag)
        if trade_type in [1,2]:
            ex_ZAR.loc[:,'use_bet']=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())), ex_ZAR.fnd_wgt.values-ex_ZAR.Sec_code.map(lambda x: zxclusion[x][0]),ex_ZAR.use_bet.values)
        
        
    ex_ZAR.loc[:,'abs_act_bet']=ex_ZAR['use_bet'].abs()
    
    if trade_type==3:
        tn_o=to_bs
    else:
        tn_o = min(abs((rank_tab[rank_tab.Sec_code=='ZAR']['fnd_wgt'].values)), 
                    max(rank_tab[rank_tab.act_bet<0]['act_bet'].abs().sum(),rank_tab[rank_tab.act_bet>0]['act_bet'].abs().sum())
                  )+np.where(trade_type==1,-tgt_eff_cash,tgt_eff_cash)+ \
                  np.where((ex_ZAR.re_flag==1)&(ex_ZAR.Sec_code.isin(list(excl_xls.keys()))),
                           np.where(trade_type==1, np.where(ex_ZAR.act_bet<0,ex_ZAR.use_bet,0 ),
                                    np.where(trade_type==2, np.where(ex_ZAR.act_bet>0,ex_ZAR.use_bet,0 ),0)),0).sum()
      #  print("Tgt eff cash is:", tgt_eff_cash)
    
    if trade_type in [1,2]:
        ex_ZAR.loc[:,'abs_act_bet']=np.where(ex_ZAR.Sec_code.isin(list(excl_xls.keys())),0,ex_ZAR['abs_act_bet'].abs())
        ex_ZAR.loc[:,'cum_trade']= np.where((ex_ZAR.use_bet.abs()>=min_trd_thrs)&(ex_ZAR.abs_act_bet.cumsum().values<=tn_o)
                                              ,ex_ZAR.abs_act_bet.values,0)
        ex_ZAR.loc[:,'cum_trade']=np.where((ex_ZAR.re_flag==1)&(ex_ZAR.Sec_code.isin(list(excl_xls.keys()))),
                                   np.where(trade_type==1, np.where(ex_ZAR.act_bet<0,ex_ZAR.use_bet.abs(),ex_ZAR.cum_trade),
                                    np.where(trade_type==2, np.where(ex_ZAR.act_bet>0,ex_ZAR.use_bet.abs(),ex_ZAR.cum_trade ),ex_ZAR.cum_trade)),ex_ZAR.cum_trade)
    elif trade_type==3:
        
       # if (abs(ex_ZAR['pos_bet_sells'].sum()) > abs(ex_ZAR['neg_bet_buys'].sum())):
        if pop_neg_bets:
            ex_ZAR['cum_trade']= ex_ZAR['neg_bet_buys'].abs().values
            ex_ZAR = ex_ZAR.sort_values(['neg_bet_buys'], ascending = True)
            h_tno=tn_o-ex_ZAR.loc[:,'cum_trade'].sum()-abs(neg_bet_excl)
        else:
            ex_ZAR['cum_trade']= ex_ZAR['pos_bet_sells'].abs().values
            ex_ZAR = ex_ZAR.sort_values(['pos_bet_sells'], ascending = False)
            
            
            
        excess=min(np.int((h_tno)/min_trd_thrs), len(ex_ZAR)-len(ex_ZAR[ex_ZAR.cum_trade!=0]))
        if excess>0:
        #    print("yes")
            r=(np.arange(len(ex_ZAR.index[ex_ZAR['cum_trade'] !=0].tolist()),len(ex_ZAR.index[ex_ZAR['cum_trade'] !=0].tolist())+(excess))).tolist()
            if len(r)<=1:
                r=np.floor(r[0]).astype(int)
            if trade_type in [2,3]:    
                ex_ZAR.cum_trade.iloc[r] = np.where(ex_ZAR['fnd_wgt'].iloc[r] < min_trd_thrs, ex_ZAR['fnd_wgt'].iloc[r] , min_trd_thrs)
                
                ex_ZAR.loc[:,'cum_trade'] = np.where((ex_ZAR.use_bet>0)&((ex_ZAR['fnd_wgt']-ex_ZAR['cum_trade']) < min_hold), ex_ZAR['fnd_wgt'].values , ex_ZAR['cum_trade'].values)
                ex_ZAR.loc[:,'re_flag'] = np.where((ex_ZAR['re_flag']==0)&((ex_ZAR['cum_trade'] < min_trd_thrs)|((ex_ZAR['fnd_wgt']-ex_ZAR['cum_trade'])<min_hold)), 1 , ex_ZAR['re_flag'].values) # Sell out
                
            elif trade_type==1:
                ex_ZAR.cum_trade.iloc[r] = min_trd_thrs
                ex_ZAR.loc[:,'re_flag'] = 0
            else:
                pass ## place holder for (two-way)
        else:
            pass
#            if trade_type==1:
#                pass
#            elif trade_type==2:
#                ex_ZAR.loc[:,'re_flag'] =ex_ZAR.loc[:,'re_flag'] 
#            else:
#                pass # placeholder
        
    else: # placholder
     #   ex_ZAR=ex_ZAR.sort_values(['abs_act_bet'],ascending = False)
     #   ex_ZAR.loc[:,'cum_trade']= np.where(((ex_ZAR.use_bet.abs()>=min_trd_thrs)&(ex_ZAR.abs_act_bet.cumsum().values<=abs(np.Inf))),
     #                                        ex_ZAR.abs_act_bet.values,0)*np.where(ex_ZAR['act_bet']<0,1,-1)
     #   ex_ZAR.loc[:,'re_flag'] = 0 # placeholder
         pass
        
        # need to add two-way here
        #.......
    if trade_type in [1,2]:        
        dif_pos=tn_o-ex_ZAR['cum_trade'].sum()-np.where((ex_ZAR.re_flag==1)&(ex_ZAR.Sec_code.isin(list(excl_xls.keys()))),
                               np.where(trade_type==1, np.where(ex_ZAR.act_bet<0,ex_ZAR.use_bet,0 ),
                                        np.where(trade_type==2, np.where(ex_ZAR.act_bet>0,ex_ZAR.use_bet,0 ),0)),0).sum()
      
    else:
        dif_pos=tn_o-ex_ZAR['cum_trade'].sum()
        
    ex_ZAR.loc[:,'part_trade']=dif_pos*np.where(ex_ZAR['re_flag']==1,0,(ex_ZAR.cum_trade/(ex_ZAR[ex_ZAR.re_flag==0].cum_trade.sum())))+ex_ZAR.cum_trade
    if trade_type==3:
        if pop_neg_bets:
            ex_ZAR['neg_bet_buys']=ex_ZAR['part_trade'].values
        else:
            ex_ZAR['pos_bet_sells']=-1*ex_ZAR['part_trade'].values
        
        ex_ZAR.loc[:,'part_trade']= ex_ZAR['neg_bet_buys'].fillna(0)-ex_ZAR['pos_bet_sells'].fillna(0).abs()
    else:
        pass

        
    rank_tab = rank_tab.merge(ex_ZAR[['Port_code','Sec_code','part_trade']], how='left', on = ['Port_code','Sec_code'])
    rank_tab.loc[:,'part_trade']=rank_tab.part_trade.fillna(0)
    
    if trade_type==1:
        rank_tab.loc[:,'part_trade']= np.where(rank_tab.Sec_code=='ZAR', -1*ex_ZAR.part_trade.sum(), rank_tab.part_trade.values)
        rank_tab.loc[:,'new_wgt']=rank_tab.part_trade.values+rank_tab.fnd_wgt.values
    elif trade_type==2:
        rank_tab.loc[:,'part_trade']= np.where(rank_tab.Sec_code=='ZAR', 1*ex_ZAR.part_trade.sum(), -1*rank_tab.part_trade.values)
        rank_tab.loc[:,'new_wgt']=rank_tab.part_trade.values+rank_tab.fnd_wgt.values
    else:
        rank_tab.loc[:,'part_trade']= np.where(rank_tab.Sec_code=='ZAR', -(1*ex_ZAR.part_trade.sum()), rank_tab.part_trade.values)
        rank_tab.loc[:,'new_wgt']=rank_tab.part_trade.values+rank_tab.fnd_wgt.values
            
        
    rank_tab.loc[:,'new_act_bet']=rank_tab.new_wgt-rank_tab.bmk_wgt
    max_bet=((rank_tab[(rank_tab.Sec_code!='ZAR')].new_act_bet.abs()).max())
    
    return [max_bet,buffer,rank_tab,ex_ZAR]

z1=trade_fx(n_comb, dfprt_comp_agg_R_B_q, excl_xls, zxclusion,min_trd_thrs=0.0005, buffer=0.0005, fnd='ALSCPF', trade_type=3)
#trade_fx(n_comb, dfprt_comp_agg_R_B_q, min_trd_thrs=0.0005, buffer=0.0005, fnd='DSALPC', trade_type=3)

#trade_fx(n_comb, dfprt_comp_agg_R_B_q, min_trd_thrs, buffer=0.00049731, tgt_eff_cash=n_comb[n_comb.Port_code==fund]['fin_teff_cash'],
#         fnd=fund, trade_type=2)

#out=trade_fx(n_comb, dfprt_comp_agg_R_B_q, min_trd_thrs, buffer, tgt_eff_cash=n_comb[n_comb.Port_code=='OMCD02']['fin_teff_cash'],
#         fnd='OMCD01', trade_type=2)
#out=trade_fx(n_comb, dfprt_comp_agg_R_B_q, min_trd_thrs, buffer, tgt_eff_cash=n_comb[n_comb.Port_code=='OMCM02']['fin_teff_cash'],
#         fnd='OMCM02', trade_type=2,excep=False)


#out[2].to_csv('c:\\data\omcd01.csv')


#dfprt_comp.to_csv('c:\\data\\test.csv')

#dfprt.AssetType3.unique()

        

# Map composites

# Map index futures

# Map SSF

# NPL


