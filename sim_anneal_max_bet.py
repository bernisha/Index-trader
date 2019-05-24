# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import numpy as np


#trade_fx(n_comb, dfprt_comp_agg_R_B_q, min_trd_thrs, buffer, trade_type=1, fnd='ALSCPF')


def sim_annl_fx(func,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,min_trd_thrs,tdr_typ,exp, min_hldg,fund, s0, niter=1000, step=0.1, ex_buf=0.00001):
  # Initialize
  ## s stands for state
  ## f stands for function value
  ## b stands for best
  ## c stands for current
  ## n stands for neighbor
  s_b = s0
  s_c = s0
  s_n = s0
  
  s_list=[]
  fn_list=[]
  
  f_b_ = func(n_comb, dfprt_comp_agg_R_B_q,excep_xls, excl_xls, zxclusion,min_trd_thrs, s0,fnd=fund,trade_type=tdr_typ, excep=exp,min_hold=min_hldg) ## calculate the max active bet (max)
  f_b = f_b_[0]
  f_c = f_b
  f_n = f_c
  cnt=0
#  print("Iteration:", cnt, "The current state value: ", s_c, ", current function value is: ", f_c ) 
 
  for k in range(1, niter):
      tmp = (1-step)**k
      s_n = s_c + np.random.normal(0,0.000005,1)
      f_n_ = func(n_comb, dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion, min_trd_thrs, s_n, fnd=fund,trade_type=tdr_typ,excep=exp,min_hold=min_hldg)
      f_n=f_n_[0]
      
      # update current state
      if ((s_n>0)and((f_n < f_c)and(f_n >= (s0-ex_buf))and(f_n < (s0+ex_buf)))or(np.random.uniform(0.0, 1.0, 1) < np.exp(-(f_n - f_c)/tmp))):
          s_c = s_n
          f_c = f_n
   #       print("Iteration:", cnt, "The current state value: ", s_c, ", current function value is: ", f_c )
          s_list.append(s_c)
          fn_list.append(f_n)
 
    # update best state
      if ((s_n>0)and((f_n < f_b)and(f_n >= (min_hldg-ex_buf))and(f_n < (min_hldg+ex_buf)))and(k<(niter-1))):
        s_b = s_n
        f_b = f_n
        i_b = k
        cnt = cnt+1
        f_b_ = f_n_
        print("k is:", k, " cnt is:", cnt, "i_b is:", i_b)
      if ('i_b' in locals()):
          if ((k>10)and((k-cnt)>30)and((i_b-cnt)>10)):
              print("Time to stop")
           #   print("k is:", k, " cnt is:", cnt, "i_b is:", i_b)
              break
                
      elif ((k==(niter-1))and(~('s_b' in locals()))):
          print("Can't find a solution")
          lst=abs(s0-np.array(fn_list))#np.where((buffer-np.array(fn_list))>0,buffer-np.array(fn_list),9)
          s_b = s_list[[i for i,x in enumerate(lst) if x == min(lst)][0]]
          f_b_ = func(n_comb, dfprt_comp_agg_R_B_q,excep_xls, excl_xls, zxclusion,min_trd_thrs, s_b, fnd=fund,trade_type=tdr_typ,excep=exp,min_hold=min_hldg)
          f_b = f_b_[0]
          i_b = 999
          break
   #   print("k is:", k, " cnt is:", cnt)
  print("The best state value: ", s_b, ", best function value is: ", f_b, 'found at iteration:', i_b )
  return [niter, s_b, f_b, i_b, tmp, cnt, f_b_[2]]
  del [niter, s_b, f_b, i_b, tmp, cnt ]
    
  
def fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer,min_trd_thrs,tdr_typ,fund,exp,min_hldg,nt=100,stp=0.05,ex_bf=0.000001,mx_bet=0.0005):  
    f_cnt=1  
    ooh=sim_annl_fx(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,min_trd_thrs,tdr_typ,exp,min_hldg,fund,s0=buffer,niter=nt,step=stp,ex_buf=ex_bf)
    ooh_aah=ooh[2]
    while((f_cnt<7)&(ooh[3]==999)&(ooh[2]>mx_bet)&(f_cnt>1)&(abs(ooh_aah-ooh[2])>1e-10)):
        buffer=buffer/2
        ooh_aah=ooh[2]
        ooh=sim_annl_fx(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,min_trd_thrs,tdr_typ, exp,min_hldg,fund,s0=buffer,niter=nt,step=stp,ex_buf=ex_bf)
        f_cnt+=1
        
        print(f_cnt)
    else:
        print('solution found')    
    
    trades=ooh[6]
    trades['tradesX']=np.where(trades['Sec_code']!='ZAR',np.floor(((trades['part_trade']*trades['tot_fnd_val'])/trades['U_Price']).fillna(0).values),0)
    trades['tradesY']=np.where(trades['Sec_code']=='ZAR', -(((trades['tradesX']*trades['U_Price'])).fillna(0).values).sum(),trades['tradesX'].values )
    trades['fnl_fnd_wgt']=trades['fnd_wgt'].values + ((trades['tradesY']*trades['U_Price'])/trades['tot_fnd_val']).fillna(0).values
    trades['fnl_fnd_wgt_check']=np.where((trades['fnl_fnd_wgt'].values <0),0, trades['fnl_fnd_wgt'].values) # shot-sell
    trades['trades']=np.where(trades['Sec_code'].isin(excep_xls), 0, 
                      np.where(trades['fnl_fnd_wgt'].values<0,np.where(trades.tradesY>0,trades['Quantity'].values,-trades['Quantity'].values),  trades['tradesY'].values))  # short sell
    trades['fnl_act_bet']=((trades['fnl_fnd_wgt_check']-trades['bmk_wgt']).fillna(0)).values
    trades=trades.drop(['tradesX','tradesY','fnl_fnd_wgt'], axis=1)
    trades.rename(columns={'fnl_fnd_wgt_check':'fnl_fndwgt'}, inplace=True)
    trades['Action']=np.where(trades.trades>0, 'B',np.where(trades.trades<0,'S','N'))
    cnt=str('Trades:'+str(len(trades[trades.trades.abs()>0]))+ ',Buys:'+ str(len(trades[trades.Action=='B']))+ ',Sells:'+str(len(trades[trades.Action=='S'])))
    ooh.pop(6)

    return [ooh+[cnt]+[trades]]
#    return [ooh+[trades]]

z1=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0000,min_trd_thrs=0.0005,tdr_typ=3,fund='ALSCPF',exp=True,min_hldg=0.00001,mx_bet=0.0005)
z2=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0003,min_trd_thrs=0.0005,tdr_typ=1,fund='ALSUPF',exp=True,min_hldg=0.00001,mx_bet=0.0005)

#z2=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0005,min_trd_thrs=0.0005,tdr_typ=2,fund='DSALPC',exp=False,min_hldg=0.0001)
z3=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0000,min_trd_thrs=0.0005,tdr_typ=3,fund='UMSMMF',exp=True,min_hldg=0.00001)
z4=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0000,min_trd_thrs=0.0005,tdr_typ=3,fund='OMALMF',exp=True,min_hldg=0.00001)


z5=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0000,min_trd_thrs=0.0005,tdr_typ=3,fund='OMCD01',exp=True,min_hldg=0.00001)

#z5=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0000,min_trd_thrs=0.0005,tdr_typ=3,fund='UMSWMF',exp=True,min_hldg=0.00001)

z6=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0003,min_trd_thrs=0.0005,tdr_typ=1,fund='SASEMF',exp=True,min_hldg=0.00001)
z7=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,excep_xls,excl_xls, zxclusion,buffer=0.0001,min_trd_thrs=0.0005,tdr_typ=3,fund='CORPEQ',exp=True,min_hldg=0.00001)

fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,buffer,min_trd_thrs,tdr_typ=2,fund='SASEMF')

z2=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,buffer,min_trd_thrs,tdr_typ=2,fund='DSALPC')

fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,buffer,min_trd_thrs,tdr_typ=1,fund='OMCD01')
z3=fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,buffer,min_trd_thrs,tdr_typ=2,exp=False,fund='OMCM02',)

fnd_best(trade_fx,n_comb,dfprt_comp_agg_R_B_q,buffer,min_trd_thrs,tdr_typ=1,fund='DALSIC')
