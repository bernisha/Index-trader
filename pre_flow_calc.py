# -*- coding: utf-8 -*-
"""
Created on Mon Mar 26 13:09:43 2018

@author: blala
"""

#import future
import sys
#sys.path.append('C:\Program Files (x86)\WinPython\python-3.6.5.amd64\lib\site-packages\IPython\extensions')
#sys.path.append('C:\Program Files (x86)\WinPython\settings\.ipython')

#for p in sys.path:
#    print(p)

import numpy as np

import pandas as pd

import datetime as dt
from datetime import datetime, timedelta
import glob
import os
#from pydatastream import Datastream
#from business_calendar import Calendar, MO, TU, WE, TH, FR
import pyodbc
from write_excel import excel_fx as exl_rep
#from write_excel import input_fx as inp
from write_excel import select_fund as sf

np.seterr(divide='ignore', invalid='ignore')

#DWE =  Datastream(username="DS:ZTQN002", password="SPACE356",proxy='172.23.18.187:3128')
#DWE.system_info()
#DWE.sources()


#data_xls = pd.read_excel('\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES\\Decalog Valuation\\UFMPosCash20170811.xls', 'Position', index_col=None)
#type(data_xls)
#data_xls.tail()

"""
Set parameters for trading
"""
# Benchmark list

# Check before you run file
#run_prog=inp()

#if run_prog[0]=='N':
#   quit()
#else:

    #startDate = datetime.today().date()
startDate = datetime.today().date()
#startDate = datetime.strptime('Sep 15 2017', '%b %d %Y').date()
pd.options.display.max_rows = 200
#testing=True

# Benchmark settings
folder_yr = datetime.strftime(startDate, "%Y")
folder_mth = datetime.strftime(startDate, "%m")
folder_day = datetime.strftime(startDate, "%d")

# Fund settings
dirtoimport_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES\\Decalog Valuation\\'
#dirtoimport_file= 'H:\\Bernisha\\Work\\IndexTrader\\Data\\required_inputs\\'

# directory to export report to
#dirtooutput_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES\\Futures Trades'
dirtooutput_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES'
#output_folder='\\'.join([dirtooutput_file ,folder_yr, folder_mth])
output_folder=str('\\'.join([dirtooutput_file ,folder_yr, folder_mth,folder_day])+'\\Futures Trades')


if not os.path.exists(output_folder):
    os.makedirs(output_folder)


# Map fund and benchmark settings 

dic_om_index = {'DRIEQC':['OM Responsible Equity Fund','CSIESG'],
'DSWIXC':['SWIX Index Fund','JSESWIXALSICAPPED'],
'CORPEQ':['OMLAC Shareholder Protected Equity Portfolio','JSESWIXALSI'],
'MFEQTY':['M&F Protected Equity Portfolio','JSESWIXALSI'],
'DSALPC':['SA Listed Property Index Fund','JSESAPY'],
'USWIMF':['Momentum SWIX Index Fund','JSESWIXALSI'],
'OMRTMF':['RAFI40 Unit Trust','JSERAFI40'],
'LEUUSW':['Life Equity UPF','JSESWIXALSICAPPED'],
'LEIUSW':['Life Equity IPF','JSESWIXALSICAPPED'],
'SASEMF':['SASRIA','JSESWIXALSI'],
'BIDLMF':['Bidvest Life CAPI','JSECAPIALSI'],
'BIIDMF':['Bidvest Insurance CAPI','JSECAPIALSI'],
'ALSCPF':['Assupol CPF','JSESWIXALSI'],
'ALSIPF':['Assupol IPF','JSESWIXALSI'],
'ALSUPF':['Assupol UPF','JSESWIXALSI'],
'UMSMMF':['Samancor Group Provident Fund','JSESWIXALSI'],
'OMSI01':['OM CAPPED SWIX FUND','JSESWIXALSICAPPED'],
'UMSWMF':['Momentum SWIX 40 Index Fund','JSESWIX40'],
 #   'UMC1MF':['Anglo Corp CW','JSESWIXALSI'],
'OMALMF':['Top40 Unit Trust','JSETOP40'],
'DALSIC':['All Share Index Fund','JSEALSI'],
}

#dic_users={'blala':['BLL','blala'], 'test':['TST','test'], 'sbisho':['SB','sbisho']}
dic_users={'blala':['BLL','blala'], 'test':['TST','test'], 'sbisho':['SB','sbisho'], 'tmfelang2':['TM','tmfelang'], 'abalfour':['AB','abalfour'], 'sparker2':['SP','sparker'], 'fsibiya':['FS','fsibiya']}

#dic_om_index={
#            'DSALPC':['SA Listed Property Index Fund','JSESAPY',1,5,8,0.0005,0.0022,1,'Option 1 Gross Rate in cents per share'],
#            'CORPEQ':['OMLAC Shareholder Protected Equity Portfolio','JSESWIXALSI',1,5,8,0.0002,0.0013,2,'Option 1 Gross Rate in cents per share'],
#            'DALSIC':['All Share Index Fund','JSEALSI',1,5,8,0.0002,0.0013,3,'Option 1 Gross Rate in cents per share'],
#            'DSWIXC':['SWIX Index Fund','JSESWIXALSICAPPED',1,5,8,0.0002,0.0013,4,'Option 1 Gross Rate in cents per share'],
#            'MFEQTY':['M&F Protected Equity Portfolio','JSESWIXALSI',1,5,8,0.0002,0.0013,2,'Option 1 Gross Rate in cents per share'],
#            'USWIMF':['Momentum SWIX Fund','JSESWIXALSI',1,5,8,0.0002,0.0013,2,'Option 1 Gross Rate in cents per share'],
#            'BIIDMF':['Bidvest Insurance Fund','JSECAPIALSI',1,5,8,0.0002,0.0013,2,'Option 1 Gross Rate in cents per share'],
#            'LEUUSW':['Bidvest Insurance Fund','JSESWIXALSICAPPED',1,5,8,0.0002,0.0013,2,'Option 1 Gross Rate in cents per share'],
##            'OMCC01':['OM Core Conservative','JSESWIXALSICAPPED',1,5,8,0.0002,0.0013,2,'Option 1 Gross Rate in cents per share'],
#            }  

override=['SSF DIV']             

# Public Holidays
pub_holidays = (pd.read_excel("C:\\IndexTrader\\required_inputs\\public_holidays.xlsx"))['pub_holidays'].tolist()
#cal = Calendar(holidays=pub_holidays)

 # Determine list of funds to trade
lst_fund=sf()


# Import cash limits
cash_lmt_x = pd.read_csv('C:\\IndexTrader\\required_inputs\\cash_limits.csv')
cash_lmt_x=cash_lmt_x[cash_lmt_x.P_Code.isin(lst_fund)]
cash_lmt_dict=cash_lmt_x.set_index(['P_Code'])[['Min_EffCash','Max_EffCash']].T.to_dict()
    

# Import Flows
#cash_flows_eff = pd.read_csv('H:\\Bernisha\\Work\\IndexTrader\\Data\\required_inputs\\flows.csv')
cash_flows_eff = pd.read_csv('C:\\IndexTrader\\required_inputs\\flows.csv')
cash_flows_eff=(cash_flows_eff[cash_flows_eff.Port_code.isin(lst_fund)]).drop('Trade',1)


# Map Sec type to more descriptive asset classes

def assetClass(Sec_type, ins_code,sec_nam):

    #ssf=['OMLS'+str((cash_flows_eff['fut_sufx'].values)[0]), 'OMAS'+str((cash_flows_eff['fut_sufx'].values)[0])]
    ssf=['S']
    #excp=['OMLF'+str((cash_flows_eff['fut_sufx'].values)[0]),'OMAF'+str((cash_flows_eff['fut_sufx'].values)[0])]
    excp=['F']
    ind_fut=[str((cash_flows_eff['fut_sufx'].values)[0])] # index future suffix
    
    if Sec_type == 'CASH : CALL ACC':
        return "A. Total cash,Settled cash,Cash on call,Total cash,C. CALL"
    elif Sec_type=='CASH : SAFEX AC':
        return "A. Total cash,Settled cash,Futures margin,Total cash,D. SAFEX"
    elif Sec_type == "CURRENCY" and sec_nam=='VAL':
        return "A. Total cash,Settled cash,Val cash,Total cash,A. VAL"
    elif Sec_type=="PAYABLE" and sec_nam=='DIF':
        return "A. Total cash,Unsettled cash,Dif cash,Total cash,B. DIF"
    elif Sec_type=='FUTRE STCK INDX':
        return str("B. Futures Exposure,"+"Index Future,"+str(ins_code[0:4]+ind_fut[0])+",Futures Exposure,A. INDEX FUTURES")


#    elif Sec_type=='FUTURE : EQUITY' and ins_code in(ssf) :
    elif Sec_type=='FUTURE : EQUITY' and ins_code[3:4] in(ssf):
#        return str("Futures Exposure,"+"SSF,"+str(ssf[0]))
        return str("B. Futures Exposure,"+"SSF,null"+",Futures Exposure"+",B. SSF")
    elif Sec_type=='EQ : ORDINARY':
        return "Equity Exposure,Equity,null,Equity Exposure,EQUITY"
    elif Sec_type=='EQ : RIGHTS':
        return "Equity Exposure,Equity Rights,null,Equity Exposure,EQUITY"
    elif Sec_type=='EQ : FOREIGN':
        return "Equity Exposure,Equity Foreign,null,Equity Exposure,EQUITY"
    elif ins_code[3:4] in(excp):
#        return str("Dividend Exposure,"+"SSF Div,"+str(excp[0]))
        return str("Dividend Exposure,"+"SSF Div,null,Dividend Exposure,SSF DIV")
    elif Sec_type=="FUND : LOC EQ":
        return str("Equity Exposure,"+"Equity Fund,"+str(ins_code)+",Equity Exposure,EQUITY")
    else:
        return "Other,null,null,Other,OTHER"
        
        
def res_ind(dat,des,ind=['Trade_date','Port_code','AssetType1','AssetType2','AssetType3','AssetType4','AssetType5','Quantity','EffExposure','MarketValue','FundValue','Close_price']):
    dat=dat.reset_index()
    dat['AssetType1']=des
    dat['AssetType2']='null'
    dat['AssetType3']='null'
    dat['AssetType4']='null'
    dat['AssetType5']='null'
    dat=dat[ind]
    return dat
        
"""
Fund, Benchmark, Corporate Action data import
"""
#newest = max(glob.iglob(dirtoimport_file+'fund_data/*.xls'), key=os.path.getmtime)
newest = max(glob.iglob(dirtoimport_file+'*.xls'), key=os.path.getmtime)
#str(dirtoimport_file+newest)
#newest

fund_xls = pd.read_excel(newest,converters={'Portfolio code':str, 'Price date': pd.to_datetime, 
'Security type (name)':str, 
'Security name':str,
'Security ISIN code':str,
'Security acronym':str,
'Close price':float,
'Quantity held':float,
'Market price value':float},
)
#fund_xls.dtypes

fund_xls.columns = ['Port_code','Price_date','Sec_type','Sec_name','Sec_ISIN','Sec_code','Close_price','Quantity','Market_price']
fund_xls['Close_price']=pd.to_numeric(fund_xls.Close_price.values, errors='coerce') 
fund_xls['Quantity']=pd.to_numeric(fund_xls.Quantity.values, errors='coerce') 
fund_xls['Market_price']=pd.to_numeric(fund_xls.Market_price.values, errors='coerce') 

fund_obj = fund_xls.select_dtypes(['object'])
fund_xls[fund_obj.columns] = fund_obj.apply(lambda x: x.str.strip())


df=fund_xls.copy()
df=df[(df.Port_code.isin(dic_om_index.keys()))]
df=df[df.Port_code.isin(lst_fund)]
      
df.loc[:,'Benchmark_code']=df.Port_code.map(lambda x:dic_om_index[x][1])
      
df['Trade_date']=startDate
df['AssetType1']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name)).split(",")[0],axis=1)
df['AssetType2']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name)).split(",")[1],axis=1)
df['AssetType3']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name)).split(",")[2],axis=1)
df['AssetType4']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name)).split(",")[3],axis=1)
df['AssetType5']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name)).split(",")[4],axis=1)
df['MarketValue']= np.where(df[['AssetType1']].isin(['B. Futures Exposure','Dividend Exposure']),0, df[['Market_price']])
df['EffExposure']= df[['Market_price']]

df['Close_price']=np.where((df['AssetType2'].isin(['Index Future']))&(df['Quantity'].values!=0),
                                                      (df['Market_price'].values/df['Quantity'].values)/10, 
                                                       df['Close_price'].values)
                           
dfprt=(df[['Trade_date','Port_code','AssetType1','AssetType5','AssetType3','Sec_code','Close_price','Quantity','MarketValue','EffExposure']]).copy()

#Remove SSF Dividend Exposure
dfprt.loc[:,'EffExposure']=np.where(dfprt[['AssetType5']].isin(override),0, dfprt[['EffExposure']])
dfprt=dfprt[~(dfprt.Port_code.isnull())]
dfprt=dfprt[~(dfprt.Quantity.isnull())]
dfprt_preflow=dfprt.copy()
            
if ~cash_flows_eff.empty:
    cash_flows_eff['Trade_date']=startDate
    cash_flows_eff['AssetType1']='A. Total cash'
    cash_flows_eff['AssetType2']='Settled cash'
    cash_flows_eff['AssetType3']='Cash flow'
    cash_flows_eff['AssetType4']='Cash flow'
    cash_flows_eff['AssetType5']='Cash flow'
    cash_flows_eff['Sec_code']='ZAR'
    cash_flows_eff['Close_price']=1
    cash_flows_eff['Quantity']= cash_flows_eff[['Inflow']]
    cash_flows_eff['MarketValue']= cash_flows_eff[['Inflow']]
    cash_flows_eff['EffExposure']= cash_flows_eff[['Inflow']]
    cash_flows=cash_flows_eff[['Trade_date', 'Port_code','AssetType1', 'AssetType5', 'AssetType3', 'Sec_code','Close_price', 'Quantity','MarketValue','EffExposure']]
    

dfprt=dfprt.append(cash_flows)  


'''Generate cash calc
- Consolidate the current holdings of each fund 
1) Calculate the cash holdings
2) Calculate the futures exposure (both Index and SSF)
3) Calculate the equity market value
4) Calculate the special fund (cash and equity exposure)

'''
     

def fx_dta(dfprt_x=dfprt):
    dfprt_1=dfprt_x.groupby(['Trade_date','Port_code','AssetType1','AssetType5','AssetType3']).agg({'EffExposure':'sum','MarketValue':'sum','Quantity':'sum','Close_price':'max'})
    dfprt_1=dfprt_1.reset_index()
    dfprt_2= (dfprt_1.groupby(['Trade_date','Port_code']).agg({'MarketValue':'sum'})).reset_index()
    dfprt_1=pd.merge( dfprt_1,dfprt_2, on=['Trade_date','Port_code'])
    dfprt_1.rename(columns={'MarketValue_x':'MarketValue', 'MarketValue_y':'FundValue'}, inplace=True)
    dfprt_1=dfprt_1[['Trade_date','Port_code','AssetType1','AssetType5','AssetType3','MarketValue','EffExposure','Quantity','FundValue','Close_price']]
    dfprt_1=dfprt_1.groupby(['Trade_date','Port_code','AssetType1','AssetType5','AssetType3']).agg({'EffExposure':'sum','MarketValue':'sum','FundValue':'max','Quantity':'max','Close_price':'max'})
    
    req_sum={'EffExposure':'sum','MarketValue':'sum','FundValue':'max','Quantity':'max','Close_price':'max'}
    total_cash= (dfprt_1[(dfprt_1.index.get_level_values('AssetType1').isin(['A. Total cash']))]).reset_index().groupby(['Trade_date','Port_code']).agg(req_sum)
    
    effective_cash=((total_cash-(dfprt_1[(dfprt_1.index.get_level_values('AssetType1').isin(['B. Futures Exposure']))]).reset_index().groupby(['Trade_date','Port_code']).agg(req_sum)).fillna(0))
    effective_cash['MarketValue']=0
    effective_cash['FundValue']=total_cash[['FundValue']].values
    effective_cash['EffExposure']=np.where(effective_cash[['EffExposure']].values==0,total_cash[['EffExposure']].values, effective_cash[['EffExposure']].values)
    
    
    cash_dat=res_ind(effective_cash,'Effective cash').reset_index()
    cash_dat['Trade_date']=startDate
    cash_dat=(cash_dat[['Trade_date', 'Port_code','AssetType1','AssetType5','AssetType3', 'Quantity','EffExposure','MarketValue','FundValue','Close_price']])
    new_dat=((pd.concat([dfprt_1.reset_index(),cash_dat],axis=0).reset_index().drop('index',axis=1)).sort_values(['Port_code','AssetType1','AssetType5','AssetType3'])).set_index(['Trade_date','Port_code','AssetType1','AssetType5','AssetType3'])
    new_dat['EffWgt']=new_dat[['EffExposure']].values/new_dat[['FundValue']].values
    new_dat['MktWgt']=new_dat[['MarketValue']].values/new_dat[['FundValue']].values
    n_1 = new_dat.reset_index()
    n_1=n_1.groupby(['Port_code','AssetType1']).agg({'EffExposure':'sum','EffWgt':'sum'})
    n_1=n_1[~(n_1.index.get_level_values('AssetType1').isin(['Dividend Exposure']))]
    n_2=n_1.reset_index()
    fnd_value=(total_cash[['FundValue']].reset_index().set_index('Port_code')[['FundValue']]).reset_index()
    fnd_value['AssetType1']='Fund Value'
    fnd_value['EffWgt']=1
    fnd_value.columns= ['Port_code','EffExposure','AssetType1','EffWgt']  
    fnd_value=fnd_value[n_2.columns]
    n_3=n_2.append(fnd_value)
    n_3=n_3.reset_index().pivot(index='Port_code', columns='AssetType1',values='EffExposure')
    n_4=n_2.reset_index().pivot(index='Port_code', columns='AssetType1',values='EffWgt')
    
    n_3.columns=[sym.replace(" ", "")+'_R'  for sym in n_3.columns]
    n_4.columns=[sym.replace(" ", "")+'_p'  for sym in n_4.columns]
    
    n_comb=n_3.merge(n_4, left_index=True, right_index=True)   
    n_comb[['B.FuturesExposure_R']]=(n_comb[['B.FuturesExposure_R']]).fillna(0)   
    n_comb[['B.FuturesExposure_p']]=(n_comb[['B.FuturesExposure_p']]).fillna(0)   
    lst = [new_dat, n_comb]
    return lst

# Pre flow
new_dat_preflow=fx_dta(dfprt_x=dfprt_preflow)
new_dat_pf=new_dat_preflow[0]
n_comb_pf=new_dat_preflow[1]

# Post flow
new_dat_x=fx_dta(dfprt_x=dfprt)
new_dat=new_dat_x[0]
n_comb=new_dat_x[1]
        
    
    
    #fut_price=((new_dat[new_dat.index.get_level_values('AssetType2')=='Index Future']['Close_price']).reset_index())[['Port_code','AssetType3','Close_price']]        
    no_fut=((new_dat[new_dat.index.get_level_values('AssetType2')=='Index Future'][['Quantity']]).reset_index())[['Port_code','AssetType3','Quantity']]        
    fut_code1=(cash_lmt_x[['P_Code', 'Future_Code']]).copy()
    fut_code1.loc[:,'Future_Code']= np.where(fut_code1['Future_Code']=='NoFuture', 'NoFuture', fut_code1['Future_Code']+str((cash_flows_eff['fut_sufx'].values)[0]))
    no_fut = no_fut.merge(fut_code1, how='right',left_on=['Port_code'],right_on=['P_Code'] )
    no_fut['AssetType3'] = np.where(no_fut.AssetType3.isnull(), no_fut.Future_Code.values, no_fut.AssetType3.values)
    no_fut=(no_fut[['P_Code','AssetType3','Quantity']]).merge((((new_dat[new_dat.index.get_level_values('AssetType2')=='Index Future'][['Close_price']]).reset_index())[['AssetType3','Close_price']]).drop_duplicates(['AssetType3']) ,
                           how='left',left_on=['AssetType3'],right_on=['AssetType3'])
    no_fut=no_fut.fillna(0)
    no_fut.columns= ['Port_code', 'AssetType3',  'Quantity',  'Close_price']
    
    n_comb=(n_comb.reset_index()).merge(no_fut, how='left',left_on=['Port_code'],right_on=['Port_code'])
    
    # Get Inflow information & override effective cash if applicable 
    
    cash_lmt=pd.merge(cash_lmt_x, cash_flows_eff[['Port_code','Eff_cash','Inflow']], how='left',left_on=['P_Code'],right_on=['Port_code']) 
    cash_lmt.pop('Port_code')
    cash_lmt=cash_lmt.rename(columns = {'Eff_cash':'Ovd_Effcash'})
    cash_lmt['Tgt_EffCash1']=np.where(cash_lmt[['Ovd_Effcash']].isnull(), cash_lmt[['Tgt_EffCash']].values, cash_lmt[['Ovd_Effcash']])
    
    
    # Get Futures codes
    
    get_Futurecodes=fut_code1
    cash_lmt=cash_lmt.merge(get_Futurecodes, how='left',left_on=['P_Code'],right_on=['P_Code'])
    #cash_lmt=cash_lmt.drop(['Port_code'], axis=1)
    
    
    n_comb=pd.merge(n_comb, cash_lmt, how='left',left_on=['Port_code'],right_on=['P_Code'])
    n_comb.loc[:,'FundValue_p']=1
    
    # Create Breach Flag
    # Emx_Tmx - Both Eff cash and Total cash above max
    # Emx_Tmn - Eff cash above max and Total cash below mn
    # Emx_Twb - Eff cash above max, Total cash within bounds
    # Emn_Tmx - Eff cash below min, Total cash above max
    # Emn_Tmn - Eff cash below min, Total cash below min
    # Emn_Twb - Eff cash below min, Total cash within bounds
    # Ewb_Twb - No breach
    
    n_comb.loc[:,'Flag']= np.where((n_comb['Effectivecash_p'].values>=n_comb['Max_EffCash'].values),
                              np.where((n_comb['Totalcash_p'].values>=n_comb['Max_TotalCash'].values), 
                                       'Emx_Tmx',
                                np.where((n_comb['Totalcash_p'].values<=n_comb['Min_TotalCash'].values),
                                         'Emx_Tmn',
                                          'Emx_Twb')),
                                   np.where((n_comb['Effectivecash_p'].values<=n_comb['Min_EffCash'].values),
                                       np.where((n_comb['Totalcash_p'].values>=n_comb['Max_TotalCash'].values),
                                                'Emn_Tmx',
                                   np.where((n_comb['Totalcash_p'].values<=n_comb['Min_TotalCash'].values),
                                            'Emn_Tmn',
                                            'Emn_Twb')), 'Ewb_Twb'
                                      ))
                                             
    
    # Scenario 1 Futures trade
    n_comb.loc[:,'Trade']= np.where(n_comb.Future_Code_y=="NoFuture", "No Trade",
                                    np.where((n_comb['Effectivecash_p'].values>=n_comb['Max_EffCash'].values), 'Buy', 
                                             np.where((n_comb['Effectivecash_p'].values<=n_comb['Min_EffCash'].values), 'Sell', 'No Trade')))
    n_comb.loc[:,'Trade']=np.where((n_comb.Trade=="No Trade")&(~n_comb.Ovd_Effcash.isnull())&(~(n_comb.Future_Code_y=="NoFuture")),
                                    np.where((n_comb['Effectivecash_p'].values>=n_comb['Tgt_EffCash1'].values), 'Buy', 
                                             np.where((n_comb['Effectivecash_p'].values<=n_comb['Tgt_EffCash1'].values), 'Sell', 'No Trade')),n_comb.Trade)
    n_comb.loc[:,'No. Futures']=np.where(n_comb[['Trade']].isin(['Buy','Sell']), np.rint(((n_comb[['Effectivecash_p']].values-n_comb[['Tgt_EffCash1']].values)*n_comb[['FundValue_R']].values)/(n_comb[['Close_price']].values*10)), 0)
    n_comb.loc[:,'FuturesExposure_TT']=(n_comb[['No. Futures']].values)*10*n_comb[['Close_price']].values
    n_comb.loc[:,'FuturesExposure_TTp']=n_comb[['FuturesExposure_TT']].values/n_comb[['FundValue_R']].values
 
    # Equity Trade
    
    n_comb.loc[:,'TradeEq']= np.where(n_comb.Future_Code_y=="NoFuture", "No Trade",
                                    np.where((n_comb['Effectivecash_p'].values>=n_comb['Max_EffCash'].values), 'Buy', 
                                             np.where((n_comb['Effectivecash_p'].values<=n_comb['Min_EffCash'].values), 'Sell', 'No Trade')))
    
    
    
    
    # Check for negative effective cash
    n_comb.loc[:,'Trade'] = np.where(((-(n_comb[['No. Futures']].values*n_comb[['Close_price']].fillna(0).values*10)/n_comb[['FundValue_R']].values)+n_comb[['Effectivecash_p']].values)<0, 0, n_comb[['Trade']])
    n_comb.loc[:,'No. Futures']=np.where(((-(n_comb[['No. Futures']].values*n_comb[['Close_price']].fillna(0).values*10)/n_comb[['FundValue_R']].values)+n_comb[['Effectivecash_p']].values)<0, 0, n_comb[['No. Futures']])
    
    n_comb.loc[:,'Trade']=np.where(n_comb['No. Futures']==0, 'No Trade', n_comb.Trade.values)
    
    
    n_comb.loc[:,'Effectivecash_T']= (-(n_comb[['No. Futures']].values*n_comb[['Close_price']].fillna(0).values*10)/n_comb[['FundValue_R']].values)+n_comb[['Effectivecash_p']].values
    
    
    n_comb.loc[:,'Effectivecash_TR']= n_comb[['Effectivecash_T']].values*n_comb[['FundValue_R']].values
    
    n_comb.loc[:,'EquityExposure_T']=n_comb[['EquityExposure_p']].values
    n_comb.loc[:,'EquityExposure_TR']=n_comb[['EquityExposure_R']].values
    
    n_comb.loc[:,'FundValue_T']=1
    n_comb.loc[:,'FundValue_TR']=n_comb[['FundValue_R']].values
    
    
    n_comb.loc[:,'FuturesExposure_T']=n_comb[['FuturesExposure_p']].fillna(0).values+((n_comb[['No. Futures']].values*n_comb[['Close_price']].fillna(0).values*10)/n_comb[['FundValue_R']].values)
    n_comb.loc[:,'FuturesExposure_TR']=n_comb[['FuturesExposure_R']].fillna(0).values+((n_comb[['No. Futures']].values*n_comb[['Close_price']].fillna(0).values*10))
    
    n_comb.loc[:,'Totalcash_T']=n_comb[['Totalcash_p']].values
    n_comb.loc[:,'Totalcash_TR']=n_comb[['Totalcash_R']].values
    n_comb.loc[:,'Check cash']=np.where((n_comb['Totalcash_T'].values>=n_comb['Max_TotalCash'].values), 'Reduce cash', 
                                             np.where((n_comb['Totalcash_T'].values<=n_comb['Min_TotalCash'].values), 'Increase cash', ''))
    n_comb.loc[:,'Inflow_p'] = n_comb['Inflow']/n_comb['FundValue_R'].values
    
    
    # Add Pre-Flow
    
    n_comb_pf.columns = [str(col) + '_pf' for col in n_comb_pf.columns]
    n_comb=pd.merge(n_comb, n_comb_pf.reset_index(), how='left',left_on=['Port_code'],right_on=['Port_code'])
    n_comb.loc[:,'FundValue_p_pf']=1
   
    
    n_comb_eff_n=n_comb[['Port_code','FundValue_R_pf','EquityExposure_R_pf','Totalcash_R_pf','FuturesExposure_R_pf','Effectivecash_R_pf',
                       'FundValue_R','EquityExposure_R','Totalcash_R','FuturesExposure_R','Effectivecash_R','Tgt_EffCash1', 'No. Futures',
                       'AssetType3','Trade','FundValue_TR','EquityExposure_TR','Totalcash_TR','FuturesExposure_TR','Effectivecash_TR','Check cash','Min_EffCash',
                       'Max_EffCash', 'Min_TotalCash', 'Max_TotalCash','Tgt_EffCash','Inflow']]
    
    n_comb_eff=n_comb_eff_n.copy()
    n_comb_eff.loc[:,'ExposureType']=''
    n_comb_eff.loc[:,'Tgt_EffCash1']=np.nan
    n_comb_eff.loc[:,'No. Futures']=n_comb[['Close_price']].values
    n_comb_eff.loc[:,'Trade']=np.nan
    n_comb_eff.loc[:,'Check cash']=np.nan
    n_comb_eff.loc[:,'Min_EffCash']=np.nan
    n_comb_eff.loc[:,'Max_EffCash']=np.nan
    n_comb_eff.loc[:,'Min_TotalCash']=np.nan
    n_comb_eff.loc[:,'Max_TotalCash']=np.nan
    n_comb_eff.loc[:,'Tgt_EffCash']=np.nan
    n_comb_eff.loc[:,'AssetType3']=np.nan
    
    n_comb_eff_n=n_comb[['Port_code','FundValue_p_pf','EquityExposure_p_pf','Totalcash_p_pf','FuturesExposure_p_pf','Effectivecash_p_pf',
                       'FundValue_p','EquityExposure_p','Totalcash_p','FuturesExposure_p','Effectivecash_p','Tgt_EffCash1', 'No. Futures',
                       'AssetType3','Trade','FundValue_T','EquityExposure_T','Totalcash_T','FuturesExposure_T','Effectivecash_T','Check cash','Min_EffCash',
                       'Max_EffCash', 'Min_TotalCash', 'Max_TotalCash','Tgt_EffCash','Inflow_p']]
    n_comb_effp=n_comb_eff_n.copy()
    n_comb_effp.loc[:,'ExposureType']='(%)'
    
    
    n_comb_effp.columns = n_comb_eff.columns
    
    
    n_comb_eff=n_comb_eff.append(n_comb_effp)
    n_comb_eff.loc[:,'Checked']=''
    #n_comb_eff.loc[:,'Authorised']=''
    #n_comb_eff.sort(['Port_code','ExposureType'], ascending=False).set_index(['Port_code','ExposureType']).to_csv('c:\data\check.csv')
    
    n_comb_eff_1=n_comb_eff.sort_values(['Port_code','ExposureType'], ascending=True).set_index(['Port_code','ExposureType'])
    n_comb_eff_1['Trade_YN']=''   
    n_comb_eff_1['Comment']=''
    # write excel report
    exl_rep(output_folder,dic_users,n_comb_eff_1,startDate)
    #excel_fx(output_folder,dic_users,n_comb_eff_1,startDate)
    #exl_rep('c:\\data\\',dic_users,n_comb_eff_1,startDate)
    print("\nReport Complete")
    #n_comb_eff.sort(['Port_code','ExposureType'], ascending=False).set_index(['Port_code','ExposureType']).to_html(open('c:\data\check.html','w'),formatters='{:,.0f}'.format)
    
    
    #grp=n_comb_eff.groupby(['Port_code','ExposureType']).agg({'FundValue_R':'sum','EquityExposure_R':'sum'})
    
    
    #n_comb_eff=n_comb_eff.pivot(index='Port_code', columns='ExposureType')
    
    
    
    
    #n_comb.to_csv('c:\data\check.csv')
        
    
        