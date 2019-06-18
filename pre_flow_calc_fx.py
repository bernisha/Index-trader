
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 26 13:09:43 2018
@author: blala
"""
def pre_flow_calcFx(response,automatic=True,orders=False,testing=False):
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
    #from write_excel import excel_fx as exl_rep
    #from write_excel import input_fx as inp
    from write_excel import select_fund as sf
    from write_excel import CashFlowFlag as cff
    from write_excel import trade_calc as t_c
    from write_excel import trade_calc_automatic as t_c_a
    from write_excel import bulk_cash_excel_report as bcer
    from write_excel import cash_flow_validity_fx as cfvf
    from write_excel import assetClassB as assetClass
    from write_excel import res_indB as res_ind
    from write_excel import fx_dtaB as fx_dta
    
   
    if testing:
        response= 'yes'
        automatic = True
        orders=False
        
    
    if response:
        start_time = datetime.now() 
        
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
        startDate = datetime.today()
        #startDate = datetime.strptime('Sep 15 2017', '%b %d %Y').datetime()
        #startDate = datetime.today()- timedelta(days=1)
        pd.options.display.max_rows = 200
        #testing=True
        
        # Benchmark settings
        folder_yr = datetime.strftime(startDate, "%Y")
        folder_mth = datetime.strftime(startDate, "%m")
        folder_day = datetime.strftime(startDate, "%d")
        
        # Fund settings
       # dirtoimport_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES\\Decalog Valuation\\'
        #dirtoimport_file= 'H:\\Bernisha\\Work\\IndexTrader\\Data\\required_inputs\\'
        dirtoimport_file='\\\\za.investment.int\\DFS\\SSDecalogUmbono\\IndexationPosFile\\'
        dirtoimport_cashfile = '\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES\\Decalog Valuation\\' 
     
        
        # directory to export report to
        #dirtooutput_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES\\Futures Trades'
        dirtooutput_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES'
        #output_folder='\\'.join([dirtooutput_file ,folder_yr, folder_mth])
        output_folder=str('\\'.join([dirtooutput_file ,folder_yr, folder_mth,folder_day])+'\\BatchTrades')
        
        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        
        # Map fund and benchmark settings 
        
          # Pull in fund dictionary
        fnd_dict=pd.read_csv('C:\\IndexTrader\\required_inputs\\fund_dictionary.csv')
        dic_om_index=fnd_dict.set_index(['FundCode']).T.to_dict('list')
            
            
        user_dict=pd.read_csv('C:\\IndexTrader\\required_inputs\\user_dictionary.csv')
        dic_users=user_dict.set_index(['username']).T.to_dict('list')
            
        fnd_excp= ['DSALPC','OMCC01','OMCD01','OMCD02','OMCM01','OMCM02','PPSBTA','PPSBTB']
        
        #dic_om_index = {'DRIEQC':['OM Responsible Equity Fund','CSIESG'],
        #'DSWIXC':['SWIX Index Fund','JSESWIXALSICAPPED'],
        #'CORPEQ':['OMLAC Shareholder Protected Equity Portfolio','JSESWIXALSI'],
        #'MFEQTY':['M&F Protected Equity Portfolio','JSESWIXALSI'],
        #'DSALPC':['SA Listed Property Index Fund','JSESAPY'],
        #'USWIMF':['Momentum SWIX Index Fund','JSESWIXALSI'],
        #'OMRTMF':['RAFI40 Unit Trust','JSERAFI40'],
        #'LEUUSW':['Life Equity UPF','JSESWIXALSICAPPED'],
        #'LEIUSW':['Life Equity IPF','JSESWIXALSICAPPED'],
        #'SASEMF':['SASRIA','JSESWIXALSI'],
        #'BIDLMF':['Bidvest Life CAPI','JSECAPIALSI'],
        #'BIIDMF':['Bidvest Insurance CAPI','JSECAPIALSI'],
        #'ALSCPF':['Assupol CPF','JSESWIXALSI'],
        #'ALSIPF':['Assupol IPF','JSESWIXALSI'],
        #'ALSUPF':['Assupol UPF','JSESWIXALSI'],
        #'UMSMMF':['Samancor Group Provident Fund','JSESWIXALSI'],
        #'OMSI01':['OM CAPPED SWIX FUND','JSESWIXALSICAPPED'],
        #'UMSWMF':['Momentum SWIX 40 Index Fund','JSESWIX40'],
        # #   'UMC1MF':['Anglo Corp CW','JSESWIXALSI'],
        #'OMALMF':['Top40 Unit Trust','JSETOP40'],
        #'DALSIC':['All Share Index Fund','JSEALSI'],
        #}
        #
        ##dic_users={'blala':['BLL','blala'], 'test':['TST','test'], 'sbisho':['SB','sbisho']}
        #dic_users={'blala':['BLL','blala'], 'test':['TST','test'], 'sbisho':['SB','sbisho'], 'tmfelang2':['TM','tmfelang'], 'abalfour':['AB','abalfour'], 'sparker2':['SP','sparker'], 'fsibiya':['FS','fsibiya']}
        
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
        cash_flows_eff = pd.read_csv('C:\\IndexTrader\\required_inputs\\flows.csv',thousands=',')
        cash_flows_act_name = cash_flows_eff[cash_flows_eff.Trade==1].Port_code
        cash_flows_eff=(cash_flows_eff[cash_flows_eff.Port_code.isin(lst_fund)]).drop('Trade',1)
        
        
        
         # Import futures
        fut_flow=pd.merge(cash_lmt_x[['P_Code','Future_Code']], cash_flows_eff[['Port_code','fut_sufx']], how='right', left_on=['P_Code'], right_on=['Port_code'] )
        fut_flow['Sec_code']= fut_flow[['Future_Code', 'fut_sufx']].apply(lambda x: ''.join(x), axis=1)
        fut_flow['Sec_code']=np.where(fut_flow.Future_Code=='NoFuture','NoFuture',fut_flow.Sec_code.values)
            
        # Map Sec type to more descriptive asset classes
        
        #def assetClassB(Sec_type, ins_code,sec_nam,cash_flows_eff):
        #
        #    #ssf=['OMLS'+str((cash_flows_eff['fut_sufx'].values)[0]), 'OMAS'+str((cash_flows_eff['fut_sufx'].values)[0])]
        #    ssf=['S']
        #    #excp=['OMLF'+str((cash_flows_eff['fut_sufx'].values)[0]),'OMAF'+str((cash_flows_eff['fut_sufx'].values)[0])]
        #    excp=['F']
        #    ind_fut=[str((cash_flows_eff['fut_sufx'].values)[0])] # index future suffix
        #    
        #    if Sec_type == 'CASH : CALL ACC':
        #        return "A. Total cash,Settled cash,Cash on call,Total cash,C. CALL"
        #    elif Sec_type=='CASH : SAFEX AC':
        #        return "A. Total cash,Settled cash,Futures margin,Total cash,D. SAFEX"
        #    elif Sec_type == "CURRENCY" and sec_nam=='VAL':
        #        return "A. Total cash,Settled cash,Val cash,Total cash,A. VAL"
        #    elif Sec_type=="PAYABLE" and sec_nam=='DIF':
        #        return "A. Total cash,Unsettled cash,Dif cash,Total cash,B. DIF"
        #    elif Sec_type=='FUTRE STCK INDX':
        #        return str("B. Futures Exposure,"+"Index Future,"+str(ins_code[0:4]+ind_fut[0])+",Futures Exposure,A. INDEX FUTURES")
        #
        #
        ##    elif Sec_type=='FUTURE : EQUITY' and ins_code in(ssf) :
        #    elif Sec_type=='FUTURE : EQUITY' and ins_code[3:4] in(ssf):
        ##        return str("Futures Exposure,"+"SSF,"+str(ssf[0]))
        #        return str("B. Futures Exposure,"+"SSF,null"+",Futures Exposure"+",B. SSF")
        #    elif Sec_type=='EQ : ORDINARY':
        #        return "Equity Exposure,Equity,null,Equity Exposure,EQUITY"
        #    elif Sec_type=='EQ : RIGHTS':
        #        return "Equity Exposure,Equity Rights,null,Equity Exposure,EQUITY"
        #    elif Sec_type=='EQ : FOREIGN':
        #        return "Equity Exposure,Equity Foreign,null,Equity Exposure,EQUITY"
        #    elif ins_code[3:4] in(excp):
        ##        return str("Dividend Exposure,"+"SSF Div,"+str(excp[0]))
        #        return str("Dividend Exposure,"+"SSF Div,null,Dividend Exposure,SSF DIV")
        #    elif Sec_type=="FUND : LOC EQ":
        #        return str("Equity Exposure,"+"Equity Fund,"+str(ins_code)+",Equity Exposure,EQUITY")
        #    else:
        #        return "Other,null,null,Other,OTHER"
        #        
        #        
        #def res_ind(dat,des,ind=['Trade_date','Port_code','AssetType1','AssetType2','AssetType3','AssetType4','AssetType5','Quantity','EffExposure','MarketValue','FundValue','Close_price']):
        #    dat=dat.reset_index()
        #    dat['AssetType1']=des
        #    dat['AssetType2']='null'
        #    dat['AssetType3']='null'
        #    dat['AssetType4']='null'
        #    dat['AssetType5']='null'
        #    dat=dat[ind]
        #    return dat
                
        """
        Fund, Benchmark, Corporate Action data import
        """
        #newest = max(glob.iglob(dirtoimport_file+'fund_data/*.xls'), key=os.path.getmtime)
        newest = max(glob.iglob(dirtoimport_file+'*.xls'), key=os.path.getmtime)
        newest_cash=max(glob.iglob(dirtoimport_cashfile+'*.xls'), key=os.path.getmtime)
        #str(dirtoimport_file+newest)
        #newest
        
        fund_xls = pd.read_excel(newest,sheet_name=0,converters={'Portfolio':str, 'Price Date': pd.to_datetime, 
        'Inst Type':str, 
        'Inst Name':str,
        'ISIN':str,
        'Instrument':str,
        'Quote Close':float,
        'Qty':float,
        'Market Val':float,
        'Delta':float,	
        'Origin':str},
        )

        fund_xls=fund_xls.drop(['Delta'],axis=1)
        if orders:
            pass
        else:
            fund_xls=fund_xls[fund_xls.Origin=='POSITION']
        fund_xls=fund_xls.drop(['Origin'],axis=1)
        #fund_xls.dtypes
        
        fund_xls.columns = ['Port_code','Price_date','Sec_type','Sec_name','Sec_ISIN','Sec_code','Close_price','Quantity','Market_price']
        fund_xls['Close_price']=pd.to_numeric(fund_xls.Close_price.values, errors='coerce') 
        fund_xls['Quantity']=pd.to_numeric(fund_xls.Quantity.values, errors='coerce') 
        fund_xls['Market_price']=pd.to_numeric(fund_xls.Market_price.values, errors='coerce') 
        
        fund_obj = fund_xls.select_dtypes(['object'])
        fund_xls[fund_obj.columns] = fund_obj.apply(lambda x: x.str.strip())
        
        
        df=fund_xls.copy()
        df=df[(df.Port_code.isin(dic_om_index.keys()))]
              
        df.loc[:,'Benchmark_code']=df.Port_code.map(lambda x:dic_om_index[x][1])
        df.loc[:,'TypeFund']=df.Port_code.map(lambda x:dic_om_index[x][2])
        
        df['Trade_date']=startDate
        df['AssetType1']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[0],axis=1)
        df['AssetType2']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[1],axis=1)
        df['AssetType3']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[2],axis=1)
        df['AssetType4']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[3],axis=1)
        df['AssetType5']=df.apply(lambda r: (assetClass(r.Sec_type,r.Sec_code, r.Sec_name,r.TypeFund,cash_flows_eff)).split(",")[4],axis=1)
        df['MarketValue']= np.where(df[['AssetType1']].isin(['B. Futures Exposure','Dividend Exposure']),0, df[['Market_price']])
        df['EffExposure']= df[['Market_price']]
        
        df['Close_price']=np.where((df['AssetType2'].isin(['Index Future']))&(df['Quantity'].values!=0),
                                                              (df['Market_price'].values/df['Quantity'].values)/10, 
                                                               df['Close_price'].values)
        # Futures insert
        if ~fut_flow.empty:
            fut_flow = pd.merge(fut_flow[['Port_code','Sec_code']], (df[['Trade_date','AssetType1','AssetType5','AssetType3','Sec_code','Sec_type','Close_price']]).drop_duplicates(['Sec_code']), on=['Sec_code'], how='left').fillna(0)
            fut_flow['Quantity']=0
            fut_flow['MarketValue']=0
            fut_flow['EffExposure']=0
            fut_flow['Trade_date']=startDate
            fut_flow['AssetType1']='B. Futures Exposure'
            fut_flow['AssetType5']='A. INDEX FUTURES'
            fut_flow['AssetType3']=fut_flow.Sec_code
        
            fut_flow=(fut_flow[['Trade_date','Port_code','AssetType1','AssetType5','AssetType3','Sec_code','Sec_type','Close_price','Quantity','MarketValue','EffExposure']]).copy()
            fut_flow=fut_flow[~(fut_flow.Sec_code=='NoFuture')]
           
        df=df[df.Port_code.isin(lst_fund)]
                                   
        dfprt=(df[['Trade_date','Port_code','AssetType1','AssetType5','AssetType3','Sec_code', 'Sec_type','Close_price','Quantity','MarketValue','EffExposure']]).copy()
        
        # Remove cash and other non-equity asset classes for multi-asset class
        dfprt['MarketValue']=np.where((dfprt.Port_code.map(lambda x:dic_om_index[x][2])=='M')&(dfprt.AssetType1!='Equity Exposure'), 0, dfprt.MarketValue.values)
        dfprt['EffExposure']=np.where((dfprt.Port_code.map(lambda x:dic_om_index[x][2])=='M')&(dfprt.AssetType1!='Equity Exposure'), 0, dfprt.EffExposure.values)
        dfprt= dfprt[~((dfprt.AssetType1=='Other')&(dfprt.Port_code.map(lambda x:dic_om_index[x][2])=='M'))]
        
        #Remove SSF Dividend Exposure
        dfprt.loc[:,'EffExposure']=np.where(dfprt[['AssetType5']].isin(override),0, dfprt[['EffExposure']])
        dfprt=dfprt[~(dfprt.Port_code.isnull())]
        dfprt=dfprt[~(dfprt.Quantity.isnull())]
        dfprt_preflow=dfprt.copy()
        
        # Add futures structureback               
            
        dfprt_preflow=dfprt_preflow.append(fut_flow,sort=True)  
          
        
                    
        if ~cash_flows_eff.empty:
            xx=cfvf(cash_flows_eff, newest_cash,startDate, lst_fund,bf=0.01)
            cash_flows_eff=cash_flows_eff.merge((xx[1])[['Port_code','Inflow_use']], on=['Port_code'], how='left')
            cash_flows_eff=cash_flows_eff[['Port_code', 'Inflow_use', 'Eff_cash', 'fut_sufx']]
            cash_flows_eff.columns=['Port_code', 'Inflow', 'Eff_cash', 'fut_sufx']
            cash_flows_eff['Trade_date']=startDate
            cash_flows_eff['AssetType1']='A. Total cash'
            cash_flows_eff['AssetType2']='Settled cash'
            cash_flows_eff['AssetType3']='Cash flow'
            cash_flows_eff['AssetType4']='Cash flow'
            cash_flows_eff['AssetType5']='Cash flow'
            cash_flows_eff['Sec_code']='ZAR'
            cash_flows_eff['Sec_type']='VAL'
            cash_flows_eff['Close_price']=1
            cash_flows_eff['Quantity']= cash_flows_eff[['Inflow']]
            cash_flows_eff['MarketValue']= cash_flows_eff[['Inflow']]
            cash_flows_eff['EffExposure']= cash_flows_eff[['Inflow']]
            
            cash_flows=cash_flows_eff[['Trade_date', 'Port_code','AssetType1', 'AssetType5', 'AssetType3', 'Sec_code','Sec_type','Close_price', 'Quantity','MarketValue','EffExposure']]
            
            chx_flw=(xx[1])
            chx_flw=chx_flw[chx_flw.Port_code.isin(cash_flows_act_name.tolist())]
        else:
            chx_flw=pd.DataFrame(columns=['Port_code','Inflow_use','ActFlow'])
        
        dfprt1=dfprt_preflow.append(cash_flows, sort=True)  
        dfprt=dfprt1.drop(['Sec_type'],axis=1)
        
        
        
        
        
        '''Generate cash calc
        - Consolidate the current holdings of each fund 
        1) Calculate the cash holdings
        2) Calculate the futures exposure (both Index and SSF)
        3) Calculate the equity market value
        4) Calculate the special fund (cash and equity exposure)
        
        '''
             
        
        #def fx_dta(dfprt_x=dfprt, res_ind):
        #    dfprt_1=dfprt_x.groupby(['Trade_date','Port_code','AssetType1','AssetType5','AssetType3']).agg({'EffExposure':'sum','MarketValue':'sum','Quantity':'sum','Close_price':'max'})
        #    dfprt_1=dfprt_1.reset_index()
        #    dfprt_2= (dfprt_1.groupby(['Trade_date','Port_code']).agg({'MarketValue':'sum'})).reset_index()
        #    dfprt_1=pd.merge( dfprt_1,dfprt_2, on=['Trade_date','Port_code'])
        #    dfprt_1.rename(columns={'MarketValue_x':'MarketValue', 'MarketValue_y':'FundValue'}, inplace=True)
        #    dfprt_1=dfprt_1[['Trade_date','Port_code','AssetType1','AssetType5','AssetType3','MarketValue','EffExposure','Quantity','FundValue','Close_price']]
        #    dfprt_1=dfprt_1.groupby(['Trade_date','Port_code','AssetType1','AssetType5','AssetType3']).agg({'EffExposure':'sum','MarketValue':'sum','FundValue':'max','Quantity':'max','Close_price':'max'})
        #    
        #    req_sum={'EffExposure':'sum','MarketValue':'sum','FundValue':'max','Quantity':'max','Close_price':'max'}
        #    total_cash= (dfprt_1[(dfprt_1.index.get_level_values('AssetType1').isin(['A. Total cash']))]).reset_index().groupby(['Trade_date','Port_code']).agg(req_sum)
        #    
        #    effective_cash=((total_cash-(dfprt_1[(dfprt_1.index.get_level_values('AssetType1').isin(['B. Futures Exposure']))]).reset_index().groupby(['Trade_date','Port_code']).agg(req_sum)).fillna(0))
        #    effective_cash['MarketValue']=0
        #    effective_cash['FundValue']=total_cash[['FundValue']].values
        #    effective_cash['EffExposure']=np.where(effective_cash[['EffExposure']].values==0,total_cash[['EffExposure']].values, effective_cash[['EffExposure']].values)
        #    
        #    
        #    cash_dat=res_ind(effective_cash,'Effective cash').reset_index()
        #    cash_dat['Trade_date']=startDate
        #    cash_dat=(cash_dat[['Trade_date', 'Port_code','AssetType1','AssetType5','AssetType3', 'Quantity','EffExposure','MarketValue','FundValue','Close_price']])
        #    new_dat=((pd.concat([dfprt_1.reset_index(),cash_dat],axis=0, sort=True).reset_index().drop('index',axis=1)).sort_values(['Port_code','AssetType1','AssetType5','AssetType3'])).set_index(['Trade_date','Port_code','AssetType1','AssetType5','AssetType3'])
        #    new_dat['EffWgt']=new_dat[['EffExposure']].values/new_dat[['FundValue']].values
        #    new_dat['MktWgt']=new_dat[['MarketValue']].values/new_dat[['FundValue']].values
        #    n_1 = new_dat.reset_index()
        #    n_1=n_1.groupby(['Port_code','AssetType1']).agg({'EffExposure':'sum','EffWgt':'sum'})
        #    n_1=n_1[~(n_1.index.get_level_values('AssetType1').isin(['Dividend Exposure']))]
        #    n_2=n_1.reset_index()
        #    fnd_value=(total_cash[['FundValue']].reset_index().set_index('Port_code')[['FundValue']]).reset_index()
        #    fnd_value['AssetType1']='Fund Value'
        #    fnd_value['EffWgt']=1
        #    fnd_value.columns= ['Port_code','EffExposure','AssetType1','EffWgt']  
        #    fnd_value=fnd_value[n_2.columns]
        #    n_3=n_2.append(fnd_value)
        #    n_3=n_3.reset_index().pivot(index='Port_code', columns='AssetType1',values='EffExposure')
        #    n_4=n_2.reset_index().pivot(index='Port_code', columns='AssetType1',values='EffWgt')
        #    
        #    n_3.columns=[sym.replace(" ", "")+'_R'  for sym in n_3.columns]
        #    n_4.columns=[sym.replace(" ", "")+'_p'  for sym in n_4.columns]
        #    
        #    n_comb=n_3.merge(n_4, left_index=True, right_index=True)   
        #    n_comb[['B.FuturesExposure_R']]=(n_comb[['B.FuturesExposure_R']]).fillna(0)   
        #    n_comb[['B.FuturesExposure_p']]=(n_comb[['B.FuturesExposure_p']]).fillna(0)   
        #    lst = [new_dat, n_comb]
        #    return lst
        
        # Pre flow
        new_dat_preflow=fx_dta(dfprt_preflow,startDate)
        new_dat_pf=new_dat_preflow[0]
        n_comb_pf=new_dat_preflow[1]
        
        time_elapsed = datetime.now() - start_time 
        
        print('Time elapsed (hh:mm:ss.ms) {}'.format(time_elapsed))
        
        # Post flow
        new_dat_x=fx_dta(dfprt, startDate)
        new_dat=new_dat_x[0]
        n_comb=new_dat_x[1]
                
        # Determing the number of futures
         #fut_price=((new_dat[new_dat.index.get_level_values('AssetType2')=='Index Future']['Close_price']).reset_index())[['Port_code','AssetType3','Close_price']]        
        no_fut=((new_dat[new_dat.index.get_level_values('AssetType5')=='A. INDEX FUTURES'][['Quantity']]).reset_index())[['Port_code','AssetType5','Quantity']]        
        fut_code1=(cash_lmt_x[['P_Code', 'Future_Code']]).copy()
        fut_code1.loc[:,'Future_Code']= np.where(fut_code1['Future_Code']=='NoFuture', 'NoFuture', fut_code1['Future_Code']+str((cash_flows_eff['fut_sufx'].values)[0]))
        no_fut = no_fut.merge(fut_code1, how='right',left_on=['Port_code'],right_on=['P_Code'] )
        no_fut['AssetType5'] = np.where(no_fut.AssetType5.isnull(), no_fut.Future_Code.values, no_fut.AssetType5.values)
        no_fut=(no_fut[['P_Code','AssetType5','Quantity']]).merge((((new_dat[new_dat.index.get_level_values('AssetType5')=='A. INDEX FUTURES'][['Close_price']]).reset_index())[['Port_code','AssetType5','Close_price']]).drop_duplicates(['Port_code','AssetType5']) ,
                               how='left',left_on=['P_Code','AssetType5'],right_on=['Port_code','AssetType5'])
        no_fut=no_fut.fillna(0)
        no_fut=no_fut[['Port_code', 'AssetType5',  'Quantity',  'Close_price']]
        #no_fut.columns= ['Port_code', 'AssetType5',  'Quantity',  'Close_price']
        
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
        n_comb.loc[:,'Inflow_p']= np.where(n_comb.Inflow.isnull().values,0 ,  n_comb.Inflow.values/n_comb.FundValue_R.values ) 
        n_comb.loc[:,'Mid_Totalcash_p'] = n_comb[['Max_TotalCash', 'Min_TotalCash']].mean(axis=1) # placeholder
        
            
        "' Create Breach Flag '"
        #   Inflow 	
        #  Var1	    test1	                 Var2	      test2		                Action
        #	Eff cash	up (breach upper bound)	Total cash	   up (breach upper bound)	  Trade equity + futures
        #	Eff cash	up (breach upper bound)	Total cash	   within bounds		         Trade futures only
        #	Eff cash	up (breach upper bound)	Total cash	   down (breach lower bound)  Trade equity + futures
        						
        #	Eff cash	within bounds	          Total cash	   up (breach upper bound)		Trade equity
        #	Eff cash	within bounds	          Total cash	   within bounds		          No action
        #	Eff cash	within bounds	          Total cash	   down (breach lower bound)	Trade equity
        						
        #Outflow						
        #  Var1	    test1	                     Var2	       test2		                      Action
        #	Eff cash	down (breach lower bound)	Total cash 	down (breach lower bound)		Trade equity + futures
        #	Eff cash	down (breach lower bound)	Total cash	   within bounds		              Trade futures only
        #	Eff cash	down (breach lower bound)	Total cash	   down (breach upper bound)		Trade equity + futures
        						
        #	Eff cash	within bounds	             Total cash	   down (breach lower bound)		Trade equity
        #	Eff cash	within bounds	             Total cash	   within bounds		              No action
        #	Eff cash	within bounds              Total cash	   down (breach upper bound)		Trade equity
        						
        #Else						Override
        n_comb.columns =['Port_code', 'Totalcash_R', 'FuturesExposure_R', 'Effectivecash_R',
                           'EquityExposure_R', 'FundValue_R', 'Totalcash_p',
                           'FuturesExposure_p', 'Effectivecash_p', 'EquityExposure_p',
                           'AssetType5', 'Quantity', 'Close_price', 'P_Code', 'Future_Code_x',
                           'Min_EffCash', 'Max_EffCash', 'Min_TotalCash', 'Max_TotalCash',
                           'Tgt_EffCash', 'Tgt_TotalCash', 'Ovd_Effcash', 'Inflow', 'Tgt_EffCash1',
                           'Future_Code_y', 'FundValue_p','Inflow_p','Mid_Totalcash_p']   
                     
        n_comb=n_comb.merge((xx[1])[['Port_code','ActFlow']], how='left', on=['Port_code'])
        n_comb['ActFlow_p']=n_comb.ActFlow/n_comb.FundValue_R
        
        n_comb['CashFlowFlag']=n_comb.apply(lambda r: (cff(r.Effectivecash_p,r.Totalcash_p, r.Max_TotalCash, r.Min_TotalCash, r.Max_EffCash, r.Min_EffCash, r.Mid_Totalcash_p, r.ActFlow_p,r.FuturesExposure_p)),axis=1)
               
         
                
        n_comb['fin_teff_cash']=n_comb.apply(lambda r: (t_c(r.CashFlowFlag, r.Tgt_EffCash1,r.Tgt_TotalCash, r.Future_Code_y,r.Max_EffCash, r.Min_EffCash,r.Ovd_Effcash,
                                                                   r.Effectivecash_p, r.Totalcash_p, r.FundValue_R, r.Close_price, r.FuturesExposure_p,r.ActFlow_p)[0]),axis=1)
        n_comb['fin_tot_cash']=n_comb.apply(lambda r: (t_c(r.CashFlowFlag, r.Tgt_EffCash1,r.Tgt_TotalCash, r.Future_Code_y,r.Max_EffCash, r.Min_EffCash,r.Ovd_Effcash,
                                                                   r.Effectivecash_p, r.Totalcash_p, r.FundValue_R, r.Close_price, r.FuturesExposure_p, r.ActFlow_p)[1]),axis=1)
        n_comb['InvType'] = np.where(n_comb.Inflow.values > 0, 'Investment', np.where(n_comb.Inflow.values < 0, 'Withdrawal Pay(t)', 'No cash flow'))        
        
        
        if automatic:
            n_comb['trd_fut'] = n_comb.apply(lambda r: (t_c_a(r.Port_code,r.CashFlowFlag, r.Tgt_EffCash1,r.Tgt_TotalCash, r.Future_Code_y,r.Max_EffCash, r.Min_EffCash,r.Ovd_Effcash,
                                                                   r.Effectivecash_p, r.Totalcash_p, r.FundValue_R, r.Close_price, r.FuturesExposure_p, r.ActFlow_p, r.Quantity)[2]),axis=1).astype(float)
 
            n_comb['tot_fut']=n_comb.apply(lambda r: (t_c_a(r.Port_code,r.CashFlowFlag, r.Tgt_EffCash1,r.Tgt_TotalCash, r.Future_Code_y,r.Max_EffCash, r.Min_EffCash,r.Ovd_Effcash,
                                                                   r.Effectivecash_p, r.Totalcash_p, r.FundValue_R, r.Close_price, r.FuturesExposure_p, r.ActFlow_p, r.Quantity)[3]),axis=1).astype(float)
            n_comb['cash_bpm']=n_comb.apply(lambda r: (t_c_a(r.Port_code,r.CashFlowFlag, r.Tgt_EffCash1,r.Tgt_TotalCash, r.Future_Code_y,r.Max_EffCash, r.Min_EffCash,r.Ovd_Effcash,
                                                                   r.Effectivecash_p, r.Totalcash_p, r.FundValue_R, r.Close_price, r.FuturesExposure_p, r.ActFlow_p, r.Quantity)[4]),axis=1)
            n_comb['eq_trade']=n_comb.apply(lambda r: (t_c_a(r.Port_code,r.CashFlowFlag, r.Tgt_EffCash1,r.Tgt_TotalCash, r.Future_Code_y,r.Max_EffCash, r.Min_EffCash,r.Ovd_Effcash,
                                                                   r.Effectivecash_p, r.Totalcash_p, r.FundValue_R, r.Close_price, r.FuturesExposure_p, r.ActFlow_p, r.Quantity)[5]),axis=1)
         #   n_comb.to_hdf(str('c:/data/n_comb_'+str(startDate.date())+'.hdf'),'w', data_columns=True, format='table')
         #   dfprt.to_hdf(str('c:/data/df_'+ str(startDate.date())+'.hdf'),'w', data_columns=True,format='table')
            
            
        
        run_btch=bcer(startDate,new_dat_pf,new_dat, n_comb,dic_users,dic_om_index, newest, output_folder,fnd_excp,chx_flw,automatic)
       # print("\nReport Complete")
        return run_btch
    else:
       # print("Exit")
        return "Error in \n creating \n batch" 
        
#pre_flow_calcFx("yes")
            
           
              
#pre_flow_calcFx(response='yes',automatic=True,orders=False,testing=False)