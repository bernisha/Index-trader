# -*- coding: utf-8 -*-
"""
Created on Mon Oct  8 15:05:09 2018

@author: Blala
"""

# Write excel functions

import pandas as pd
import numpy as np
import datetime as dt
import os
from datetime import datetime, timedelta
import openpyxl as px
from openpyxl.styles import colors, Font, Border, Side ,Protection
import openpyxl.cell
from openpyxl import load_workbook


start_time = datetime.now() 

output_folder= 'c:/data/'


output_file = output_folder+'\\BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsx'
st_row = 19
st_it = st_row+1

st_col= 'C'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

#hdr = new_dat_pf.columns.tolist()


new = new_dat_pf.reset_index()
new= new[~new.AssetType5.isin(['SSF DIV'])]

new_flow = n_comb
inv=(new_flow[['Port_code','fin_teff_cash','fin_tot_cash', 'InvType']]).set_index('Port_code').T.to_dict('list')
    
#new.pivot(index=['Port_code','AssetType1'], columns= 'AssetType3', values='EffExposure')

#h=pd.pivot_table(new,  index=['Port_code'],columns=['AssetType1','AssetType3'], values=['EffExposure','EffWgt'],aggfunc=np.sum)


qf=pd.DataFrame([], dtype='object')
futures_dict = dict()
cashflow_dict = dict()
fnd_excp = ['DSALPC']

for fnd in lst_fund:
    fnd_sel= str(fnd+' Total')
    new_prt= new[new.Port_code.isin([fnd])]
    new_frt= new[~new.AssetType5.isin(new_prt.AssetType5.unique().tolist())]
    new_frt1=new_frt.copy()
    new_frt1.loc[:,'Port_code'] = fnd
    new_frt1.loc[:,'Close_price':'MktWgt'] = np.nan
    new_prt=new_prt.append(new_frt1)
 
    h=pd.pivot_table(new_prt,  index=['Port_code','AssetType1','AssetType5'],values=['EffExposure','EffWgt'], aggfunc=np.sum)
    g=pd.pivot_table(new_prt,  index=['Port_code','AssetType1','AssetType5'],values=['Quantity','Close_price'], aggfunc=np.sum)
    g=g[g.index.get_level_values(2).isin(['A. INDEX FUTURES'])]
    g.columns = ['EffExposure','EffWgt']
    g = g[['EffWgt','EffExposure']]
    g = g.reset_index()
    g['AssetType5'] = 'A. No. Futures'
    g.columns = ['Port_code', 'AssetType1', 'AssetType5', 'EffExposure', 'EffWgt']
    g=g.set_index(['Port_code','AssetType1','AssetType5'])
    
    # Create futures dictionary
    p=(new_prt[new_prt.AssetType5.isin(['A. INDEX FUTURES'])][['Port_code','AssetType3']]).drop_duplicates('Port_code').set_index('Port_code').T.to_dict('list')
    
    if fnd in fnd_excp:
        p[fnd] = ['No Future']
    
    # Cash flow dicitionary
    c=(((new_dat[(new_dat.index.get_level_values('AssetType5').isin(['Cash flow']))&(new_dat.index.get_level_values('Port_code').isin([fnd]))]['MarketValue']).reset_index())[['Port_code','MarketValue']]).set_index('Port_code').T.to_dict('list')
    
    # Inv Type dictionary
    
    
    futures_dict.update(p)
    cashflow_dict.update(c)
    
    
#h.query('Port_code==["OMSI01"]')
#h.stack()
#h= h.T

    df1 = h.groupby(level=[0,1]).sum()

    df1.index = pd.MultiIndex.from_arrays([df1.index.get_level_values(0), 
                                           df1.index.get_level_values(1)+ '', 
                                           len(df1.index) * ['']])
    df2 = h.groupby(level=0).sum() 
    df2['EffExposure'] = df2.EffExposure.values-(df1[(df1.index.get_level_values(1).isin(['B. Futures Exposure']))].reset_index()).EffExposure.values
    df2['EffWgt'] = 1
    
    df2.index = pd.MultiIndex.from_arrays([df2.index.values + ' Total',
                                           len(df2.index) * [''], 
                                           len(df2.index) * ['']])
    df2_1=df2.reset_index()
    df2_1.loc[df2_1.level_0.isin([fnd_sel]),'EffExposure'] = new_prt.FundValue.head(1).values
    df2_1.loc[df2_1.level_0.isin([fnd_sel]),'EffWgt'] = 1
    df2 = df2_1.set_index(['level_0','level_1','level_2'])

#df2.index.get_level_values('CORPEQ Total').isin(['CORPEQ'])
#df = pd.concat([h, df1, df2]).sort_index(level=[0])
    df = pd.concat([df1,df2,h,g]).sort_index(level=[1])
    df=df.reset_index(level=0, drop=True)
    d = dict(zip(df.columns, ['EffExp', 'EffWgt']))
    df=df.rename(columns=d,level=0)
   # df['EffExposure']=pd.Series(["R{: ,.2f}".format(val) for val in df['EffExposure']], index = df.index)
   # df['EffWgt']=pd.Series(["{0:.4f}%".format(val * 100) for val in df['EffWgt']], index = df.index)
#wf =pd.concat([df1,df2,h]).sort_index(level=[1])
#df = pd.concat([h, df2]).sort_index(level=[0])

    qf= pd.concat([qf,df], axis=1)
    

qf=qf.rename(index={'':'Fund Value','A. Total cash':'Total cash','B. Futures Exposure':'Futures Exposure'}, level=0)    
qf=qf.rename(index={'A. VAL':'VAL', 
                    'B. DIF':'DIF',
                    'A. INDEX FUTURES':'INDEX FUTURES',
                    'A. No. Futures':'No. Futures',
                    'B. SSF':'SSF',
                    'C. CALL': 'CALL',
                    'D. SAFEX':'SAFEX'}, level=1)    

    
#h.to_excel(writer, sheet_name='Sheet1', startrow=st_row, header=  hdr,index_label = ['Portfolio Code','Fund', 'Level 1','Level 2', 'Level 3'])

    
    
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
qf.to_excel(writer, sheet_name='Sheet1', startrow=st_row, startcol =0, header = False)
#wf.to_excel(writer, sheet_name='Sheet1', startrow=st_row, startcol =5)


workbook  = writer.book
worksheet = writer.sheets['Sheet1']




cell_format1 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':12})
cell_format2 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':11})
cell_format2_1 = workbook.add_format({'bold': False, 'font_color': 'black', 'font':11})
cell_format2_2 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':11,  'border':1})
cell_format2_2.set_text_wrap() 
cell_format2_2.set_font_name('Calibri')
cell_format2_3 = workbook.add_format({'bold': False, 'font_color': 'black', 'font':11,'bg_color':'#CCFFFF'})
cell_format2_3.set_font_size(10)
cell_format2_4 = workbook.add_format({'bold': True, 'font_color': 'white', 'font':11,'bg_color':'#003366'})


cell_format3 = workbook.add_format({'bold': True, 'bg_color':'#CCFFFF', 'font':11, 'locked':False })
cell_format4 = workbook.add_format({'bold': True, 'bg_color':'#339966', 'font':11, 'locked':False })
cell_format5 = workbook.add_format({'bold': True, 'bg_color':'#C0C0C0', 'font':11, 'locked':False,'align': 'center','border': 1})
cell_format6 = workbook.add_format({'bold': True, 'bg_color':'#CCFFFF', 'font':11, 'locked':False,'align': 'center','border':1})

cell_format5_1 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': 'R#,##0','font_color': 'red'  })
cell_format5_2 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': '0.00%','font_color': 'red'  })

cell_format6_1 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': 'R#,##0', 'font_color': 'red' })
cell_format6_2 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': '0.00%', 'font_color': 'red' })

cell_format5_3 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': 'R#,##0','font_color': 'black'  })
cell_format5_4 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': '0.00%','font_color': 'black'  })

cell_format6_3 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': 'R#,##0', 'font_color': 'black' })
cell_format6_4 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': '0.00%', 'font_color': 'black' })


cell_format7 = workbook.add_format({'bold': True, 'bg_color':'#CCFFCC', 'font':11, 'locked':False,'align': 'center','border': 1})
cell_format8 = workbook.add_format({'bold': True, 'font':11, 'font_color': '#339966', 'locked':False,'align': 'center','border': 1})

td_cell_format_1 = workbook.add_format({'bold': True, 'bg_color':'#FFC7CE', 'font':10, 'locked':False,'font_color': '#9C0006'  })
td_cell_format_2 = workbook.add_format({'bold': True, 'bg_color':'#C6EFCE', 'font':10, 'locked':False,'font_color': '#006100'  })

td_cell_format_3 = workbook.add_format({'bold': True, 'bg_color':'#FF0000', 'font':10, 'locked':False,'font_color': '#9C0006','align': 'center'  })
td_cell_format_4 = workbook.add_format({'bold': True, 'bg_color':'#CCCCFF', 'font':11, 'locked':False,'font_color': '#0000FF','align': 'center'  })
td_cell_format_4.set_font_size(11)
                                      
format1 = workbook.add_format({'num_format': 'R#,##0'})
format2 = workbook.add_format({'num_format': '0.000%'})     


format1_b = workbook.add_format({'bold': True,'num_format': 'R#,##0'})
format2_b = workbook.add_format({'bold': True,'num_format': '0.000%'})                                    

worksheet.write_string('A1', 'Batch Fund Cash Calc',cell_format1)
  
#        worksheet.write('A2', '<portfolio owner>')
worksheet.write('A3', 'Date', cell_format2)
#worksheet.write('B1', fund)
#worksheet.write('B2', Manager)
worksheet.write('B3', datetime.strftime(datetime.today(), "%Y-%m-%d %H:%M:%S"), cell_format2_1)
worksheet.write_string('A4', 'Prepared by',cell_format2)
worksheet.write_string('B4',str(dic_users[os.environ.get("USERNAME").lower()][1]).upper(), cell_format2_1)
worksheet.write_string('A5', 'Authorised by',cell_format2)
worksheet.merge_range('B5:C5','', cell_format3)

worksheet.write_string('A10', 'CASHFLOWS',cell_format1 )
worksheet.write_string('B10', 'Type of Flow',cell_format1 )
worksheet.write_string('B11', 'Amount Flow',cell_format1 )

worksheet.write_string('A13', 'TARGET CASH LEVELS',cell_format1 )
worksheet.write_string('B14', 'Target Cash Exposure',cell_format1 )
worksheet.write_string('B15', 'Target Total Cash',cell_format1 )
worksheet.write_string('B16', 'Futures',cell_format1 )
worksheet.write_string('B17', 'Benchmark',cell_format1 )





#import string  
#alp=list(string.ascii_uppercase) 
alp=[]
def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string
    #print(string)
#print (colnum_string(1))


for a in range(1,(len(lst_fund))*2+5):
#    print(a)
    x_alp=colnum_string(a)
#    print(x_alp)
    alp.append(x_alp)
#    print(alp)
    


st_col=2
fmts=[cell_format5,cell_format6]
fmts2_1=[cell_format5_1,cell_format6_1] #red R
fmts2_2=[cell_format5_3,cell_format6_3] # black R

fmts3_1=[cell_format5_2,cell_format6_2] # red %
fmts3_2=[cell_format5_4,cell_format6_4] # black %

jet=0
worksheet.set_column('A:A', 24)
worksheet.set_column('B:B', 24)

for j in lst_fund:
 #   print(j)
    get_pl=str(alp[st_col]+str(st_row-12)+":"+alp[st_col+1]+str(st_row-12))
    get_pl2=str(alp[st_col]+str(st_row-8))
    get_pl3=str(alp[st_col+1]+str(st_row-8))
    get_pl4=str(alp[st_col]+str(st_row-9))
    get_pl5=str(alp[st_col]+str(st_row-5))
    get_pl6=str(alp[st_col]+str(st_row-4))
    get_pl7=str(alp[st_col+1]+str(st_row-5))
    get_pl8=str(alp[st_col+1]+str(st_row-4))

    worksheet.set_column(str(alp[st_col]+":"+alp[st_col]), 15,format1)
    worksheet.set_column(str(alp[st_col+1]+":"+alp[st_col+1]), 11,format2)
    worksheet.merge_range(get_pl,j, fmts[jet])
    worksheet.write_number(get_pl2, 0,fmts2_2[jet])
    worksheet.write_string(str(alp[st_col]+str(st_row-6)), 'Target',cell_format2)
    worksheet.write_string(str(alp[st_col+1]+str(st_row-6)), 'Post Trade',cell_format2)
    worksheet.write_formula(get_pl3,str('='+get_pl2+'/('+str(alp[st_col]+str(st_row+1))+'+'+get_pl2+')'),fmts[jet])
    worksheet.merge_range(str(alp[st_col]+str(st_row-3)+":"+alp[st_col+1]+str(st_row-3)), futures_dict[j][0] ,cell_format7)
    worksheet.merge_range(str(alp[st_col]+str(st_row-2)+":"+alp[st_col+1]+str(st_row-2)), dic_om_index[j][1] ,cell_format8)
    worksheet.write_number(str(alp[st_col]+str(st_row-8)), cashflow_dict[j][0] ,format1)
    worksheet.write_number(str(alp[st_col]+str(st_row-5)), inv[j][0] ,fmts3_2[jet])
    worksheet.write_number(str(alp[st_col]+str(st_row-4)), inv[j][1] ,fmts3_2[jet])
    
    worksheet.write_string(get_pl4, inv[j][2] ,cell_format1)
    worksheet.data_validation(get_pl4, {'validate': 'list',
                                        'source': ['Investment', 
                                                   'Hedged Withdrawal', 
                                                   'Hedged With Pay(t)',
                                                   'Withdrawal Pay(t)',
                                                   'No cash flow'],
                                        'input_title': 'Enter a cash flow',
                                        'input_message': 'Type & Value'} )
    
     
   # worksheet.write_formula(get_pl5, str('='+str(alp[st_col+1]+str(st_row+26))),fmts3_2[jet])
   # worksheet.write_formula(get_pl6, str('='+str(alp[st_col+1]+str(st_row+17))),fmts3_2[jet])
    
    worksheet.write_formula(get_pl7, str('='+str(alp[st_col+1]+str(st_row+56))),fmts3_2[jet])
    worksheet.write_formula(get_pl8, str('='+str(alp[st_col+1]+str(st_row+47))),fmts3_2[jet])
    
    worksheet.conditional_format(get_pl2, {'type': 'cell','criteria': '<','value': 0,'format': fmts2_1[jet]})                              
    worksheet.conditional_format(get_pl2, {'type': 'cell','criteria': '>=','value': 0,'format': fmts2_2[jet]})   
    worksheet.conditional_format(get_pl3, {'type': 'cell','criteria': '<','value': 0,'format': fmts3_1[jet]})                              
    worksheet.conditional_format(get_pl3, {'type': 'cell','criteria': '>=','value': 0,'format': fmts3_2[jet]})      
    
    worksheet.conditional_format(get_pl4,{'type': 'cell','criteria': 'equal to','value': '"No cash flow"','format': cell_format1})
    worksheet.conditional_format(get_pl4,{'type': 'cell','criteria': '!=','value': '"No cash flow"','format': td_cell_format_4})
    worksheet.conditional_format(get_pl5,{'type': 'cell','criteria': '!=','value': inv[j][0] ,'format': td_cell_format_4})
    worksheet.conditional_format(get_pl6,{'type': 'cell','criteria': '!=','value': inv[j][1] ,'format': td_cell_format_4})
    
    for g in range(st_row, st_row+ 25):
        if g in [st_row+1, st_row+2,st_row+7, st_row+11,st_row+13]:
            
            worksheet.conditional_format(str(alp[st_col]+str(g)),{'type': 'cell','criteria': '!=','value': inv[j][1] ,'format': format1_b})
            worksheet.conditional_format(str(alp[st_col+1]+str(g)),{'type': 'cell','criteria': '!=','value': inv[j][1] ,'format': format2_b})
    
    #worksheet.conditional_format(get_pl4,{'type': 'cell','criteria': 'equal to','value': '"No cash flow"','format': cell_format1})
    
    st_col=st_col+2
    jet=np.where(jet==0,1,0)
   # print(get_pl)

# grouping categories
del g

worksheet.set_row((st_row+2), None, None, {'level': 1})
worksheet.set_row((st_row+3), None, None, {'level': 1})
worksheet.set_row((st_row+4), None, None, {'level': 1})
worksheet.set_row((st_row+5), None, None, {'level': 1})
#worksheet.set_row((st_row+6), None, None, {'level': 1})

worksheet.set_row((st_row+7), None, None, {'level': 1})
worksheet.set_row((st_row+8), None, cell_format2_3,{'level': 1})
worksheet.set_row((st_row+9), None, None, {'level': 1})

worksheet.set_row((st_row+11), None, None, {'level': 1})
worksheet.set_row((st_row+13), None, None, {'level': 1})

worksheet.set_row(st_row-1, None, cell_format2_4)
worksheet.merge_range(str('A'+str(st_row)+':B'+str(st_row)),'Pre Trade', cell_format2_4)

# end grouping
def form_at(st_no=(st_row+len(qf)),_label_='Inflow'):
    in_st=st_no
    in_cell_format1 = workbook.add_format({'bold': True, 'font_color': 'black', 'border':1, 'align': 'center', 'valign': 'top' })
    bn_cell_format1 = workbook.add_format({'bold': True, 'font_color': 'black', 'border':1, 'align': 'center', 'valign': 'top','bg_color':'#FF8080' })
    
    worksheet.set_row(in_st, None, cell_format2_4)
    worksheet.merge_range(str('A'+str(in_st+1)+':B'+str(in_st+1)),_label_, cell_format2_4)
    worksheet.write_string(str('A'+str(in_st+2)), 'Fund Value',in_cell_format1)
    worksheet.merge_range(str('A'+str(in_st+3)+':A'+str(in_st+7)),'Total cash', in_cell_format1)
    worksheet.merge_range(str('A'+str(in_st+8)+':A'+str(in_st+11)),'Futures Exposure', in_cell_format1)
    worksheet.merge_range(str('A'+str(in_st+12)+':A'+str(in_st+13)),'Effective cash', in_cell_format1)
    worksheet.merge_range(str('A'+str(in_st+14)+':A'+str(in_st+15)),'Equity Exposure', in_cell_format1)
    
    worksheet.write_string(str('B'+str(in_st+2)), '',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+3)), '',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+4)), 'VAL',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+5)), 'DIF',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+6)), 'CALL',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+7)), 'SAFEX',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+8)), '',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+9)), 'INDEX FUTURES',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+10)), 'No. Futures',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+11)), 'SSF',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+12)), '',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+13)), 'null',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+14)), '',in_cell_format1)
    worksheet.write_string(str('B'+str(in_st+15)), 'EQUITY',in_cell_format1)
    
    worksheet.set_row((in_st+3), None, None, {'level': 1})
    worksheet.set_row((in_st+4), None, None, {'level': 1})
    worksheet.set_row((in_st+5), None, None, {'level': 1})
    worksheet.set_row((in_st+6), None, None, {'level': 1})
    #worksheet.set_row((st_row+6), None, None, {'level': 1})
    
    worksheet.set_row((in_st+8), None, None, {'level': 1})
    worksheet.set_row((in_st+9), None, cell_format2_3,{'level': 1})
    worksheet.set_row((in_st+10), None,None,  {'level': 1})
    
    worksheet.set_row((in_st+12), None, None, {'level': 1})
    worksheet.set_row((in_st+14), None, None, {'level': 1})
    
    if _label_=='Post Trade':
         worksheet.write_string(str('B'+str(in_st+17)), 'BARRA CASH',bn_cell_format1)
        

form_at()
form_at(st_no=(st_row+2*len(qf)+1),_label_='Trade')
form_at(st_no=(st_row+3*len(qf)+2),_label_='Post Trade')


# Write formula

# Inflow

in_st = (st_row+1*len(qf)+1)
in_col=2

informat1_n = workbook.add_format({'bold': False,'num_format': 'R#,##0'})
informat2_n = workbook.add_format({'bold': False,'num_format': '0.000%'})                                    
informat1_b = workbook.add_format({'bold': True,'num_format': 'R#,##0'})
informat2_b = workbook.add_format({'bold': True,'num_format': '0.000%'})                   

            
for f in lst_fund:
  #  print(f)
    worksheet.write_formula(str(alp[in_col]+str(in_st+1)),str('='+str(alp[in_col]+str(in_st+2))+'+'+str(alp[in_col]+str(in_st+13))),informat1_b)
    worksheet.write_formula(str(alp[in_col]+str(in_st+2)),str('=SUM('+str(alp[in_col]+str(in_st+3))+':'+str(alp[in_col]+str(in_st+6))+')'),informat1_b)
    worksheet.write_formula(str(alp[in_col]+str(in_st+3)),str('='+str(alp[in_col]+str(in_st-12))+'+IF('+str(alp[in_col]+str(in_st-24))+\
                                '="Hedged Withdrawal",0,'+str(alp[in_col]+str(in_st-23))+')'),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+4)),str('='+str(alp[in_col]+str(in_st-11))),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+5)),str('='+str(alp[in_col]+str(in_st-10))),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+6)),str('='+str(alp[in_col]+str(in_st-9))),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+7)),str('=SUM('+str(alp[in_col]+str(in_st+8))+','+str(alp[in_col]+str(in_st+10))+')'),informat1_b)
    worksheet.write_formula(str(alp[in_col]+str(in_st+8)),str('='+str(alp[in_col]+str(in_st-7))),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+10)),str('='+str(alp[in_col]+str(in_st-5))),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+11)),str('='+str(alp[in_col]+str(in_st+12))),informat1_b)
    worksheet.write_formula(str(alp[in_col]+str(in_st+12)),str('='+str(alp[in_col]+str(in_st+2))+'-'+str(alp[in_col]+str(in_st+7))),informat1_n)
    worksheet.write_formula(str(alp[in_col]+str(in_st+13)),str('='+str(alp[in_col]+str(in_st+14))),informat1_b)
    worksheet.write_formula(str(alp[in_col]+str(in_st+14)),str('='+str(alp[in_col]+str(in_st-1))),informat1_n)
    
    for g in range(in_st+1, in_st+qf.shape[0]+1):
        if g in [in_st+1,in_st+2,in_st+7,in_st+11,in_st+13]:
            worksheet.write_formula(str(alp[in_col+1]+str(g)),str('='+str(alp[in_col]+str(g))+'/'+str(alp[in_col]+'$'+str(in_st+1))),informat2_b)
        #    print(str('B='+str(alp[in_col]+str(g))+'/'+str(alp[in_col]+'$'+str(in_st+1))))
        elif g!=(in_st+9):
            worksheet.write_formula(str(alp[in_col+1]+str(g)),str('='+str(alp[in_col]+str(g))+'/'+str(alp[in_col]+'$'+str(in_st+1))),informat2_n)
        #    print(str('N='+str(alp[in_col]+str(g))+'/'+str(alp[in_col]+'$'+str(in_st+1))))
        else:
        #    print(g)
            gg=1
    in_col=in_col+2

del f 
del g
del gg
# Trade

td_st = (st_row+2*len(qf)+2)
td_col=2


for f in lst_fund:
    worksheet.write_formula(str(alp[td_col]+str(td_st+1)),str('='+str(alp[td_col]+str(td_st+8))+'+'+str(alp[td_col]+str(td_st+14))),informat1_b)
    worksheet.write_formula(str(alp[td_col]+str(td_st+2)),str('=SUM('+str(alp[td_col]+str(td_st+3))+':'+str(alp[td_col]+str(td_st+6))+')'),informat1_b)
    worksheet.write_formula(str(alp[td_col]+str(td_st+3)),str('=-'+str(alp[td_col]+str(td_st+6))),informat1_n)
    worksheet.write_formula(str(alp[td_col]+str(td_st+4)),str('='+str(alp[td_col]+str(td_st+8))+'+'+str(alp[td_col]+str(td_st+11))),informat1_n)
    worksheet.write_formula(str(alp[td_col]+str(td_st+6)),\
                            str('=IF(ISNUMBER('+str(alp[td_col]+str(td_st+9))+'*('+str(alp[td_col]+str(td_st-24))+'/'+ \
                                                    str(alp[td_col]+str(td_st-21))+')),'+str(alp[td_col]+str(td_st+9))+ \
                                                    '*('+str(alp[td_col]+str(td_st-24))+'/'+str(alp[td_col]+str(td_st-21))+'),0)'), informat1_n)
     
    worksheet.write_formula(str(alp[td_col]+str(td_st+7)),str('='+str(alp[td_col]+str(td_st+8))),informat1_b)
    worksheet.write_formula(str(alp[td_col]+str(td_st+8)),\
                            str('=IF(ISERROR('+str(alp[td_col]+str(td_st+9))+'*'+str(alp[td_col+1]+str(td_st-21))+'*10),0,'+\
                                                   str(alp[td_col]+str(td_st+9))+'*'+str(alp[td_col+1]+str(td_st-21))+'*10)'), informat1_n)
    worksheet.write_formula(str(alp[td_col]+str(td_st+9)),\
                            str('=IF('+str(alp[td_col]+str(td_st-33))+'="No Future",0,ROUNDDOWN((('+str(alp[td_col]+str(td_st-34))+'-'+str(alp[td_col]+str(td_st-35))+\
                                '-'+str(alp[td_col+1]+str(td_st-8))+')*'+str(alp[td_col]+str(td_st-14))+')/'+ \
                                    str(alp[td_col+1]+str(td_st-21))+'/10,0))'),cell_format2_1)
    worksheet.write_formula(str(alp[td_col+1]+str(td_st+9)),str('=IF('+str(alp[td_col]+str(td_st+9))+'<0,"SOC",IF('+str(alp[td_col]+str(td_st+9))\
                                                            +'>0,"BOC",""))'),informat1_n) 
    
    worksheet.write_formula(str(alp[td_col]+str(td_st+11)),str('='+str(alp[td_col]+str(td_st+12))),informat1_n)
    worksheet.write_formula(str(alp[td_col]+str(td_st+12)),str('=('+str(alp[td_col]+str(td_st-14))+'*'+str(alp[td_col]+str(td_st-35))+')-'\
                                +str(alp[td_col]+str(td_st-4))),informat1_n)
    worksheet.write_formula(str(alp[td_col]+str(td_st+13)),str('='+str(alp[td_col]+str(td_st+14))),informat1_n)
    worksheet.write_formula(str(alp[td_col]+str(td_st+14)),str('=-'+str(alp[td_col]+str(td_st+4))),informat1_n)
    worksheet.write_formula(str(alp[td_col+1]+str(td_st+14)),str('=IF('+str(alp[td_col]+str(td_st+14))+'<-0.01,"SELL EQTY",IF('+str(alp[td_col]+str(td_st+14))\
                                                            +'>0.01,"BUY EQTY",""))'),informat1_n) 
    
    for g in range(td_st+1, td_st+qf.shape[0]+1):
      #  print(g)
        if g in [td_st+1,td_st+2,td_st+7,td_st+11,td_st+13]:
            worksheet.write_formula(str(alp[td_col+1]+str(g)),str('='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st-14))),informat2_b)
     #       print(str('B='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st+1))))
        elif g in [td_st+3,td_st+4,td_st+6,td_st+8,td_st+12]:
            worksheet.write_formula(str(alp[td_col+1]+str(g)),str('='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st-14))),informat2_n)
     #       print(str('N='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st+1))))
        else:
            gg=1
        
        worksheet.conditional_format(str(alp[td_col]+str(td_st+9)), {'type': 'cell','criteria': '<','value': 0,'format': td_cell_format_1})    
        worksheet.conditional_format(str(alp[td_col]+str(td_st+9)), {'type': 'cell','criteria': '>','value': 0,'format': td_cell_format_2})
        worksheet.conditional_format(str(alp[td_col+1]+str(td_st+9)),{'type': 'cell','criteria': '=','value': '"SOC"','format': td_cell_format_1}) 
        worksheet.conditional_format(str(alp[td_col+1]+str(td_st+9)),{'type': 'cell','criteria': '=','value': '"BOC"','format': td_cell_format_2}) 
        worksheet.conditional_format(str(alp[td_col+1]+str(td_st+14)),{'type': 'cell','criteria': '=','value': '"SELL EQTY"','format': td_cell_format_1}) 
        worksheet.conditional_format(str(alp[td_col+1]+str(td_st+14)),{'type': 'cell','criteria': '=','value': '"BUY EQTY"','format': td_cell_format_2}) 
        worksheet.conditional_format(str(alp[td_col]+str(td_st+14)),{'type': 'cell','criteria': '<','value': -0.01,'format': td_cell_format_1})
        worksheet.conditional_format(str(alp[td_col]+str(td_st+14)),{'type': 'cell','criteria': '>','value': 0.01,'format': td_cell_format_2}) 
    
    td_col=td_col+2
                             
del f 
del g
del gg

# Post trade

pt_st = (st_row+3*len(qf)+3)
pt_col=2
pt_cell_format_1 = workbook.add_format({'bold': True, 'bg_color':'#FFC7CE', 'font':10, 'locked':False,'font_color': '#9C0006','num_format': '0.000%' })
pt_cell_format_2 = workbook.add_format({'bold': True, 'bg_color':'#C6EFCE', 'font':10, 'locked':False,'font_color': '#006100','num_format': '0.000%'  })

for f in lst_fund:
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+1)),str('='+str(alp[pt_col]+str(pt_st+2))+'+'+str(alp[pt_col]+str(pt_st+13))),informat1_b)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+2)),str('=SUM('+str(alp[pt_col]+str(pt_st+3))+':'+str(alp[pt_col]+str(pt_st+6))+')'),informat1_b)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+3)),str('='+str(alp[pt_col]+str(pt_st-27))+'+'+str(alp[pt_col]+str(pt_st-12))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+4)),str('='+str(alp[pt_col]+str(pt_st-26))+'+'+str(alp[pt_col]+str(pt_st-11))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+5)),str('='+str(alp[pt_col]+str(pt_st-25))+'+'+str(alp[pt_col]+str(pt_st-10))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+6)),str('='+str(alp[pt_col]+str(pt_st-24))+'+'+str(alp[pt_col]+str(pt_st-9))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+7)),str('='+str(alp[pt_col]+str(pt_st-23))+'+'+str(alp[pt_col]+str(pt_st-8))),informat1_b)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+8)),str('='+str(alp[pt_col]+str(pt_st-22))+'+'+str(alp[pt_col]+str(pt_st-7))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+9)),str('='+str(alp[pt_col]+str(pt_st-36))+'+'+str(alp[pt_col]+str(pt_st-6))),cell_format2_3)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+10)),str('='+str(alp[pt_col]+str(pt_st-35))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+11)),str('='+str(alp[pt_col]+str(pt_st+12))),informat1_b)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+12)),str('='+str(alp[pt_col]+str(pt_st-18))+'+'+str(alp[pt_col]+str(pt_st-3))),informat1_n)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+13)),str('='+str(alp[pt_col]+str(pt_st+14))),informat1_b)
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+14)),str('='+str(alp[pt_col]+str(pt_st-16))+'+'+str(alp[pt_col]+str(pt_st-1))),informat1_n)
     
     worksheet.write_formula(str(alp[pt_col]+str(pt_st+16)),str('='+str(alp[pt_col]+str(pt_st+11))+'-'+str(alp[pt_col]+str(pt_st-13))),informat1_n)
     
     
     for g in range(pt_st+1, pt_st+qf.shape[0]+3):
     #   print(g)
        if g in [pt_st+1,pt_st+2,pt_st+7,pt_st+11,pt_st+13,pt_st+16]:
            worksheet.write_formula(str(alp[pt_col+1]+str(g)),str('='+str(alp[pt_col]+str(g))+'/'+str(alp[pt_col]+'$'+str(pt_st+1))),informat2_b)
     #       print(str('B='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st+1))))
        elif g in [pt_st+3,pt_st+4,pt_st+5,pt_st+6,pt_st+8,pt_st+12,pt_st+14]:
            worksheet.write_formula(str(alp[pt_col+1]+str(g)),str('='+str(alp[pt_col]+str(g))+'/'+str(alp[pt_col]+'$'+str(pt_st+1))),informat2_n)
     #       print(str('N='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st+1))))
        else:
            gg=1
   
     worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+11)), {'type': 'cell','criteria': '<','value': 0,'format': pt_cell_format_1})    
     worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+11)), {'type': 'cell','criteria': '>','value': 0,'format': pt_cell_format_2})
       
    
     pt_col=pt_col+2
                       
del f 
del g
del gg

worksheet.freeze_panes(19,2)
worksheet.write_comment('C10', 'Enter cash flow info below', {'start_col': 5,'start_row': 7, 'x_scale': 1.2, 'y_scale': 0.25, 'visible': True ,'font_size': 11, 'bold':True ,'color': '#FFCC99'})

writer.save()
workbook.close()

time_elapsed = datetime.now() - start_time 
print('Time elapsed (hh:mm:ss.ms) {}'.format(time_elapsed))


"' ################################################################################################################# '"
wb = load_workbook(output_file)
ws = wb.active
ws.delete_cols(2)
wb.save('c:\\data\\test_file.xlsx')
    # or by name: ws = wb['SheetName']
col = 'A'
    del_col(ws, col)
    wb.save('output_book.xlsx')

 delete_column = openpyxl.cell.column_index_from_string(col)