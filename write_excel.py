# -*- coding: utf-8 -*-
"""
Created on Thu Mar 29 15:02:40 2018

@author: blala
"""
"""
'******************************************************************************************************************************************************************************    
'                                                      Create a Pandas Excel writer using XlsxWriter as the engine
'                                                       Futures report with required number of futures trades
'******************************************************************************************************************************************************************************    
"""


def excel_fx(output_folder,dic_users,n_comb_eff_1,startDate):
    
    import pandas as pd
    import numpy as np
    import datetime as dt
    import os
    from datetime import datetime, timedelta
    import openpyxl as px
    from openpyxl.styles import colors, Font, Border, Side ,Protection
    
    
    output_file = output_folder+'\\IndexFutRep_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsx'
    st_row = 7
    st_it = st_row+1
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    hdr= ['FundValue', 'EquityExposure', 'Totalcash', 'FuturesExposure','Effectivecash', 
          'Cash Flow','Totalcash', 'Effectivecash',
          'Tgt_EffCash', 'No. Futures / Price', 'FutureCode','Trade','FundValue', 'EquityExposure', 'Totalcash',
          'FuturesExposure', 'Effectivecash', 'Check cash', 'TradeSignal','TradeComment','Checked by']        
    n_comb_dta=n_comb_eff_1[['FundValue_R_pf', 'EquityExposure_R_pf', 'Totalcash_R_pf', 'FuturesExposure_R_pf','Effectivecash_R_pf',
                            'Inflow', 'Totalcash_R', 'Effectivecash_R',
                       #     'FundValue_R', 'EquityExposure_R', 'Totalcash_R', 'FuturesExposure_R','Effectivecash_R',
                            'Tgt_EffCash1', 'No. Futures', 'AssetType3','Trade',
                            'FundValue_TR', 'EquityExposure_TR', 'Totalcash_TR','FuturesExposure_TR', 'Effectivecash_TR', 'Check cash','Trade_YN','Comment','Checked']]
    n_comb_dta.to_excel(writer, sheet_name='Sheet1', startrow=st_row, header=  hdr,index_label = ['Portfolio Code',' '])
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # colour source (hex codes): http://dmcritchie.mvps.org/excel/colors.htm
    # Convert the dataframe to an XlsxWriter Excel object.
    cell_format1 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':12})
    cell_format2 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':11})
    cell_format2_1 = workbook.add_format({'bold': False, 'font_color': 'black', 'font':11})
    cell_format2_2 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':11,  'border':1})
    cell_format2_2.set_text_wrap() 
    cell_format2_2.set_font_name('Calibri')
    cell_format3 = workbook.add_format({'bold': True, 'bg_color':'#FFFF99', 'font':11, 'locked':False })
    cell_format4 = workbook.add_format({'bold': True, 'bg_color':'#339966', 'font':11, 'locked':False })
    unlocked = workbook.add_format({'locked': False})
    locked = workbook.add_format({'locked': True})
    
    
    worksheet.write_string('A1', 'Indexation Futures Report',cell_format1)
  
    #        worksheet.write('A2', '<portfolio owner>')
    worksheet.write('A3', 'Date', cell_format2)
    #worksheet.write('B1', fund)
    #worksheet.write('B2', Manager)
    worksheet.write('B3', datetime.strftime(datetime.today(), "%Y-%m-%d %H:%M:%S"), cell_format2_1)
    worksheet.write_string('A4', 'Prepared by',cell_format2)
    worksheet.write_string('B4',str(dic_users[os.environ.get("USERNAME").lower()][1]).upper(), cell_format2_1)
    worksheet.write_string('A5', 'Authorised by',cell_format2)
    worksheet.merge_range('B5:C5','', cell_format3)

    #worksheet.write_string('A7', 'USWIMF',merge_format1)
      
    
    merge_format1 = workbook.add_format({'bold': 1,'border': 0,'align': 'center','valign': 'vcenter','fg_color': '#C6EFCE', 'font_color':'#006100'})
    merge_format2 = workbook.add_format({'bold': 1,'border': 0,'align': 'center','valign': 'vcenter','fg_color': '#CCFFFF', 'font_color':'#003366'}) # blue
    merge_format3 = workbook.add_format({'bold': 1,'border': 0,'align': 'center','valign': 'vcenter','fg_color': '#FF99CC', 'font_color':'#993366'}) # pink
    merge_format4 = workbook.add_format({'bold': 1,'fg_color': '#C6EFCE', 'font_color':'#006100','num_format': '0.000%'})
    merge_format5 = workbook.add_format({'bold': 1,'border': 0,'align': 'center','valign': 'vcenter','fg_color': '#CCCCFF', 'font_color':'#800080'}) # blue
    merge_format6 = workbook.add_format({'bold': 1,'border': 0,'align': 'center','valign': 'vcenter','fg_color': '#FF8080', 'font_color':'#800000'}) # blue
    
    
    #worksheet.merge_range('C5:G5', 'Pre Trade', merge_format1)
    worksheet.merge_range(str('C'+str(st_row)+':'+'G'+str(st_row)), 'Pre Trade', merge_format1)
    worksheet.merge_range(str('H'+str(st_row)+':'+'J'+str(st_row)), 'Flow', merge_format5)
    worksheet.write_string(str('K'+str(st_row)), 'Target', merge_format6)
    worksheet.merge_range(str('L'+str(st_row)+':'+'N'+str(st_row)), 'Trade', merge_format2)
    worksheet.merge_range(str('O'+str(st_row)+':'+'T'+str(st_row)), 'Post Trade', merge_format3)
    worksheet.merge_range(str('U'+str(st_row)+':'+'W'+str(st_row)),  'Sign-off', merge_format2)
    
    worksheet.write(7, 11, "No. Futures / Price", cell_format2_2)
    
    worksheet.set_column('A:A', 13)
    worksheet.set_column('B:B', 4)
    worksheet.set_column('C:C', 17)
    worksheet.set_column('D:D', 17)
    worksheet.set_column('E:E', 15)
    worksheet.set_column('F:F', 15)
    worksheet.set_column('G:G', 13)
        
    worksheet.set_column('H:H', 13)        
    worksheet.set_column('I:I', 12)        
    worksheet.set_column('J:J', 11)
    

    worksheet.set_column('K:K',10)
    worksheet.set_column('L:L', 11)
    worksheet.set_column('M:M', 11)
    worksheet.set_column('N:N', 11)        
    worksheet.set_column('O:O', 15)        
    worksheet.set_column('P:P', 14)        
    worksheet.set_column('Q:Q', 12)
    worksheet.set_column('R:R', 13)
    worksheet.set_column('S:S', 14)
    worksheet.set_column('T:T', 11)        
    worksheet.set_column('U:U', 10)        
    worksheet.set_column('V:V', 13)        
    worksheet.set_column('W:W', 11)
#    worksheet.set_column('F:F', 15, None, {'level': 1})
    # Get the xlsxwriter workbook and worksheet objects.
    #workbook  = writer.book
    #worksheet = writer.sheets['Sheet1']
    
    # Add some cell formats.
    format1 = workbook.add_format({'num_format': 'R#,##0'})
    format2 = workbook.add_format({'num_format': '0.000%'})
    format3 = workbook.add_format({'num_format': '0', 'font_color': 'black', 'bold':True})
    format4 = workbook.add_format({'num_format': '0', 'font_color': 'black', 'bold':False})
    format5 = workbook.add_format({'bold': 1,'bg_color': '#C6EFCE', 'font_color':'#006100'})
    format6 = workbook.add_format({'bold': 1,'bg_color': '#FFC7CE', 'font_color':'#9C0006'})
    format7 = workbook.add_format({'num_format': '0.000%','bold':True,'font_color': 'red' })
    format8 = workbook.add_format({'num_format': '0.000%','bold':True,'font_color': 'green' })
    
    # Note: It isn't possible to format any cells that already have a format such
    # as the index or headers or any cells that contain dates or datetimes.
    
    # Set the column width and format.
    #worksheet.set_column('B:B', 18, format1)
    
    # Set the format but not the column width.
    #worksheet.set_column('C:C', None, format2)
    len1=n_comb_eff_1.shape[0]
    for i in range(st_it,len1+st_it,2):
   #    print(i)
       worksheet.set_row(i, 18, format1)
    
    for j in range(st_it,len1+st_it,2):
 #      print(i+1)
       worksheet.set_row(j+1, 18, format2)
    
      
    worksheet.conditional_format(str('L'+str(st_row+2)+':'+'L100'), {'type': 'cell',
                                                                     'criteria': 'between',
                                                                     'minimum': 1,
                                                                     'maximum': 9999,
                                                                     'format': format3})
    worksheet.conditional_format(str('L'+str(st_row+2)+':'+'L100'), {'type': 'cell',
                                                                     'criteria': '>',
                                                                     'value': 10000,
                                                                     'format': format1})
    worksheet.conditional_format(str('L'+str(st_row+2)+':'+'L100'), {'type': 'cell',
                                                                     'criteria': '=',
                                                                     'value': 0,
                                                                     'format': format4})
    worksheet.conditional_format(str('L'+str(st_row+2)+':'+'L100'), {'type': 'cell',
                                                                     'criteria': '<',
                                                                     'value': 0,
                                                                     'format': format3})
  
    worksheet.conditional_format(str('N'+str(st_row+2)+':'+'N100'), {'type': 'cell',
                                                                     'criteria': '=',
                                                                     'value': '"Buy"',
                                                                     'format': format5})
    worksheet.conditional_format(str('N'+str(st_row+2)+':'+'N100'), {'type': 'cell',
                                                                     'criteria': '=',
                                                                     'value': '"Sell"',
                                                                     'format': format6})
    
           
    for i in range(1, len(n_comb_eff_1),2):
   #     print(i+6)
        if n_comb_eff_1['Effectivecash_R'].iloc[i] < n_comb_eff_1['Min_EffCash'].iloc[i]:
            worksheet.write(i+st_it, 9, n_comb_dta['Effectivecash_R'].iloc[i], format7)
        elif n_comb_eff_1['Effectivecash_R'].iloc[i] > n_comb_eff_1['Max_EffCash'].iloc[i]:
            worksheet.write(i+st_it, 9, n_comb_dta['Effectivecash_R'].iloc[i], format7)
    
        if n_comb_eff_1['Totalcash_R'].iloc[i] < n_comb_eff_1['Min_TotalCash'].iloc[i]:
            worksheet.write(i+st_it, 8, n_comb_dta['Totalcash_R'].iloc[i], format7)
            worksheet.write(i+st_it, 16, n_comb_dta['Totalcash_TR'].iloc[i], format7)
            worksheet.write(i+st_it, 19, n_comb_dta['Check cash'].iloc[i], format6)
        elif n_comb_eff_1['Totalcash_R'].iloc[i] > n_comb_eff_1['Max_TotalCash'].iloc[i]:
            worksheet.write(i+st_it, 8, n_comb_dta['Totalcash_R'].iloc[i], format7)
            worksheet.write(i+st_it, 16, n_comb_dta['Totalcash_TR'].iloc[i], format7)
            worksheet.write(i+st_it, 19, n_comb_dta['Check cash'].iloc[i], format6)
        
        if n_comb_eff_1['Effectivecash_TR'].iloc[i] < n_comb_eff_1['Min_EffCash'].iloc[i]:
            worksheet.write(i+st_it, 18, n_comb_dta['Effectivecash_TR'].iloc[i], format7)
        elif n_comb_eff_1['Effectivecash_TR'].iloc[i] > n_comb_eff_1['Max_EffCash'].iloc[i]:
            worksheet.write(i+st_it, 18, n_comb_dta['Effectivecash_TR'].iloc[i], format7)
        else:
            worksheet.write(i+st_it, 18, n_comb_dta['Effectivecash_TR'].iloc[i], format8)
            
        if (n_comb_eff_1['Inflow'].iloc[i] < 0)&(~np.isnan(n_comb_eff_1['Inflow'].iloc[i])) :
            worksheet.write(i+st_it, 7, n_comb_dta['Inflow'].iloc[i], format7)
        elif (n_comb_eff_1['Inflow'].iloc[i] > 0)&(~np.isnan(n_comb_eff_1['Inflow'].iloc[i])):
            worksheet.write(i+st_it, 7, n_comb_dta['Inflow'].iloc[i], format8)
       
        
        if n_comb_eff_1['Tgt_EffCash1'].iloc[i] != n_comb_eff_1['Tgt_EffCash'].iloc[i]:
            worksheet.write(i+st_it, 10, n_comb_dta['Tgt_EffCash1'].iloc[i], merge_format4)
       
           
    # Close the Pandas Excel writer and output the Excel file.
    worksheet.autofilter(str('A'+str(st_it)+':'+'W100'))
    worksheet.set_row(st_row, 29.25)
#    worksheet.write(st_row, 1, 'Portfolio code', cell_format3)
#    worksheet.protect()
    
  #  worksheet.set_column('R:R', None, unlocked)
    
    for j in range(1, len(n_comb_eff_1),2):
 #       print(n_comb_dta.index.values[j][0])
 #       if n_comb_eff_1['Check cash'].ix[j] == '':
         worksheet.write(j+st_it, 20, '', cell_format3)
         worksheet.write(j+st_it, 21, '', cell_format3)
         worksheet.write(j+st_it, 22, '', cell_format3)
      #   worksheet.write(str('A'+str(j+st_it)), n_comb_dta.index.values[j][0], format3)
     #   worksheet.write_formula(str('A'+str(j+st_it)), str('='+str('A'+str(j+st_it))))  
   #     else:
   #         worksheet.write(j+st_it, 18, 0, cell_format4)
            #worksheet.write(str('R'+str(j+7)), '',  cell_format3)
  #      worksheet.write(str('R'+str(j)), '',  cell_format4)    

       # print(str('R'+str(j+6)))
    
 #   cell.protection = Protection(locked=False)
    writer.save()
    workbook.close()
    
       
    mywb = px.load_workbook(output_file)

    mysheet = mywb.active
    border1=Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='hair'), 
                     bottom=Side(style='hair'))
   # top = Border(top=border.top)
    
    
    for j in range(1, len(n_comb_eff_1),2):
     #  print(str('A'+str(j+st_it))+":"+str('A'+str(j+st_it+1)))
       mysheet.unmerge_cells(str('A'+str(j+st_it))+":"+str('A'+str(j+st_it+1)))
       mysheet.cell(row = j+st_it+1, column = 1).value = str("="+str('A'+str(j+st_it)))
       mysheet.cell(row = j+st_it+1, column = 1).font=Font(color=colors.WHITE)
       mysheet.cell(row = j+st_it+1, column = 1).border = border1
       
 #   mysheet.protection.sheet = True   
 #   area =mysheet['A7':'Q50']
 #   area.protection = Protection(locked=False)
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"Trade at spot,Trade at close, Do not trade"', allow_blank=True)
    dv.error ='Your entry is not in the list'
    dv.errorTitle = 'Invalid Entry'
    dv.prompt = 'Please select from the list'
    dv.promptTitle = 'List Selection'
    mysheet.add_data_validation(dv)
    dv.add(str('V'+str(st_it+1))+":"+str('V'+str(st_it+100)))
    
    from openpyxl.worksheet.datavalidation import DataValidation
    dw = DataValidation(type="list", formula1='"0,1"', allow_blank=True)
    dw.error ='Your entry is not in the list'
    dw.errorTitle = 'Invalid Entry'
    dw.prompt = 'Please select from the list'
    dw.promptTitle = 'List Selection'
    mysheet.add_data_validation(dw)
    dw.add(str('U'+str(st_it+1))+":"+str('U'+str(st_it+100)))
   
    
    
    mysheet.protection.sheet = True
    mysheet.protection.password = 'Flower'
    mysheet.protection.autoFilter=False
 #   for h in range(1, len(n_comb_eff_1),1):
 #       cell = mysheet[str('A'+str(h))]
 #       cell.protection = Protection(autofilter=True, locked=False)
    from openpyxl.comments import Comment
    comment = Comment('Sign-off','PM')
    comment.height = 20
    mysheet["B5"].comment = comment
    
    mywb.save(output_file) 
    
    os.startfile(output_folder)
"""
'******************************************************************************************************************************************************************************    
'                                                     Define user input to check if flows are updated
'******************************************************************************************************************************************************************************    
"""
    
def input_fx(termi_nate_cnt=5):
    cnt=0
    loop=True
    while loop:
        d1a = input ("1. Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?: ")
        
        if d1a=="Y":
            print ("Futures report generation in progress",end='', flush=True)
            break
        elif d1a=="N":
            print ("Please update flows file",end='', flush=True)
            break
        else:
            cnt=cnt+1
            #print(cnt)
            print("Invalid input, please select the correct option")
            if cnt==termi_nate_cnt:
                print("You have run out of options, default option selected")
                d1a='N'
                #loop=False
                #print(loop)
                break
    
    if loop:
        x = [d1a]

    return x    
"""
'******************************************************************************************************************************************************************************    
'                                                     Define futures load function for futures tloader to Decalog
'******************************************************************************************************************************************************************************    
"""

# Tloader for futures

def tloader_fmt_futures(termi_nate_cnt=5):

    import sys
    import pandas as pd
    import numpy as np
    import datetime as dt
    from datetime import datetime, timedelta
    import glob
    import os
    #from pydatastream import Datastream
    #from business_calendar import Calendar, MO, TU, WE, TH, FR
    import pyodbc
    import xlrd

        
    startDate = datetime.today().date()
    
    folder_yr = datetime.strftime(startDate, "%Y")
    folder_mth = datetime.strftime(startDate, "%m")
    folder_day = datetime.strftime(startDate, "%d")
    
    
    dirtoimport_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES'
    
    #output_folder='\\'.join([dirtooutput_file ,folder_yr, folder_mth])
    input_folder=str('\\'.join([dirtoimport_file ,folder_yr, folder_mth,folder_day])+'\\Futures Trades\\')
    
    dic_users={'blala':['BLL','blala'], 'test':['TST','test'], 'sbisho':['SB','sbisho'], 'tmfelang2':['TM','tmfelang'], 'abalfour':['AB','abalfour'], 'sparker2':['SP','sparker'], 'fsibiya':['FS','fsibiya']}
    dirtooutput_file= 'U:\\Production\\In\\'
    
    newest = max(glob.iglob(input_folder+'IndexFutRep_*.xlsx'), key=os.path.getmtime)
    
    check = xlrd.open_workbook(newest)
    sh='Sheet1'
        
    ls = []
    cols = range(0,23)
    for i in cols:
        #    print(i)
        ls.append(i)
    
    fund_xls = pd.read_excel(newest, sheet_name = sh, skiprows =7, usecols = ls)

    checksheet = check.sheet_by_name('Sign-Off')
    value_app = checksheet.cell(5, 0)
    
    checksheet2 = check.sheet_by_name('Sheet1')
    
    run_load=1
    for j in range(9,(len(fund_xls)+9),2):
        value = checksheet2.cell(j, 20)
        if str(value) == 'number:1.0':
            value_cell = checksheet2.cell(j, 21)
            if str(value_cell)[0:11]=="text:'Trade":
              run_load_x=1    
            else:
              run_load=0
            #  print("Please enter a trade comment before loading")
         #        break
  
    if run_load==1:
        
        if (str(value_app)== "text:'Approved'"):
        
            fund_xls_ex= fund_xls[fund_xls.TradeComment.isin(['Trade at spot','Trade at close'])]
            fund_xls_ex=fund_xls_ex.copy()
            fund_xls_ex['TradeShort']= np.where(fund_xls_ex['Trade'].values=='Sell', 'SOC', 
                                                         np.where(fund_xls_ex['Trade'].values=='Buy', 'BOC', ''))
            fund_xls_ex['Nom']=abs(fund_xls_ex['No. Futures / Price'].values)
            fund_xls_x=fund_xls_ex[['FutureCode', 'Portfolio Code', 'TradeShort', 'Nom']]
            fund_xls_x=fund_xls_x.copy()
            fund_xls_x['Instruction']='Rebalance Portfolio'
            fund_xls_x['MP']='MP'
            fund_xls_x['Blank']=''
            fund_xls_x['TradeIns']=  fund_xls_ex['TradeComment']
            
            
            with open(str(dirtooutput_file+"FuturesTrade"+folder_yr+folder_mth+folder_day+'.txt'), "w") as fin:
          #  with open(str('c:\\data\\'+"FuturesTrade"+folder_yr+folder_mth+folder_day+'.txt'), "w") as fin:
                #fin.write('\n'.join((fund_xls_ex.values.tolist())[0]))
                for i in range(0,len(fund_xls_ex)):
                    print(i)
                    st=fund_xls_x.values.tolist()[i]
                    sf=st[0]+','+st[1]+',',st[2]+','+str(int(st[3]))+','+st[4]+','+st[5]+','+st[6]+','+st[7]+'\n'
                    sh=''.join(sf)
                    #sf=st[0]+','+st[1]+',',st[2]+','+st[4]+','+st[5]+','+st[6]+','+st[7]
                    fin.write(sh)
                # Get sheet names
            print("Trades loaded into Decalog!!")
                
        else:
            print("Trade has not been approved, please sign-off before loading, Trades are not loaded!!")
            
    else:
        print("Please enter a trade comment before loading, Trades are not loaded!!")
        

"""
        
'******************************************************************************************************************************************************************************    
'                                                     Define input function for Equity tloader to Decalog
'******************************************************************************************************************************************************************************    
"""
        
def tloader_fmt_equity(termi_nate_cnt=5):

    import sys
    import pandas as pd
    import numpy as np
    import datetime as dt
    from datetime import datetime, timedelta
    import glob
    import os
    #from pydatastream import Datastream
    #from business_calendar import Calendar, MO, TU, WE, TH, FR
    import pyodbc
    
    
    startDate = datetime.today().date()
    
    folder_yr = datetime.strftime(startDate, "%Y")
    folder_mth = datetime.strftime(startDate, "%m")
    folder_day = '08'# datetime.strftime(startDate, "%d")
    
    
    dirtoimport_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES'
    
    #output_folder='\\'.join([dirtooutput_file ,folder_yr, folder_mth])
    input_folder=str('\\'.join([dirtoimport_file ,folder_yr, folder_mth,folder_day])+'\\Equity Trades\\')
    
    dic_users={'blala':['BLL','blala'], 'test':['TST','test'], 'sbisho':['SB','sbisho'], 'tmfelang2':['TM','tmfelang'], 'abalfour':['AB','abalfour'], 'sparker2':['SP','sparker'], 'fsibiya':['FS','fsibiya']}
    dirtooutput_file= 'U:\\Production\\In\\'
    
    newest = max(glob.iglob(input_folder+'Trade*.xlsx'), key=os.path.getmtime)
    
    sh='TradeList'
    
    ls = []
    cols = range(0,11)
    for i in cols:
        print(i)
        ls.append(i)
        
    fund_xls = pd.read_excel(newest, sheet_name = sh, skiprows =1, usecols = ls)
    #fund_xls['AlpCode']= (fund_xls['Asset ID'])[1:]
    fund_xls = fund_xls[fund_xls['Asset ID'] != 'ZAR']
    fund_xls.loc[:,'AlpCode'] = fund_xls['Asset ID'].apply(lambda x : x[2:] if x.startswith("ZA") else x)   
    fund_xls.loc[:,'TradeShort']= np.where(fund_xls['Trade Type'].values=='SELL', 'S', 
                                                 np.where(fund_xls['Trade Type'].values=='BUY', 'B', ''))
    
#     fund_xls_ex= fund_xls.loc[fund_xls['Trade Comment'] == 1]
#    fund_xls_ex['Nom']=abs(fund_xls_ex['No. Futures'].values)
#    fund_xls_ex=fund_xls_ex[['FutureCode', 'Portfolio Code', 'TradeShort', 'Nom']]
#    fund_xls_ex['Instruction']='Rebalance Portfolio'
#    fund_xls_ex['MP']='MP'
#    fund_xls_ex['Blank']=''
#    fund_xls_ex['TradeIns']='Trade at spot'
#    
    
    with open(str(dirtooutput_file+"EquityTrade"+folder_yr+folder_mth+folder_day+'.txt'), "w") as fin:
  #  with open(str('c:\\data\\'+"EquityTrade"+folder_yr+folder_mth+folder_day+'.txt'), "w") as fin:
        #fin.write('\n'.join((fund_xls_ex.values.tolist())[0]))
        for i in range(0,len(fund_xls)):
    #    for i in range(0,10):
            print(i)
            st=fund_xls.values.tolist()[i]
            sf=st[11]+','+st[1]+','+st[12]+','+str(int(abs(st[5])))+',''Rebalance Portfolio'+','+'MP'+',,'+'BLALA:Trade at Close'+'\n'
            sh=''.join(sf)
            #sf=st[0]+','+st[1]+',',st[2]+','+st[4]+','+st[5]+','+st[6]+','+st[7]
            fin.write(sh)
    # Get sheet names

"""    
'******************************************************************************************************************************************************************************    
'                                                                   Select funds to trade
'******************************************************************************************************************************************************************************    
"""

def select_fund():
    import pandas as pd
    get_fund_list = pd.read_csv('C:\\IndexTrader\\required_inputs\\flows.csv')
    get_funds = (get_fund_list[(get_fund_list.Trade==1)])['Port_code'].tolist()

    return get_funds