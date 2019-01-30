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


def excel_fx(output_folder,dic_users,n_comb_eff_1,startDate,newest):
    
    import pandas as pd
    import numpy as np
    import datetime as dt
    import os
    from datetime import datetime, timedelta
    import openpyxl as px
    from openpyxl.styles import colors, Font, Border, Side ,Protection
    import time
    
    
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
    worksheet.write_string('D3', 'File used:',cell_format2)
    worksheet.write_string('E3', newest,cell_format2_1)
    worksheet.write_string('D4', 'Timestamp:',cell_format2)
    worksheet.write_string('E4', time.ctime(os.path.getmtime(newest)) ,cell_format2_1)
   
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

def select_fund(struc=True):
    import pandas as pd
    get_fund_list = pd.read_csv('C:\\IndexTrader\\required_inputs\\flows.csv')
    get_funds = (get_fund_list[(get_fund_list.Trade==1)])['Port_code'].tolist()
    if struc:
        get_funds = list(set(get_funds + ['CORPEQ'])) # Add corpeq to get the underlyng aset classifications    
    return get_funds


"""    
'******************************************************************************************************************************************************************************    
'                                                                   Systematic rules to get trading action
'******************************************************************************************************************************************************************************    
"""


def CashFlowFlag(Eff_cash,Total_cash, mx_totcash, mn_totcash, mx_effcash, mn_effcash, md_totcash, flw, fut_exp):
    
    if Eff_cash > mx_effcash:
        if Total_cash >= mx_totcash:
            if Total_cash >= (mx_totcash+md_totcash):
                if flw==0:
                     return 'Trade Equity + Futures (midpoint)'
                else:
                     return 'Trade Equity + Futures (only flow)'
            elif Total_cash < (mx_totcash+md_totcash):
                return 'Trade Equity + Futures (midpoint)'
        elif Total_cash <= mn_totcash: # unlikely to occur
            if Total_cash <= (mn_totcash-md_totcash):
                if flw==0:
                     return 'Trade Equity + Futures (midpoint)'
                else:
                     return 'Trade Equity + Futures (only flow)'
            elif Total_cash > (mn_totcash-md_totcash):
                return 'Trade Equity + Futures (midpoint)'
            #return 'Trade Equity + Futures' # unlikely to occur
        else: 
            return 'Trade Futures only'
    elif Eff_cash < mn_effcash:
        if Total_cash <= mn_totcash:
             if Total_cash <= (mn_totcash-md_totcash):
                 if flw==0:
                     return 'Trade Equity + Futures (midpoint)'
                 else:
                     return 'Trade Equity + Futures (only flow)'
             elif Total_cash > (mn_totcash-md_totcash):
                return 'Trade Equity + Futures (midpoint)'
            #return 'Trade Equity + Futures'
        elif Total_cash >= mx_totcash:
            if Total_cash >= (mx_totcash+md_totcash):
                if flw==0:
                     return 'Trade Equity + Futures (midpoint)'
                else:
                     return 'Trade Equity + Futures (only flow)'
            elif Total_cash < (mx_totcash+md_totcash):
                return 'Trade Equity + Futures (midpoint)'
         #   return 'Trade Equity + Futures' # unlikely to occur
        else:
            return 'Trade Futures only'
    else:
        if Total_cash < mn_totcash:
           return 'Trade Equity'
        elif Total_cash > mx_totcash:
            if fut_exp > mx_totcash: # Need to think around how to deal with this event
                return 'No Action'
            else:
                return 'Trade Equity'
        else:
           return 'No Action'
    
"""    
'******************************************************************************************************************************************************************************    
'                                                                   Systematic rules to get trading targets
'******************************************************************************************************************************************************************************    
"""
    
      
def trade_calc(Flag, tgt_effcash, tgt_totcash, fut_code,  mx_effcash, mn_effcash, ovrd_effcash, aeff_cash, atot_cash, fnd_val, fut_price, fut_exp, flw):
    import numpy as np

#    if Flag == 'Trade Equity + Futures':
#        tgt_totcash = tgt_totcash
#        tgt_effcash = tgt_effcash
    if Flag == 'Trade Equity + Futures (only flow)':
        t_effcash=aeff_cash-flw
        t_totcash=atot_cash-flw
        if ((t_effcash>mx_effcash)or(t_effcash<mn_effcash)):
            t_effcash=tgt_effcash
    elif Flag == 'Trade Equity + Futures (midpoint)': 
        t_effcash=aeff_cash-flw
        t_totcash=tgt_totcash
        if ((t_effcash>mx_effcash)or(t_effcash<mn_effcash)):
            t_effcash=tgt_effcash
    elif Flag == 'Trade Futures only':
         x_trd = np.where(fut_code=="NoFuture", "No Trade",
                                    np.where((aeff_cash>mx_effcash), 'Buy', 
                                             np.where((aeff_cash<mn_effcash), 'Sell', 'No Trade')))
         
         x_trd2 = np.where((x_trd=="No Trade")&(~np.isnan(ovrd_effcash))&(not(fut_code=="NoFuture")),
                                    np.where((aeff_cash>tgt_effcash), 'Buy', 
                                             np.where((aeff_cash<tgt_effcash), 'Sell', 'No Trade')),x_trd)
         no_fut = np.where(np.isin(x_trd2, ['Buy','Sell']), np.rint(((aeff_cash-tgt_effcash)*fnd_val)/(fut_price*10)), 0)
         fut_exp1=((no_fut*10*fut_price)/fnd_val)+fut_exp
         if (atot_cash - fut_exp1) < 0:
              t_totcash = np.where(np.isnan(fut_exp), 0, fut_exp) + tgt_effcash
              t_effcash = tgt_effcash
         else:
             t_effcash = np.where((atot_cash - fut_exp1) < 0, (atot_cash - fut_exp), (atot_cash - fut_exp1))
             t_totcash = atot_cash
    
    elif Flag == 'Trade Equity':
        
         t_totcash = np.where(np.isnan(fut_exp), 0, fut_exp) + tgt_effcash
         t_effcash = tgt_effcash
     
    else:
        t_totcash = atot_cash
        t_effcash = aeff_cash
         
    return [np.round(t_effcash,20),np.round(t_totcash,20)]    

"""    
'******************************************************************************************************************************************************************************    
'                                                                   Bulk cash calc excel report
'******************************************************************************************************************************************************************************    
"""

def bulk_cash_excel_report(startDate,new_dat_pf,new_dat, n_comb,dic_users,dic_om_index,newest,output_folder,fnd_excp):
    
    import pandas as pd
    import numpy as np
    import datetime as dt
    import os
    from datetime import datetime, timedelta
    import openpyxl as px
    from openpyxl.styles import colors, Font, Border, Side ,Protection
    #import openpyxl.cell
    #from openpyxl import load_workbook
    from write_excel import select_fund as sf
    import xlsxwriter
    import time
    
    start_time = datetime.now() 
    
    auto_trade=True
 #   output_folder= 'c:/data/'
    
    
    output_file = output_folder+'\\BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsx'
    st_row = 19
#    st_it = st_row+1
    
    st_col= 'C'
    #writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    #hdr = new_dat_pf.columns.tolist()
    
    
    new = new_dat_pf.reset_index()
    new= new[~new.AssetType5.isin(['SSF DIV'])]
    
    new_flow = n_comb
    inv=(new_flow[['Port_code','fin_teff_cash','fin_tot_cash', 'InvType','CashFlowFlag','Min_EffCash','Max_EffCash','Min_TotalCash','Max_TotalCash']]).set_index('Port_code').T.to_dict('list')
    
    lst_fund = sf(False)    
    #new.pivot(index=['Port_code','AssetType1'], columns= 'AssetType3', values='EffExposure')
    
    #h=pd.pivot_table(new,  index=['Port_code'],columns=['AssetType1','AssetType3'], values=['EffExposure','EffWgt'],aggfunc=np.sum)
    
    
    qf=pd.DataFrame([], dtype='object')
    futures_dict = dict()
    cashflow_dict = dict()
    
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
        
    #wf.to_excel(writer, sheet_name='Sheet1', startrow=st_row, startcol =5)
    
    qf.to_excel(writer, sheet_name='Summary', startrow=st_row, startcol =0, header = False)
    
    workbook  = writer.book
    
    workbook.filename= output_folder+'\\BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsm'
    workbook.add_vba_project('C:/IndexTrader/code/vbaProject.bin')
    
    
    #writer.save()
    
    
    worksheet = writer.sheets['Summary']
    
    
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
    #cell_format4 = workbook.add_format({'bold': True, 'bg_color':'#339966', 'font':11, 'locked':False })
    cell_format5 = workbook.add_format({'bold': True, 'bg_color':'#C0C0C0', 'font':11, 'align': 'center','border': 1})
    cell_format6 = workbook.add_format({'bold': True, 'bg_color':'#CCFFFF', 'font':11,  'align': 'center','border':1})
    
    cell_format5_1 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': 'R#,##0','font_color': 'red'  })
    cell_format5_2 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': '0.00%','font_color': 'red'  })
    
    cell_format6_1 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': 'R#,##0', 'font_color': 'red' })
    cell_format6_2 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': '0.00%', 'font_color': 'red' })
    
    cell_format5_3 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': 'R#,##0','font_color': 'black'  })
    cell_format5_4 = workbook.add_format({'bold': False, 'bg_color':'#C0C0C0', 'font':10, 'locked':False,'num_format': '0.00%','font_color': 'black'  })
    
    cell_format6_3 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': 'R#,##0', 'font_color': 'black' })
    cell_format6_4 = workbook.add_format({'bold': False, 'bg_color':'#CCFFFF', 'font':10, 'locked':False,'num_format': '0.00%', 'font_color': 'black' })
    
    
    cell_format7 = workbook.add_format({'bold': True, 'bg_color':'#CCFFCC', 'font':11, 'align': 'center','border': 1})
    cell_format8 = workbook.add_format({'bold': True, 'font':11, 'font_color': '#339966','align': 'center','border': 1})
    
    td_cell_format_1 = workbook.add_format({'bold': True, 'bg_color':'#FFC7CE', 'font':10,'font_color': '#9C0006'  })
    td_cell_format_2 = workbook.add_format({'bold': True, 'bg_color':'#C6EFCE', 'font':10,'font_color': '#006100'  })
    
    #td_cell_format_3 = workbook.add_format({'bold': True, 'bg_color':'#FF0000', 'font':10, 'locked':False,'font_color': '#9C0006','align': 'center'  })
    td_cell_format_4 = workbook.add_format({'bold': True, 'bg_color':'#CCCCFF', 'font':11, 'locked':False,'font_color': '#0000FF','align': 'center'  })
    td_cell_format_4.set_font_size(11)
                                          
    format1 = workbook.add_format({'num_format': 'R#,##0'})
    format1_1 = workbook.add_format({'num_format': 'R#,##0', 'locked': False})
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
    
    worksheet.write_string('D3', 'File used:',cell_format2)
    worksheet.write_string('E3', newest,cell_format2_1)
    worksheet.write_string('D4', 'Timestamp:',cell_format2)
    worksheet.write_string('E4', time.ctime(os.path.getmtime(newest)) ,cell_format2_1)
   
    
    
    
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
    cell_format1_1 = workbook.add_format({'bold': True, 'font_color': 'black', 'font':12, 'locked': False})
    
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
        worksheet.write_number(str(alp[st_col]+str(st_row-8)), cashflow_dict[j][0] ,format1_1)
        
        worksheet.write_string(get_pl4, inv[j][2] ,cell_format1_1)
        worksheet.data_validation(get_pl4, {'validate': 'list',
                                            'source': ['Investment', 
                                                       'Hedged Withdrawal', 
                                                       'Hedged With Pay(t)',
                                                       'Withdrawal Pay(t)',
                                                       'No cash flow'],
                                            'input_title': 'Enter a cash flow',
                                            'input_message': 'Type & Value'} )
        
        if auto_trade:
            worksheet.write_number(str(alp[st_col]+str(st_row-5)), inv[j][0] ,fmts3_2[jet])
            worksheet.write_number(str(alp[st_col]+str(st_row-4)), inv[j][1] ,fmts3_2[jet])
            worksheet.conditional_format(get_pl5,{'type': 'cell','criteria': '!=','value': inv[j][0] ,'format': td_cell_format_4})
            worksheet.conditional_format(get_pl6,{'type': 'cell','criteria': '!=','value': inv[j][1] ,'format': td_cell_format_4})
        
        else: 
            worksheet.write_formula(get_pl5, str('='+str(alp[st_col+1]+str(st_row+26))),fmts3_2[jet])
            worksheet.write_formula(get_pl6, str('='+str(alp[st_col+1]+str(st_row+17))),fmts3_2[jet])
            worksheet.conditional_format(get_pl5,{'type': 'cell','criteria': '!=','value': str('='+str(alp[st_col+1]+str(st_row+26))) ,'format': td_cell_format_4})
            worksheet.conditional_format(get_pl6,{'type': 'cell','criteria': '!=','value': str('='+str(alp[st_col+1]+str(st_row+17))) ,'format': td_cell_format_4})
        
        
        worksheet.write_formula(get_pl7, str('='+str(alp[st_col+1]+str(st_row+56))),fmts3_2[jet])
        worksheet.write_formula(get_pl8, str('='+str(alp[st_col+1]+str(st_row+47))),fmts3_2[jet])
        
        worksheet.conditional_format(get_pl2, {'type': 'cell','criteria': '<','value': 0,'format': fmts2_1[jet]})                              
        worksheet.conditional_format(get_pl2, {'type': 'cell','criteria': '>=','value': 0,'format': fmts2_2[jet]})   
        worksheet.conditional_format(get_pl3, {'type': 'cell','criteria': '<','value': 0,'format': fmts3_1[jet]})                              
        worksheet.conditional_format(get_pl3, {'type': 'cell','criteria': '>=','value': 0,'format': fmts3_2[jet]})      
        
        worksheet.conditional_format(get_pl4,{'type': 'cell','criteria': 'equal to','value': '"No cash flow"','format': cell_format1_1})
        worksheet.conditional_format(get_pl4,{'type': 'cell','criteria': '!=','value': '"No cash flow"','format': td_cell_format_4})
        
        worksheet.merge_range(str(alp[st_col]+str(st_row-3)+":"+alp[st_col+1]+str(st_row-3)), futures_dict[j][0] ,cell_format7)
        worksheet.merge_range(str(alp[st_col]+str(st_row-2)+":"+alp[st_col+1]+str(st_row-2)), dic_om_index[j][1] ,cell_format8)
       
        
        for g in range(st_row, st_row+ 25):
            if g in [st_row+1, st_row+2,st_row+7, st_row+11,st_row+13]:
                
                worksheet.conditional_format(str(alp[st_col]+str(g)),{'type': 'cell','criteria': '!=','value': inv[j][0] ,'format': format1_b})
                worksheet.conditional_format(str(alp[st_col+1]+str(g)),{'type': 'cell','criteria': '!=','value': inv[j][0] ,'format': format2_b})
        
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
    worksheet.merge_range(str('A'+str(st_row)+':B'+str(st_row)),'PRE-TRADE', cell_format2_4)
    
    # end grouping
    def form_at(st_no=(st_row+len(qf)),_label_='INCLUDING FLOWS'):
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
        
        if _label_=='POST-TRADE':
             worksheet.write_string(str('B'+str(in_st+17)), 'BARRA CASH',bn_cell_format1)
            
    
    form_at()
    form_at(st_no=(st_row+2*len(qf)+1),_label_='TRADE')
    form_at(st_no=(st_row+3*len(qf)+2),_label_='POST-TRADE')
    
    
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
                  #              str('=IF('+str(alp[td_col]+str(td_st-33))+'="No Future",0,INT((ROUND('+str(alp[td_col]+str(td_st-34))+'-'+str(alp[td_col]+str(td_st-35))+\
                                str('=IF('+str(alp[td_col]+str(td_st-33))+'="No Future",0,IF(('+str(alp[td_col]+str(td_st-34))+'-'+str(alp[td_col]+str(td_st-35))+\
                                    '-'+str(alp[td_col+1]+str(td_st-8))+')>0,ROUNDDOWN((('+str(alp[td_col]+str(td_st-34))+'-'+str(alp[td_col]+str(td_st-35))+\
                                    '-'+str(alp[td_col+1]+str(td_st-8))+')*'+str(alp[td_col]+str(td_st-14))+')/'+ \
                                        str(alp[td_col+1]+str(td_st-21))+'/10,0),ROUNDUP(ROUND((('+str(alp[td_col]+str(td_st-34))+'-'+str(alp[td_col]+str(td_st-35))+\
                                    '-'+str(alp[td_col+1]+str(td_st-8))+')*'+str(alp[td_col]+str(td_st-14))+')/'+ \
                                        str(alp[td_col+1]+str(td_st-21))+'/10,9),0)'+'))'),cell_format2_1)
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
            elif g in [pt_st+3,pt_st+4,pt_st+5,pt_st+6,pt_st+8,pt_st+10,pt_st+12,pt_st+14]:
                worksheet.write_formula(str(alp[pt_col+1]+str(g)),str('='+str(alp[pt_col]+str(g))+'/'+str(alp[pt_col]+'$'+str(pt_st+1))),informat2_n)
         #       print(str('N='+str(alp[td_col]+str(g))+'/'+str(alp[td_col]+'$'+str(td_st+1))))
            else:
                gg=1
       
        # worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+11)), {'type': 'cell','criteria': '<','value': 0,'format': pt_cell_format_1})    
        # worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+11)), {'type': 'cell','criteria': '>','value': 0,'format': pt_cell_format_2})
         
         worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+11)), {'type':'cell','criteria': 'between','minimum': inv[f][4],
                                      'maximum': inv[f][5],'format': pt_cell_format_2})
        
         worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+11)), {'type':'cell','criteria': 'not between','minimum':  inv[f][4], 
                                           'maximum':  inv[f][5],
                                           'format':   pt_cell_format_1})
           
         worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+2)), {'type':'cell','criteria': 'between','minimum':  inv[f][6], 
                                           'maximum':  inv[f][7],
                                           'format':   pt_cell_format_2})
        
         worksheet.conditional_format(str(alp[pt_col+1]+str(pt_st+2)), {'type':'cell','criteria': 'not between','minimum':  inv[f][6], 
                                           'maximum':  inv[f][7],
                                           'format':   pt_cell_format_1})
         
         pt_col=pt_col+2
                           
    del f 
    del g
    del gg
    
    worksheet.freeze_panes(19,2)
    #worksheet.write_comment('C10', 'Enter cash flow info below', {'start_col': 5,'start_row': 7, 'x_scale': 1.2, 'y_scale': 0.25, 'visible': True ,'font_size': 11, 'bold':True ,'color': '#FFCC99'})
    
    #pandaswb = writer.book
    #pandaswb.filename =  output_folder+'\\BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsm'
    #pandaswb.add_vba_project('C:/Program Files (x86)/WinPython/python-3.6.5.amd64/Scripts/vbaProject.bin')
    
    #pandaswb.save()        
    #pandaswb.close()                                                           
    writer.save()
    workbook.close()
    
    
    mywb = px.load_workbook(output_folder+'\\BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsm', keep_vba=True)
    mysheet = mywb.active
    
    border1=Border(right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
      
    
    for j in range(4, 4+len(lst_fund)*2-2,2):
     #   print(j)
        mysheet.cell(row = 7, column = j).border = border1
        mysheet.cell(row = 16, column = j).border = border1
        mysheet.cell(row = 17, column = j).border = border1
           
    
    mysheet.protection.sheet = False
    mysheet.protection.password = 'Flower'
    
    mywb.save(output_folder+'\\BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsm') 
        
    os.startfile(output_folder)  
        
    
    #import win32com.client #import Dispatch
    
    #xl = win32com.client.Dispatch("Excel.Application")  # Set up excel
    #wb=xl.Workbooks.Open(Filename = output_folder+'BatchCashCalc_'+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.xlsm')         # Open .xlsm file from step 2A
    #xl.Application.Run("Module1")                  # Run VBA_macro.bin
    #xl.Application.Run("auto_open")
    
    time_elapsed = datetime.now() - start_time 
    print('Time elapsed (hh:mm:ss.ms) {}'.format(time_elapsed))
    
    
    #workbook.add_vba_project('C:/IndexTrader/code/vbaProject.bin')
"""    
'******************************************************************************************************************************************************************************    
'                                                                   Create BPM Cash File
'******************************************************************************************************************************************************************************    
"""



def create_BPMcashfile(fnd_excp= ['DSALPC','OMCC01','OMCD01','OMCD02','OMCM01','OMCM02','PPSBTA','PPSBTB']):
    
    import win32com.client #import Dispatch
    from write_excel import select_fund as sf
    import os
    import numpy as np
    import pandas as pd
    from datetime import datetime, timedelta
    from tkinter import filedialog
    from tkinter import Tk

    lst_fund = sf(False)   
    startDate = datetime.today()
   # startDate=startDate.replace(day=28)

    #fnd_excp= ['DSALPC','OMCC01','OMCD01','OMCD02','OMCM01','OMCM02','PPSBTA','PPSBTB']
    dirtoimport_file='\\\\za.investment.int\\dfs\\dbshared\\DFM\\TRADES'
    folder_yr = datetime.strftime(startDate, "%Y")
    folder_mth = datetime.strftime(startDate, "%m")
    folder_day = datetime.strftime(startDate, "%d")

    output_folder=str('\\'.join([dirtoimport_file ,folder_yr, folder_mth,folder_day])+'\\BatchTrades\\')
    root = Tk()
    root.filename =  filedialog.askopenfilename(initialdir = output_folder,title = "choose your file",filetypes = (("jpeg files","*.xlsm"),("all files","*.*")))
    print (root.filename)
    root.withdraw()
 
    dirtooutput_file=str(output_folder+'\\CashFile\\')
    
    if not os.path.exists(dirtooutput_file):
        os.makedirs(dirtooutput_file)
    
    user_dict=pd.read_csv('C:\\IndexTrader\\required_inputs\\user_dictionary.csv')
    dic_users=user_dict.set_index(['username']).T.to_dict('list')
 
        
    xl = win32com.client.Dispatch("Excel.Application")  # Set up excel
    wb=xl.Workbooks.Open(Filename = root.filename)         # Open .xlsm file from step 2A
    ws=wb.Sheets[0]
    
    with open(str(dirtooutput_file+"BPM_Cash"+startDate.strftime('%Y%m%d %H-%M-%S')+'_'+dic_users[os.environ.get("USERNAME").lower()][1]+'.csv'), "w",newline='\n') as fin:
    
        for j in range(3, len(lst_fund)*2+2, 2):
        
            fund = ws.Cells(7, j).value
            bpm_cash=ws.Cells(80, j).value
            bpm_futures=ws.Cells(73, j).value
            fut_code=ws.Cells(16, j).value
            
            if fund in fnd_excp:
                sf=fund+',ZAR,'+str(np.round(bpm_cash,15))+'\n'
                sh=''.join(sf)
            else:
                sf=fund+',ZAR,'+str(np.round(bpm_cash,15))+'\n'+fund+','+fut_code+','+str(bpm_futures)+'\n'
                sh=''.join(sf)
                #sf=st[0]+','+st[1]+',',st[2]+','+st[4]+','+st[5]+','+st[6]+','+st[7]
            fin.write(sh)
            
    print("BPM Cash File created")
    os.startfile(dirtooutput_file)  
   
    
    
    wb.Close(False)
    del xl
    

"""    
'******************************************************************************************************************************************************************************    
'******************************************************************************************************************************************************************************    
"""


"""    
'******************************************************************************************************************************************************************************    
'                                                                   Create Cash Flow Validity Fx 
'******************************************************************************************************************************************************************************    
"""



def cash_flow_validity_fx(cash_flows_eff,newest,startDate,lst_fund, bf=0.005):

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
    
"""    
'******************************************************************************************************************************************************************************    
'******************************************************************************************************************************************************************************    
"""
