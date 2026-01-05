#!/usr/bin/env python
# coding: utf-8
import time
cstart = time.time()

import pandas as pd
import warnings
warnings.filterwarnings('ignore')
from pandas.tseries.offsets import MonthEnd
from sqlalchemy import create_engine
import pandas as pd
import numpy as np
pd.set_option.max_columns= None
import datetime as dt
from datetime import datetime
from dateutil import rrule, relativedelta
from statistics import mean,median
import math
import time
import pickle
import os
print('Start time: ',dt.datetime.now())

# In[3]:


##update last_allowed_trading_date change it as per excel last working day every month, 2 working days. 1 apr(run day), 28 mar, 27 mar.==>27mar
##update yyyy,mm,dd
last_allowed_trading_date = dt.datetime(2025,11,28)
def_date = dt.datetime(2000,1,1)

##update 6 months only - remove oldest month and add latest excel name - file names must be in ascending order of month year .
adto_months_req = ['July 2025','Aug 2025','Sep 2025','Oct 2025', 'Nov 2025','Dec 2025']
#update to current month
month = 'Jan-2026'
#update to previous month run
prev_month = 'Dec-2025'

#keep updated path of manual reviewed stocks excel.
# manual_reviewed_stocks_path = r'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\Manual Reviewed Stocks.xlsx'
manual_reviewed_stocks_path = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\raw_data\Manual Reviewed Stocks.xlsx'

raw_data_path = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\raw_data'

# update file name angel scrip category , sometimes csv,xlsx,xls, uncomment angel_scrip_category_df read line as required

angel_scrip_category_fname = 'Angel Scrip Category 01012026.xlsx'
#angel_scrip_category_df = pd.read_csv(os.path.join(raw_data_path,angel_scrip_category_fname))
#angel_scrip_category_df = angel_scrip_category_df.rename(columns= {"ISINNo" : "ISIN No"})
angel_scrip_category_df = pd.read_excel(os.path.join(raw_data_path,angel_scrip_category_fname),engine="openpyxl")
angel_scrip_category_df = angel_scrip_category_df.rename(columns= {"ISINNo" : "ISIN No"})

# angel_scrip_category_fname = 'Angel Scrip Category 01092025.xlsx'
# angel_scrip_category_df = pd.read_excel(os.path.join(raw_data_path,angel_scrip_category_fname),header =0 )
# angel_scrip_category_df = angel_scrip_category_df.rename(columns= {"ISINNo" : "ISIN No"})

# angel_scrip_category_fname = 'Angel Scrip Category 01092025.xls'
# angel_scrip_category_df = pd.read_excel(os.path.join(raw_data_path,angel_scrip_category_fname),engine="xlrd")
# angel_scrip_category_df = angel_scrip_category_df.rename(columns= {"ISINNo" : "ISIN No"})

#new added 20-09-2024
funding_fname = 'FUNDING 01012026.xlsx'
funding_df = pd.read_excel(os.path.join(raw_data_path,funding_fname) , header = 0)
print(funding_df.columns)
funding_df = funding_df.rename(columns= {"%" : "% Funding" , "%Funding" : "% Funding", "Funding %" : "% Funding","%Funding" : "% Funding", "%Funding " : "% Funding", "% Funding " : "% Funding"," %Funding " : "% Funding", " % Funding" : "% Funding"})
print(funding_df.columns)
# In[7]:
def clean_impact_cost_file(raw_data_path_):
    impact_cost_excel_path = os.path.join(raw_data_path_,'Impact Cost.xls')
    impact_cost_to_clean =  pd.read_excel(impact_cost_excel_path)
    impact_cost_cleanded = impact_cost_to_clean.dropna(how='any').reset_index(drop=True)
    # impact_cost_cleanded = impact_cost_to_clean.dropna(subset=['[Impact Cost (Latest)].1']).reset_index(drop=True)
    writer = pd.ExcelWriter( impact_cost_excel_path, engine='xlsxwriter', datetime_format='dd/mm/yyyy')
    impact_cost_cleanded.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()

clean_impact_cost_file(raw_data_path)
# angel_scrip_category_fpath = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Scrip category'


# In[8]:


# adto_monthly_path = r'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Volume Turnover data'
adto_monthly_path = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\raw_data\ADTO data'


# In[9]:




# In[10]:


prev_month_excel = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{prev_month}\Output\Risk_Rating_{prev_month}.xlsx'


# In[11]:





# In[12]:


output_dir = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\Output\Risk_Rating_{month}.xlsx'


# In[15]:


asm_f_path = os.path.join(raw_data_path,'ASM')


# In[16]:


files_asm = os.listdir(asm_f_path)
files_asm = [file for file in files_asm if '.txt' not in file]


# In[17]:


for file in files_asm:
    if 'BSE' in file:
        if 'long' in file.lower():
            asm_long_term_df_bse = pd.read_csv(os.path.join(asm_f_path,file))
        else:
            asm_short_term_df_bse = pd.read_csv(os.path.join(asm_f_path,file))
    if 'NSE' in file:
        asm_long_term_df_nse = pd.read_csv(os.path.join(asm_f_path,file))


# In[18]:


# all_sheets = pd.read_excel(r'W:\Knowledge repositiory\Temp\Risk Rating\Final Version\Risk Rating Presentation v3.xlsx',sheet_name=None)


# In[19]:


# all_sheets.keys()


# In[20]:


fno_list_sheet = pd.read_excel(os.path.join(raw_data_path,'FnO companies.xls'))
nan_count_per_row = fno_list_sheet.isnull().sum(axis=1)
fno_list_sheet = fno_list_sheet[nan_count_per_row <= 2]
fno_list_sheet = fno_list_sheet.reset_index(drop=True)


# In[21]:


financial_data = pd.read_excel(os.path.join(raw_data_path,'Financial Data.xls'))
nan_count_per_row = financial_data.isnull().sum(axis=1)
financial_data = financial_data[nan_count_per_row <= 70]
financial_data = financial_data.reset_index(drop=True)


# In[22]:


Networth_sheet = pd.DataFrame()
Networth_sheet[['CAPITALINE CODE', 'CO_NAME', '[Year End (Latest)]',
       '[Year (Latest)]', '[Networth (Latest)]', '[MODE (Latest)]']] = financial_data[['CAPITALINE CODE', 'CO_NAME', '[Year End (Latest)]',
               '[Year (Latest)]', '[Total Shareholders Funds (Latest)]', '[MODE (Latest)]']]


# In[23]:

bse_nse_group_sheet = pd.read_excel(os.path.join(raw_data_path,'BSE and NSE group n500.Xls'))
nan_count_per_row = bse_nse_group_sheet.isnull().sum(axis=1)
bse_nse_group_sheet = bse_nse_group_sheet[nan_count_per_row <= 7]
bse_nse_group_sheet = bse_nse_group_sheet.reset_index(drop=True)


# In[27]:


nifty_500_sheet = bse_nse_group_sheet[['CAPITALINE CODE','CO_NAME','[ISIN No','[Nifty 500']]
nifty_500_sheet = nifty_500_sheet[nifty_500_sheet['[Nifty 500']=='Y']
nifty_500_sheet = nifty_500_sheet.reset_index(drop=True)
nifty_500_sheet['Index'] = 'Nifty 500'

latest_mcap_sheet_bse = pd.read_excel(os.path.join(raw_data_path,'Market Cap BSE.xls'))
nan_count_per_row = latest_mcap_sheet_bse.isnull().sum(axis=1)
latest_mcap_sheet_bse = latest_mcap_sheet_bse[nan_count_per_row <= 4]
latest_mcap_sheet_bse = latest_mcap_sheet_bse.reset_index(drop=True)

latest_mcap_sheet_nse = pd.read_excel(os.path.join(raw_data_path,'Market Cap NSE.xls'))
nan_count_per_row = latest_mcap_sheet_nse.isnull().sum(axis=1)
latest_mcap_sheet_nse = latest_mcap_sheet_nse[nan_count_per_row <= 4]
latest_mcap_sheet_nse = latest_mcap_sheet_nse.reset_index(drop=True)

# In[24]:


# latest_mcap_sheet_bse = latest_mcap_sheet_bse.rename(columns= { '[Market Cap (Latest)]' : '[Market Cap (Latest)]_bse' ,
#                                                          '[Date (Latest)]' : '[Date (Latest)]_bse'})

# latest_mcap_sheet_nse = latest_mcap_sheet_nse.rename(columns= { '[Market Cap (Latest)]' : '[Market Cap (Latest)]_nse' ,
#                                                          '[Date (Latest)]' : '[Date (Latest)]_nse'})


# In[25]:
latest_mcap_sheet = bse_nse_group_sheet[['CAPITALINE CODE','CO_NAME','[ISIN No']]
latest_mcap_sheet['[Market Cap (Latest)]'] = np.nan
latest_mcap_sheet['Exchange'] = ''
for idx in range(0,len(latest_mcap_sheet['CAPITALINE CODE'])):
    co_code_f = latest_mcap_sheet['CAPITALINE CODE'][idx]
    temp_mcap_bse = latest_mcap_sheet_bse[latest_mcap_sheet_bse['CAPITALINE CODE']==co_code_f].reset_index(drop=True)
    temp_mcap_nse = latest_mcap_sheet_nse[latest_mcap_sheet_nse['CAPITALINE CODE']==co_code_f].reset_index(drop=True)

    if len(temp_mcap_bse) ==0 and len(temp_mcap_nse) == 0 :
        latest_mcap_sheet['[Market Cap (Latest)]'][idx] = np.nan
        latest_mcap_sheet['Exchange'][idx] = ''

    elif len(temp_mcap_bse) ==0 and len(temp_mcap_nse) > 0 :
        latest_mcap_sheet['[Market Cap (Latest)]'][idx] = temp_mcap_nse['[Market Cap (Latest)]'][0]
        latest_mcap_sheet['Exchange'][idx] = 'NSE'
    
    elif len(temp_mcap_bse) >0 and len(temp_mcap_nse) == 0 :
        latest_mcap_sheet['[Market Cap (Latest)]'][idx] = temp_mcap_bse['[Market Cap (Latest)]'][0]
        latest_mcap_sheet['Exchange'][idx] = 'BSE'
    
    else:
        temp_mcap_bse_date = temp_mcap_bse['[Date (Latest)]'][0]
        temp_mcap_nse_date = temp_mcap_nse['[Date (Latest)]'][0]

        if temp_mcap_bse_date >= temp_mcap_nse_date :
            latest_mcap_sheet['[Market Cap (Latest)]'][idx] = temp_mcap_bse['[Market Cap (Latest)]'][0]
            latest_mcap_sheet['Exchange'][idx] = 'BSE'
        elif temp_mcap_bse_date < temp_mcap_nse_date :
            latest_mcap_sheet['[Market Cap (Latest)]'][idx] = temp_mcap_nse['[Market Cap (Latest)]'][0]
            latest_mcap_sheet['Exchange'][idx] = 'NSE'
        
latest_mcap_sheet = latest_mcap_sheet[latest_mcap_sheet['Exchange']!=''].reset_index(drop=True)
# latest_mcap_sheet.to_excel('test_mcap.xlsx')
# import sys
# sys.exit(0)
# for idx in range(0,len(latest_mcap_sheet['[Market Cap (Latest)]'])):
#     if (not pd.isna(latest_mcap_sheet['[Market Cap (Latest)]_nse'][idx])):
#         latest_mcap_sheet['[Market Cap (Latest)]'][idx] = latest_mcap_sheet['[Market Cap (Latest)]_nse'][idx]
#         latest_mcap_sheet['Exchange'][idx] = 'NSE'
#     else:
#         latest_mcap_sheet['[Market Cap (Latest)]'][idx] = latest_mcap_sheet['[Market Cap (Latest)]_bse'][idx]
#         latest_mcap_sheet['Exchange'][idx] = 'BSE'


# In[26]:





# In[28]:


pet_check_code = pd.DataFrame()
pet_check_code[['CO_CODE', 'CO_NAME', 'YearEnd1', 'YearEnd2',
       'Sales Y1', 'Sales Y2', 'Power & Fuel Cost Y1', 'Power & Fuel Cost Y2',
       'Employee Cost Y1', 'Employee Cost Y2', 'PBT Y1', 'PBT Y2', 'PAT Y1',
       'PAT Y2', 'Tax1', 'Tax2' ,'FBTax1','FBTax2','DTax1','DTax2']] = financial_data[['CAPITALINE CODE', 'CO_NAME', '[Year End (Latest)]', '[Year End (Latest1)]',
       '[Net Sales (Latest)]', '[Net Sales (Latest1)]', '[PowerampFuel Cost (Latest)]', '[PowerampFuel Cost (Latest1)]',
       '[Employee Cost (Latest)]', '[Employee Cost (Latest1)]', '[Profit Before Tax (Latest)]', '[Profit Before Tax (Latest1)]', '[Reported Net Profit (Latest)]',
       '[Reported Net Profit (Latest1)]', '[Tax (Latest)]', '[Tax (Latest1)]','[Fringe Benefit tax (Latest)]','[Fringe Benefit tax (Latest1)]','[Deferred Tax (Latest)]','[Deferred Tax (Latest1)]']]


# In[29]:


pet_check_code['Total_Tax1'] = pet_check_code['Tax1'] + pet_check_code['FBTax1'] + pet_check_code['DTax1']
pet_check_code['Total_Tax2'] = pet_check_code['Tax2'] + pet_check_code['FBTax2'] + pet_check_code['DTax2']


# In[30]:


f_score_code = pd.DataFrame()
f_score_code[['CO_CODE', 'CO_NAME', '[Year End (Y1)]', '[Year End (Y2)]',
       '[Year End (Y3)]', '[Year (Y1)]', '[Year (Y2)]', '[Year (Y3)]',
        '[Net Sales (Y1)]', '[Net Sales (Y2)]',
       '[Net Sales (Y3)]', '[Operating Profit (Y1)]',
       '[Operating Profit (Y2)]', '[Operating Profit (Y3)]', '[Interest (Y1)]',
       '[Interest (Y2)]', '[Interest (Y3)]', '[Adjusted Net Profit (Y1)]',
       '[Adjusted Net Profit (Y2)]', '[Adjusted Net Profit (Y3)]',
       '[MODE (Y1)]', '[MODE (Y2)]', '[MODE (Y3)]',
       '[Total Shareholders Funds (Y1)]', '[Total Shareholders Funds (Y2)]',
       '[Total Shareholders Funds (Y3)]', '[Total Debt / Loan Funds (Y1)]',
       '[Total Debt / Loan Funds (Y2)]', '[Total Debt / Loan Funds (Y3)]',
       '[Cash and Bank Balance (Y1)]', '[Cash and Bank Balance (Y2)]',
       '[Cash and Bank Balance (Y3)]', '[Balance at Bank and Call Money (Y1)]',
       '[Balance at Bank and Call Money (Y2)]',
       '[Balance at Bank and Call Money (Y3)]', '[Total Assets (Y1)]',
       '[Total Assets (Y2)]', '[Total Assets (Y3)]',
       '[Net Cash from Operating Activities (Y1)]',
       '[Net Cash from Operating Activities (Y2)]',
       '[Net Cash from Operating Activities (Y3)]',
       '[Purchased of Fixed Assets (Y1)]', '[Purchased of Fixed Assets (Y2)]',
       '[Purchased of Fixed Assets (Y3)]', '[Sale of Fixed Assets (Y1)]',
       '[Sale of Fixed Assets (Y2)]', '[Sale of Fixed Assets (Y3)]',
       '[Capital Expenditure (Y1)]', '[Capital Expenditure (Y2)]',
       '[Capital Expenditure (Y3)]', '[capital WIP (Y1)]',
       '[capital WIP (Y2)]', '[capital WIP (Y3)]']] = financial_data[['CAPITALINE CODE', 'CO_NAME', '[Year End (Latest)]', '[Year End (Latest1)]',
       '[Year End (Latest2)]', '[Year (Latest)]', '[Year (Latest1)]', '[Year (Latest2)]',
        '[Net Sales (Latest)]', '[Net Sales (Latest1)]',
       '[Net Sales (Latest2)]', '[Operating Profit (Latest)]',
       '[Operating Profit (Latest1)]', '[Operating Profit (Latest2)]', '[Interest (Latest)]',
       '[Interest (Latest1)]', '[Interest (Latest2)]', '[Adjusted Net Profit (Latest)]',
       '[Adjusted Net Profit (Latest1)]', '[Adjusted Net Profit (Latest2)]',
       '[MODE (Latest)]', '[MODE (Latest1)]', '[MODE (Latest2)]',
       '[Total Shareholders Funds (Latest)]', '[Total Shareholders Funds (Latest1)]',
       '[Total Shareholders Funds (Latest2)]', '[Total Debt / Loan Funds (Latest)]',
       '[Total Debt / Loan Funds (Latest1)]', '[Total Debt / Loan Funds (Latest2)]',
       '[Cash and Bank Balance (Latest)]', '[Cash and Bank Balance (Latest1)]',
       '[Cash and Bank Balance (Latest2)]', '[Balance at Bank and Call Money (Latest)]',
       '[Balance at Bank and Call Money (Latest1)]',
       '[Balance at Bank and Call Money (Latest2)]', '[Total Assets (Latest)]',
       '[Total Assets (Latest1)]', '[Total Assets (Latest2)]',
       '[Net Cash from Operating Activities (Latest)]',
       '[Net Cash from Operating Activities (Latest1)]',
       '[Net Cash from Operating Activities (Latest2)]',
       '[Purchased of Fixed Assets (Latest)]', '[Purchased of Fixed Assets (Latest1)]',
       '[Purchased of Fixed Assets (Latest2)]', '[Sale of Fixed Assets (Latest)]',
       '[Sale of Fixed Assets (Latest1)]', '[Sale of Fixed Assets (Latest2)]',
       '[Capital Expenditure (Latest)]', '[Capital Expenditure (Latest1)]',
       '[Capital Expenditure (Latest2)]', '[capital WIP (Latest)]',
       '[capital WIP (Latest1)]', '[capital WIP (Latest2)]']]


# In[31]:


# adto_nse_f = pd.read_excel(os.path.join(raw_data_path,'NSE ADTO.xls'), sheet_name=None)
# adto_bse_f = pd.read_excel(os.path.join(raw_data_path,'BSE ADTO.xls'), sheet_name=None)


# In[32]:


# adto_nse = pd.DataFrame()
# for sheet in adto_nse_f.keys():
#     adto_nse = pd.concat([adto_nse,adto_nse_f[sheet]],ignore_index=True)


# In[33]:


# adto_bse = pd.DataFrame()
# for sheet in adto_bse_f.keys():
#     adto_bse = pd.concat([adto_bse,adto_bse_f[sheet]],ignore_index=True)


# In[34]:


promoter_pledge_sheet = pd.read_excel(os.path.join(raw_data_path,'Promoter Pledge.xls'))
nan_count_per_row = promoter_pledge_sheet.isnull().sum(axis=1)
promoter_pledge_sheet = promoter_pledge_sheet[nan_count_per_row <= 3]
promoter_pledge_sheet = promoter_pledge_sheet.reset_index(drop=True)


# In[35]:


impact_cost_both =  pd.read_excel(os.path.join(raw_data_path,'Impact Cost.xls'))
nan_count_per_row = impact_cost_both.isnull().sum(axis=1)
impact_cost_both = impact_cost_both[nan_count_per_row <= 3]
impact_cost_both = impact_cost_both.reset_index(drop=True)


# In[36]:


impact_cost_bse_sheet = pd.DataFrame()
impact_cost_bse_sheet[['CAPITALINE CODE', 'CO_NAME', '[ISIN No', '[BSE Scrip Code',
       '[Year End(YYYYMM)', '[Impact Cost']] = impact_cost_both[['CAPITALINE CODE', 'CO_NAME', '[ISIN No', 
       '[BSE Scrip Code (Latest)]','[Year End(YYYYMM) (Latest)]', '[Impact Cost (Latest)]']]


# In[37]:


impact_cost_bse_sheet = impact_cost_bse_sheet[~impact_cost_bse_sheet['[Impact Cost'].isna()].reset_index(drop=True)


# In[38]:


impact_cost_nse_sheet = pd.DataFrame()
impact_cost_nse_sheet[['CAPITALINE CODE', 'CO_NAME', '[ISIN No', '[NSE Symbol',
       '[Year End(YYYYMM)', '[Impact Cost']] = impact_cost_both[['CAPITALINE CODE', 'CO_NAME', '[ISIN No', 
       '[NSE Symbol (Latest)]','[Year End(YYYYMM) (Latest)]', '[Impact Cost (Latest)].1']]


# In[39]:


impact_cost_nse_sheet = impact_cost_nse_sheet[~impact_cost_nse_sheet['[Impact Cost'].isna()].reset_index(drop=True)


# In[40]:


# bse_cos = list(set(adto_bse['CO_NAME']))
# nse_cos = list(set(adto_nse['CO_NAME']))
# bse_nse_cos = bse_cos + nse_cos
# bse_nse_cos = sorted(list(set(bse_nse_cos)))


# In[41]:


# atdo_df = pd.DataFrame(columns=['CAPITALINE CODE','CO_NAME','Average'])
# atdo_df['CAPITALINE CODE'] = np.nan
# atdo_df['CO_NAME'] = bse_nse_cos
# for idx in range(0,len(atdo_df['CO_NAME'])) :
#     co_name = atdo_df['CO_NAME'][idx] 
    
#     co_code = bse_nse_group_sheet[bse_nse_group_sheet['CO_NAME'] == co_name ].reset_index(drop=True)
#     atdo_df['CAPITALINE CODE'][idx] = co_code['CAPITALINE CODE'][0]
    
#     nse_adto_req = adto_nse[adto_nse['CO_NAME'] == co_name ].reset_index(drop=True)
#     bse_adto_req = adto_bse[adto_bse['CO_NAME'] == co_name ].reset_index(drop=True)
#     nse_mean_adto = 0
#     bse_mean_adto = 0
#     total_dates_nse = len(nse_adto_req['[Net Turnover -Rs. Thousand'])
#     total_dates_bse = len(bse_adto_req['[Net Turnover -Rs. Thousand'])
#     max_dates_len = max(total_dates_nse,total_dates_bse)
#     if len(nse_adto_req) > 0 :
        
#         nse_mean_adto = ((nse_adto_req['[Net Turnover -Rs. Thousand'].sum())/max_dates_len)
#     if len(bse_adto_req) > 0 :
        
#         bse_mean_adto = ((bse_adto_req['[Net Turnover -Rs. Thousand'].sum())/max_dates_len)
#     bse_nse_mean_adto = nse_mean_adto+bse_mean_adto
#     atdo_df['Average'][idx] = bse_nse_mean_adto

# atdo_df['Amount Traded Rs.'] = atdo_df['Average']

# adto_sheet = atdo_df.copy()


# In[42]:


adto_df = bse_nse_group_sheet[['CAPITALINE CODE', 'CO_NAME', '[Company Long Name', '[ISIN No']]


# In[43]:


for month_fname in adto_months_req:
    path = os.path.join(adto_monthly_path,month_fname+'.xlsx')
    temp_df = pd.read_excel(path,header=1)
    adto_df[f'{month_fname}'] = 0
    for idx in range(0,len(adto_df['CAPITALINE CODE'])):
        co_code = adto_df['CAPITALINE CODE'][idx]
        temp_a = temp_df[temp_df['CAPITALINE CODE'] == co_code].reset_index()
        if len(temp_a)>0 :
            adto_df[f'{month_fname}'][idx] = temp_a['Total.1'][0]
adto_df['Median'] = np.nan
for idx in range(0,len(adto_df['CAPITALINE CODE'])):
    adto_df['Median'][idx] = median(adto_df[adto_months_req].loc[idx])

# print('ADTO done')
# adto_df.to_csv('ADTO.csv')
# import sys
# sys.exit(0)
# In[44]:


adto_df['Amount Traded Rs.'] = adto_df['Median']
adto_sheet = adto_df.copy()


# In[45]:


sales_pat_sheet = pd.DataFrame()
sales_pat_sheet = financial_data[['CAPITALINE CODE','CO_NAME','[Year End (Latest)]','[Net Sales (Latest)]','[Adjusted Net Profit (Latest)]']]


# In[46]:


close_price_sheet_bse = pd.read_excel(os.path.join(raw_data_path,'BSE latest traded date.xls'))
close_price_sheet_nse = pd.read_excel(os.path.join(raw_data_path,'NSE latest traded date.xls'))


# In[47]:


risk_rating_nonfno_isin = pd.DataFrame()
risk_rating_nonfno_isin = bse_nse_group_sheet[['CAPITALINE CODE', 'CO_NAME', '[Company Long Name', '[ISIN No']]


# In[48]:


prev_month_data = pd.read_excel(prev_month_excel,sheet_name=None)


# In[49]:


prev_month_data_fno_sheet = prev_month_data['Risk Rating_F&O']
prev_month_data_non_fno_sheet = prev_month_data['Risk Rating_Non F&O']


# In[50]:


prev_month_data_data_points = prev_month_data['Data_Points_Curr']


# In[51]:


## sheet 1


# In[52]:


# fno_list_sheet = all_sheets['Fno List']
# Networth_sheet = all_sheets['Networth']
# latest_mcap_sheet = all_sheets['latest Mcap']
# bse_nse_group_sheet = all_sheets['BSE & NSE Group']
# nifty_500_sheet = all_sheets['Nifty 500']
# pet_check_sheet = all_sheets['PET Check']
# f_score_sheet = all_sheets['F - Score']
# adto_sheet = all_sheets['ADTO']
# promoter_pledge_sheet = all_sheets['Promoter Pledge']
# impact_cost_bse_sheet = all_sheets['Impact Cost BSE']
# impact_cost_nse_sheet = all_sheets['Impact Cost NSE']
# sales_pat_sheet = all_sheets['Sales PAT']
# close_price_sheet = all_sheets['Close Price']
# risk_rating_nonfno_isin = all_sheets['Risk Rating_Non F&O']


# In[ ]:





# In[ ]:





# In[53]:


manual_rank = pd.read_excel(manual_reviewed_stocks_path,sheet_name=None)
poor_downgrd_rated_df = manual_rank['Poor Rated']
poor_to_avg_rated_df = manual_rank['Poor to Average Rated']
good_bc_to_avg_rated_df = manual_rank['Any other Rating to Average']


# In[54]:


# bse_var_sheet = all_sheets['BSE VAR']
# bse_var_sheet['Co name_lower'] = ''
# for idx in range(0,len(bse_var_sheet['BSE code'])):
#     try:
#         bse_var_sheet['Co name_lower'][idx] = bse_var_sheet['Co name'][idx].lower()
#     except:
#         bse_var_sheet['Co name_lower'][idx] = ''


# In[55]:


# pet_check_sheet.columns


# In[56]:


# pet_check_cols = pd.read_excel(r'C:\cltemp\Book400.xls')
# pet_check_cols.columns


# In[57]:


# pet_check_code[['CAPITALINE CODE', 'CO_NAME', 'YearEnd1', 'YearEnd2',
#        'Sales Y1', 'Sales Y2', 'Power & Fuel Cost Y1', 'Power & Fuel Cost Y2',
#        'Employee Cost Y1', 'Employee Cost Y2', 'PBT Y1', 'PBT Y2', 'PAT Y1',
#        'PAT Y2', 'Tax1', 'Tax2','FBTax1','FBTax2','DTax1','DTax2']] = pet_check_sheet[['CAPITALINE CODE', 'CO_NAME', '[Year End (Latest)]',
#        '[Year End (Latest1)]', '[Net Sales (Latest)]', '[Net Sales (Latest1)]',
#        '[Power & Fuel Cost (Latest)]', '[Power & Fuel Cost (Latest1)]',
#        '[Employee Cost (Latest)]', '[Employee Cost (Latest1)]',
#        '[Profit Before Tax (Latest)]', '[Profit Before Tax (Latest1)]',
#        '[Reported Net Profit (Latest)]',
#        '[Reported Net Profit (Latest1)]',
#        '[Tax (Latest)]', '[Tax (Latest1)]', '[Fringe Benefit tax (Latest)]',
#        '[Fringe Benefit tax (Latest1)]', '[Deferred Tax (Latest)]',
#        '[Deferred Tax (Latest1)]']]


# In[58]:


# pet_check_code = pd.DataFrame()
# pet_check_code[['CO_CODE', 'CO_NAME', 'YearEnd1', 'YearEnd2',
#        'Sales Y1', 'Sales Y2', 'Power & Fuel Cost Y1', 'Power & Fuel Cost Y2',
#        'Employee Cost Y1', 'Employee Cost Y2', 'PBT Y1', 'PBT Y2', 'PAT Y1',
#        'PAT Y2', 'Total_Tax1', 'Total_Tax2']] = pet_check_sheet[['CAPITALINE CODE', 'CO_NAME', 'Y1', 'Y2',
#        'Sales Y1', 'Sales Y2', 'Power & Fuel Cost Y1', 'Power & Fuel Cost Y2',
#        'Employee Cost Y1', 'Employee Cost Y2', 'PBT Y1', 'PBT Y2', 'PAT Y1',
#        'PAT Y2', 'Tax1', 'Tax2']]


# In[59]:


pet_check_code['Y1'] = np.nan
pet_check_code['Y2'] = np.nan
pet_check_code['Power Y1 %'] = np.nan
pet_check_code['Power Y2 %'] = np.nan
pet_check_code['Employee  Y1 %'] = np.nan
pet_check_code['Employee  Y2 %'] = np.nan
pet_check_code['Tax Y1%'] = np.nan
pet_check_code['Tax Y2%'] = np.nan

for idx in range(0,len(pet_check_code['CO_CODE'])):
    if not pd.isna(pet_check_code['YearEnd1'][idx]) :
        pet_check_code['Y1'][idx]= int(str(int(pet_check_code['YearEnd1'][idx]))[:4])
    if not pd.isna(pet_check_code['YearEnd2'][idx]) :
        pet_check_code['Y2'][idx]= int(str(int(pet_check_code['YearEnd2'][idx]))[:4])
    
    if ( (pd.isna(pet_check_code['Sales Y1'][idx])) | ( pet_check_code['Sales Y1'][idx]== 0)):
        ...
    else:
        pet_check_code['Power Y1 %'][idx] = pet_check_code['Power & Fuel Cost Y1'][idx] / pet_check_code['Sales Y1'][idx]
        pet_check_code['Employee  Y1 %'][idx] = pet_check_code['Employee Cost Y1'][idx] / pet_check_code['Sales Y1'][idx]
    
    
    if ( (pd.isna(pet_check_code['Sales Y2'][idx])) | ( pet_check_code['Sales Y2'][idx]== 0)):
        ...
    else:
        pet_check_code['Power Y2 %'][idx] = pet_check_code['Power & Fuel Cost Y2'][idx] / pet_check_code['Sales Y2'][idx]
        pet_check_code['Employee  Y2 %'][idx] = pet_check_code['Employee Cost Y2'][idx] / pet_check_code['Sales Y2'][idx]
        
    if ( (pd.isna(pet_check_code['PBT Y1'][idx])) | ( pet_check_code['PBT Y1'][idx]== 0)):
        ...
    else:
        if ( (pet_check_code['PBT Y1'][idx]<0) & (pet_check_code['Total_Tax1'][idx]<0) ):
            pet_check_code['Tax Y1%'][idx] = ((pet_check_code['Total_Tax1'][idx]/pet_check_code['PBT Y1'][idx])*-1)
        else:
            pet_check_code['Tax Y1%'][idx] = (pet_check_code['Total_Tax1'][idx]/pet_check_code['PBT Y1'][idx])
    
    
    if ( (pd.isna(pet_check_code['PBT Y2'][idx])) | ( pet_check_code['PBT Y2'][idx]== 0)):
        ...
    else:
        if ( (pet_check_code['PBT Y2'][idx]<0) & (pet_check_code['Total_Tax2'][idx]<0) ):
            pet_check_code['Tax Y2%'][idx] = ((pet_check_code['Total_Tax2'][idx]/pet_check_code['PBT Y2'][idx])*-1)
        else:
            pet_check_code['Tax Y2%'][idx] = (pet_check_code['Total_Tax2'][idx]/pet_check_code['PBT Y2'][idx])


# In[60]:


pet_check_code['Pow1_chk'] = np.nan
pet_check_code['Pow2_chk'] = np.nan
pet_check_code['Emp1_chk'] = np.nan
pet_check_code['Emp2_chk'] = np.nan
pet_check_code['Tax1_chk'] = np.nan
pet_check_code['Tax2_chk'] = np.nan
pet_check_code['pow_emp_sum'] = np.nan
pet_check_code['tax_emp_sum'] = np.nan
pet_check_code['pow_tax_sum'] = np.nan
pet_check_code['max_sum'] = np.nan
pet_check_code['Final'] = ''
for idx in range(0,len(pet_check_code['CO_CODE'])):
    if not pd.isna(pet_check_code['Power Y1 %'][idx]):
        if pet_check_code['Power Y1 %'][idx] > 0.01 :
            pet_check_code['Pow1_chk'][idx] = 1
        else:
            pet_check_code['Pow1_chk'][idx] = 0
    else:
        pet_check_code['Pow1_chk'][idx] = 0
    
    
    if not pd.isna(pet_check_code['Power Y2 %'][idx]):
        if pet_check_code['Power Y2 %'][idx] > 0.01 :
            pet_check_code['Pow2_chk'][idx] = 1
        else:
            pet_check_code['Pow2_chk'][idx] = 0
    else:
        pet_check_code['Pow2_chk'][idx] = 0
    
    
    if not pd.isna(pet_check_code['Employee  Y1 %'][idx]):
        if pet_check_code['Employee  Y1 %'][idx] > 0.01 :
            pet_check_code['Emp1_chk'][idx] = 1
        else:
            pet_check_code['Emp1_chk'][idx] = 0
    else:
        pet_check_code['Emp1_chk'][idx] = 0
    
    if not pd.isna(pet_check_code['Employee  Y2 %'][idx]):
        if pet_check_code['Employee  Y2 %'][idx] > 0.01 :
            pet_check_code['Emp2_chk'][idx] = 1
        else:
            pet_check_code['Emp2_chk'][idx] = 0
    else:
        pet_check_code['Emp2_chk'][idx] = 0
    
    
    if not pd.isna(pet_check_code['Tax Y1%'][idx]):
        if pet_check_code['Tax Y1%'][idx] > 0.15 :
            pet_check_code['Tax1_chk'][idx] = 1
        else:
            pet_check_code['Tax1_chk'][idx] = 0
    else:
        pet_check_code['Tax1_chk'][idx] = 0
        
    if not pd.isna(pet_check_code['Tax Y2%'][idx]):
        if pet_check_code['Tax Y2%'][idx] > 0.15 :
            pet_check_code['Tax2_chk'][idx] = 1
        else:
            pet_check_code['Tax2_chk'][idx] = 0
    else:
        pet_check_code['Tax2_chk'][idx] = 0
        
    pow_emp_sum = pet_check_code["Pow1_chk"][idx] + pet_check_code["Pow2_chk"][idx] + pet_check_code["Emp1_chk"][idx] + pet_check_code["Emp2_chk"][idx]
    tax_emp_sum = pet_check_code["Emp1_chk"][idx] + pet_check_code["Emp2_chk"][idx] + pet_check_code["Tax1_chk"][idx] + pet_check_code["Tax2_chk"][idx]
    pow_tax_sum = pet_check_code["Pow1_chk"][idx] + pet_check_code["Pow2_chk"][idx] + pet_check_code["Tax1_chk"][idx] + pet_check_code["Tax2_chk"][idx]
    
    max_sum = max(pow_emp_sum, tax_emp_sum, pow_tax_sum)
    pet_check_code['pow_emp_sum'][idx] = pow_emp_sum
    pet_check_code['tax_emp_sum'][idx] = tax_emp_sum
    pet_check_code['pow_tax_sum'][idx] = pow_tax_sum
    pet_check_code['max_sum'][idx] = max_sum
    if pet_check_code['max_sum'][idx] == 4 :
        pet_check_code['Final'][idx] = 'PET Good'


# In[61]:


# f_score_sheet.columns


# In[62]:


# f_score_code = pd.DataFrame()
# f_score_code[['CO_CODE', 'CO_NAME', '[Year End (Y1)]', '[Year End (Y2)]',
#        '[Year End (Y3)]', '[Year (Y1)]', '[Year (Y2)]', '[Year (Y3)]',
#         '[Net Sales (Y1)]', '[Net Sales (Y2)]',
#        '[Net Sales (Y3)]', '[Operating Profit (Y1)]',
#        '[Operating Profit (Y2)]', '[Operating Profit (Y3)]', '[Interest (Y1)]',
#        '[Interest (Y2)]', '[Interest (Y3)]', '[Adjusted Net Profit (Y1)]',
#        '[Adjusted Net Profit (Y2)]', '[Adjusted Net Profit (Y3)]',
#        '[MODE (Y1)]', '[MODE (Y2)]', '[MODE (Y3)]',
#        '[Total Shareholders Funds (Y1)]', '[Total Shareholders Funds (Y2)]',
#        '[Total Shareholders Funds (Y3)]', '[Total Debt / Loan Funds (Y1)]',
#        '[Total Debt / Loan Funds (Y2)]', '[Total Debt / Loan Funds (Y3)]',
#        '[Cash and Bank Balance (Y1)]', '[Cash and Bank Balance (Y2)]',
#        '[Cash and Bank Balance (Y3)]', '[Balance at Bank and Call Money (Y1)]',
#        '[Balance at Bank and Call Money (Y2)]',
#        '[Balance at Bank and Call Money (Y3)]', '[Total Assets (Y1)]',
#        '[Total Assets (Y2)]', '[Total Assets (Y3)]',
#        '[Net Cash from Operating Activities (Y1)]',
#        '[Net Cash from Operating Activities (Y2)]',
#        '[Net Cash from Operating Activities (Y3)]',
#        '[Purchased of Fixed Assets (Y1)]', '[Purchased of Fixed Assets (Y2)]',
#        '[Purchased of Fixed Assets (Y3)]', '[Sale of Fixed Assets (Y1)]',
#        '[Sale of Fixed Assets (Y2)]', '[Sale of Fixed Assets (Y3)]',
#        '[Capital Expenditure (Y1)]', '[Capital Expenditure (Y2)]',
#        '[Capital Expenditure (Y3)]', '[capital WIP (Y1)]',
#        '[capital WIP (Y2)]', '[capital WIP (Y3)]']] = f_score_sheet[['CAPITALINE CODE', 'CO_NAME', '[Year End (Y1)]', '[Year End (Y2)]',
#        '[Year End (Y3)]', '[Year (Y1)]', '[Year (Y2)]', '[Year (Y3)]',
#         '[Net Sales (Y1)]', '[Net Sales (Y2)]',
#        '[Net Sales (Y3)]', '[Operating Profit (Y1)]',
#        '[Operating Profit (Y2)]', '[Operating Profit (Y3)]', '[Interest (Y1)]',
#        '[Interest (Y2)]', '[Interest (Y3)]', '[Adjusted Net Profit (Y1)]',
#        '[Adjusted Net Profit (Y2)]', '[Adjusted Net Profit (Y3)]',
#        '[MODE (Y1)]', '[MODE (Y2)]', '[MODE (Y3)]',
#        '[Total Shareholders Funds (Y1)]', '[Total Shareholders Funds (Y2)]',
#        '[Total Shareholders Funds (Y3)]', '[Total Debt / Loan Funds (Y1)]',
#        '[Total Debt / Loan Funds (Y2)]', '[Total Debt / Loan Funds (Y3)]',
#        '[Cash and Bank Balance (Y1)]', '[Cash and Bank Balance (Y2)]',
#        '[Cash and Bank Balance (Y3)]', '[Balance at Bank and Call Money (Y1)]',
#        '[Balance at Bank and Call Money (Y2)]',
#        '[Balance at Bank and Call Money (Y3)]', '[Total Assets (Y1)]',
#        '[Total Assets (Y2)]', '[Total Assets (Y3)]',
#        '[Net Cash from Operating Activities (Y1)]',
#        '[Net Cash from Operating Activities (Y2)]',
#        '[Net Cash from Operating Activities (Y3)]',
#        '[Purchased of Fixed Assets (Y1)]', '[Purchased of Fixed Assets (Y2)]',
#        '[Purchased of Fixed Assets (Y3)]', '[Sale of Fixed Assets (Y1)]',
#        '[Sale of Fixed Assets (Y2)]', '[Sale of Fixed Assets (Y3)]',
#        '[Capital Expenditure (Y1)]', '[Capital Expenditure (Y2)]',
#        '[Capital Expenditure (Y3)]', '[capital WIP (Y1)]',
#        '[capital WIP (Y2)]', '[capital WIP (Y3)]']]


# In[63]:


# f_score_code.columns


# In[64]:


f_score_code['Sales Growth Y1%'] = np.nan
f_score_code['Sales Growth Y2%'] = np.nan
f_score_code['EBITDA Margin% Y1'] = np.nan
f_score_code['EBITDA Margin% Y2'] = np.nan
for idx in range(0,len(f_score_code['CO_CODE'])):
    
    if ( (pd.isna(f_score_code['[Net Sales (Y1)]'][idx])) | ( f_score_code['[Net Sales (Y1)]'][idx]== 0)):
        ...
    else:
        f_score_code['EBITDA Margin% Y1'][idx] = f_score_code['[Operating Profit (Y1)]'][idx]/f_score_code['[Net Sales (Y1)]'][idx]
    
    
    if ( (pd.isna(f_score_code['[Net Sales (Y2)]'][idx])) | ( f_score_code['[Net Sales (Y2)]'][idx]== 0)):
        ...
    else:
        f_score_code['Sales Growth Y1%'][idx] = ((f_score_code['[Net Sales (Y1)]'][idx] / f_score_code['[Net Sales (Y2)]'][idx]) -1 )
        f_score_code['EBITDA Margin% Y2'][idx] = f_score_code['[Operating Profit (Y2)]'][idx]/f_score_code['[Net Sales (Y2)]'][idx]
    
    
    if ( (pd.isna(f_score_code['[Net Sales (Y3)]'][idx])) | ( f_score_code['[Net Sales (Y3)]'][idx]== 0)):
        ...
    else:
        f_score_code['Sales Growth Y2%'][idx] = ((f_score_code['[Net Sales (Y2)]'][idx] / f_score_code['[Net Sales (Y3)]'][idx]) -1 )


# In[65]:


f_score_code['D/E Y1'] = np.nan
f_score_code['D/E Y2'] = np.nan
f_score_code['AT Y1'] = np.nan
f_score_code['AT Y2'] = np.nan
f_score_code['FCF Y1'] = np.nan
f_score_code['FCF Y2'] = np.nan
f_score_code['ICR Y1'] = np.nan
f_score_code['ICR Y2'] = np.nan
for idx in range(0,len(f_score_code['CO_CODE'])):
    if ( (pd.isna(f_score_code['[Total Shareholders Funds (Y1)]'][idx])) | ( f_score_code['[Total Shareholders Funds (Y1)]'][idx]== 0)):
        ...
    else:
        f_score_code['D/E Y1'][idx] = (f_score_code['[Total Debt / Loan Funds (Y1)]'][idx]-f_score_code['[Cash and Bank Balance (Y1)]'][idx]-f_score_code['[Balance at Bank and Call Money (Y1)]'][idx])/f_score_code['[Total Shareholders Funds (Y1)]'][idx]
    
    if ( (pd.isna(f_score_code['[Total Shareholders Funds (Y2)]'][idx])) | ( f_score_code['[Total Shareholders Funds (Y2)]'][idx]== 0)):
        ...
    else:
        f_score_code['D/E Y2'][idx] = (f_score_code['[Total Debt / Loan Funds (Y2)]'][idx]-f_score_code['[Cash and Bank Balance (Y2)]'][idx]-f_score_code['[Balance at Bank and Call Money (Y2)]'][idx])/f_score_code['[Total Shareholders Funds (Y2)]'][idx]
    
    if ( (pd.isna(f_score_code['[Total Assets (Y1)]'][idx])) | ( f_score_code['[Total Assets (Y1)]'][idx]== 0)):
        ...
    else:
        f_score_code['AT Y1'][idx] = f_score_code['[Net Sales (Y1)]'][idx]/f_score_code['[Total Assets (Y1)]'][idx]
    
    if ( (pd.isna(f_score_code['[Total Assets (Y2)]'][idx])) | ( f_score_code['[Total Assets (Y2)]'][idx]== 0)):
        ...
    else:
        f_score_code['AT Y2'][idx] = f_score_code['[Net Sales (Y2)]'][idx]/f_score_code['[Total Assets (Y2)]'][idx]
    
    f_score_code['FCF Y1'][idx] = f_score_code['[Net Cash from Operating Activities (Y1)]'][idx] + f_score_code['[Purchased of Fixed Assets (Y1)]'][idx]+ f_score_code['[Sale of Fixed Assets (Y1)]'][idx]+ f_score_code['[capital WIP (Y1)]'][idx]+ f_score_code['[Capital Expenditure (Y1)]'][idx]
    f_score_code['FCF Y2'][idx] = f_score_code['[Net Cash from Operating Activities (Y2)]'][idx] + f_score_code['[Purchased of Fixed Assets (Y2)]'][idx]+ f_score_code['[Sale of Fixed Assets (Y2)]'][idx]+ f_score_code['[capital WIP (Y2)]'][idx]+ f_score_code['[Capital Expenditure (Y2)]'][idx]
    
    if ( (pd.isna(f_score_code['[Interest (Y1)]'][idx])) | ( f_score_code['[Interest (Y1)]'][idx]== 0)):
        f_score_code['ICR Y1'][idx] = 0
    else:
        f_score_code['ICR Y1'][idx] = f_score_code['[Operating Profit (Y1)]'][idx]/f_score_code['[Interest (Y1)]'][idx]
    
    if ( (pd.isna(f_score_code['[Interest (Y2)]'][idx])) | ( f_score_code['[Interest (Y2)]'][idx]== 0)):
        f_score_code['ICR Y2'][idx] = 0
    else:
        f_score_code['ICR Y2'][idx] = f_score_code['[Operating Profit (Y2)]'][idx]/f_score_code['[Interest (Y2)]'][idx]
    


# In[66]:


f_score_code['Sal1_chk'] = np.nan
f_score_code['Sal2_chk'] = np.nan
f_score_code['EBITDA1_chk'] = np.nan
f_score_code['EBITDA2_chk'] = np.nan
f_score_code['PAT1_chk'] = np.nan
f_score_code['PAT2_chk'] = np.nan
f_score_code['DE1_chk'] = np.nan
f_score_code['DE2_chk'] = np.nan
f_score_code['ICR1_chk'] = np.nan
f_score_code['ICR2_chk'] = np.nan
f_score_code['AT1_chk'] = np.nan
f_score_code['AT2_chk'] = np.nan
f_score_code['FCF1_chk'] = np.nan
f_score_code['FCF2_chk'] = np.nan
f_score_code['Y1 Score'] = np.nan
f_score_code['Y2 Score'] = np.nan
f_score_code['Final Score'] = np.nan 
f_score_code['Final Check'] = ''
for idx in range(0,len(f_score_code['CO_CODE'])):
    if not pd.isna(f_score_code['Sales Growth Y1%'][idx]):
        if f_score_code['Sales Growth Y1%'][idx] > 0 :
            f_score_code['Sal1_chk'][idx] = 1
        else:
            f_score_code['Sal1_chk'][idx] = 0
    else:
        f_score_code['Sal1_chk'][idx] = 0
    
    
    if not pd.isna(f_score_code['Sales Growth Y2%'][idx]):
        if f_score_code['Sales Growth Y2%'][idx] > 0 :
            f_score_code['Sal2_chk'][idx] = 1
        else:
            f_score_code['Sal2_chk'][idx] = 0
    else:
        f_score_code['Sal2_chk'][idx] = 0
    
    
    if not pd.isna(f_score_code['EBITDA Margin% Y1'][idx]):
        if f_score_code['EBITDA Margin% Y1'][idx] > 0 :
            f_score_code['EBITDA1_chk'][idx] = 1
        else:
            f_score_code['EBITDA1_chk'][idx] = 0
    else:
        f_score_code['EBITDA1_chk'][idx] = 0
    
    if not pd.isna(f_score_code['EBITDA Margin% Y2'][idx]):
        if f_score_code['EBITDA Margin% Y2'][idx] > 0 :
            f_score_code['EBITDA2_chk'][idx] = 1
        else:
            f_score_code['EBITDA2_chk'][idx] = 0
    else:
        f_score_code['EBITDA2_chk'][idx] = 0
    
    
    if not pd.isna(f_score_code['[Adjusted Net Profit (Y1)]'][idx]):
        if f_score_code['[Adjusted Net Profit (Y1)]'][idx] > 0 :
            f_score_code['PAT1_chk'][idx] = 1
        else:
            f_score_code['PAT1_chk'][idx] = 0
    else:
        f_score_code['PAT1_chk'][idx] = 0
        
    if not pd.isna(f_score_code['[Adjusted Net Profit (Y2)]'][idx]):
        if f_score_code['[Adjusted Net Profit (Y2)]'][idx] > 0 :
            f_score_code['PAT2_chk'][idx] = 1
        else:
            f_score_code['PAT2_chk'][idx] = 0
    else:
        f_score_code['PAT2_chk'][idx] = 0
    
    
    if not pd.isna(f_score_code['D/E Y1'][idx]):
        if f_score_code['D/E Y1'][idx] <= 0.5 :
            f_score_code['DE1_chk'][idx] = 1
        else:
            f_score_code['DE1_chk'][idx] = 0
    else:
        f_score_code['DE1_chk'][idx] = 0
    
    if not pd.isna(f_score_code['D/E Y2'][idx]):
        if f_score_code['D/E Y2'][idx] <= 0.5 :
            f_score_code['DE2_chk'][idx] = 1
        else:
            f_score_code['DE2_chk'][idx] = 0
    else:
        f_score_code['DE2_chk'][idx] = 0
    
    if not pd.isna(f_score_code['ICR Y1'][idx]):
        if f_score_code['ICR Y1'][idx] >= 1 :
            f_score_code['ICR1_chk'][idx] = 1
        else:
            f_score_code['ICR1_chk'][idx] = 0
    else:
        f_score_code['ICR1_chk'][idx] = 0
    
    if not pd.isna(f_score_code['ICR Y2'][idx]):
        if f_score_code['ICR Y2'][idx] >= 1 :
            f_score_code['ICR2_chk'][idx] = 1
        else:
            f_score_code['ICR2_chk'][idx] = 0
    else:
        f_score_code['ICR2_chk'][idx] = 0
    
    
    if not pd.isna(f_score_code['AT Y1'][idx]):
        if f_score_code['AT Y1'][idx] >= 1 :
            f_score_code['AT1_chk'][idx] = 1
        else:
            f_score_code['AT1_chk'][idx] = 0
    else:
        f_score_code['AT1_chk'][idx] = 0
    
    if not pd.isna(f_score_code['AT Y2'][idx]):
        if f_score_code['AT Y2'][idx] >= 1 :
            f_score_code['AT2_chk'][idx] = 1
        else:
            f_score_code['AT2_chk'][idx] = 0
    else:
        f_score_code['AT2_chk'][idx] = 0
    
    
    if not pd.isna(f_score_code['FCF Y1'][idx]):
        if f_score_code['FCF Y1'][idx] > 0 :
            f_score_code['FCF1_chk'][idx] = 1
        else:
            f_score_code['FCF1_chk'][idx] = 0
    else:
        f_score_code['FCF1_chk'][idx] = 0
    
    if not pd.isna(f_score_code['FCF Y2'][idx]):
        if f_score_code['FCF Y2'][idx] > 0 :
            f_score_code['FCF2_chk'][idx] = 1
        else:
            f_score_code['FCF2_chk'][idx] = 0
    else:
        f_score_code['FCF2_chk'][idx] = 0
    
    f_score_code['Y1 Score'][idx] = sum([ f_score_code['Sal1_chk'][idx],f_score_code['EBITDA1_chk'][idx],f_score_code['PAT1_chk'][idx],f_score_code['DE1_chk'][idx],f_score_code['ICR1_chk'][idx],f_score_code['AT1_chk'][idx],f_score_code['FCF1_chk'][idx] ])
    f_score_code['Y2 Score'][idx] = sum([ f_score_code['Sal2_chk'][idx],f_score_code['EBITDA2_chk'][idx],f_score_code['PAT2_chk'][idx],f_score_code['DE2_chk'][idx],f_score_code['ICR2_chk'][idx],f_score_code['AT2_chk'][idx],f_score_code['FCF2_chk'][idx] ])
    f_score_code['Final Score'][idx] = sum([f_score_code['Y1 Score'][idx],f_score_code['Y2 Score'][idx]])
    if ((f_score_code['Final Score'][idx] >=7) & (f_score_code['Y1 Score'][idx]>=3)) :
        f_score_code['Final Check'][idx] = 'F-Check Passed'


# In[67]:


# f_score_code.to_csv('f_score_code.csv')


# In[68]:


impact_cost = pd.DataFrame()


# In[69]:


impact_cost[['CO_CODE', 'CO_NAME']] = pet_check_code[['CO_CODE', 'CO_NAME']]


# In[70]:


# impact_cost_bse = impact_cost_bse_sheet.groupby('CAPITALINE CODE').mean('[Impact Cost').reset_index()
# impact_cost_nse = impact_cost_nse_sheet.groupby('CAPITALINE CODE').mean('[Impact Cost').reset_index()
impact_cost_bse = impact_cost_bse_sheet.copy()
impact_cost_nse = impact_cost_nse_sheet.copy()


# In[71]:


impact_cost['NSE'] = np.nan
impact_cost['BSE'] = np.nan
impact_cost['Min of NSE/BSE'] = np.nan

for idx in range(0,len(impact_cost['CO_CODE'])):
    temp_bse_ic = impact_cost_bse[impact_cost_bse['CAPITALINE CODE'] == impact_cost['CO_CODE'][idx] ].reset_index()
    temp_nse_ic = impact_cost_nse[impact_cost_nse['CAPITALINE CODE'] == impact_cost['CO_CODE'][idx] ].reset_index()
    if len(temp_bse_ic) > 0 :
        impact_cost['BSE'][idx] = temp_bse_ic['[Impact Cost'][0]
    
    if len(temp_nse_ic) > 0 :
        impact_cost['NSE'][idx] = temp_nse_ic['[Impact Cost'][0]
    if (( not pd.isna(impact_cost['BSE'][idx])) & ( not pd.isna(impact_cost['NSE'][idx]))) :
        impact_cost['Min of NSE/BSE'][idx] = min( [impact_cost['BSE'][idx] ,impact_cost['NSE'][idx]] )
    elif ( pd.isna(impact_cost['BSE'][idx])) :
        impact_cost['Min of NSE/BSE'][idx] = impact_cost['NSE'][idx]
    elif ( pd.isna(impact_cost['NSE'][idx])) :
        impact_cost['Min of NSE/BSE'][idx] = impact_cost['BSE'][idx]
    else:
        impact_cost['Min of NSE/BSE'][idx] = np.nan


# In[72]:


risk_rating_fno = pd.DataFrame()
risk_rating_fno[['CO_CODE','CO_NAME']] = fno_list_sheet[['CAPITALINE CODE','CO_NAME']]
risk_rating_fno['F&O'] = 'F&O'
risk_rating_fno['[Total Shareholders Funds (Latest)]'] = np.nan
risk_rating_fno['Mcap'] = np.nan
risk_rating_fno['BSE Series'] = ''
risk_rating_fno['NSE Series'] = ''
risk_rating_fno['Nifty 500'] = ''
risk_rating_fno['PET Check'] = ''
risk_rating_fno['F-Score'] = ''
risk_rating_fno['F-Score'] = ''
risk_rating_fno['ADTO in Cr'] = np.nan
risk_rating_fno['Promoter Pledge%'] = np.nan
risk_rating_fno['Impact Cost'] = np.nan
risk_rating_fno['Rev'] = np.nan
 #new added 24-09-2024
risk_rating_fno['[ISIN No'] = ''
risk_rating_fno['Percent_Funding'] = np.nan
# In[73]:


for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    
    temp_nw = Networth_sheet[risk_rating_fno['CO_CODE'][idx] == Networth_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_nw)>0 :
        risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] = temp_nw['[Networth (Latest)]'][0]
    
    temp_mcap = latest_mcap_sheet[risk_rating_fno['CO_CODE'][idx] == latest_mcap_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_mcap) > 0 :
        risk_rating_fno['Mcap'][idx] = temp_mcap['[Market Cap (Latest)]'][0]
    
    temp_bsensegrp = bse_nse_group_sheet[risk_rating_fno['CO_CODE'][idx] == bse_nse_group_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_bsensegrp)>0:
        risk_rating_fno['BSE Series'][idx] = temp_bsensegrp['[BSE Group'][0]
        risk_rating_fno['NSE Series'][idx] = temp_bsensegrp['[NSE Series'][0]
        risk_rating_fno['[ISIN No'][idx] = temp_bsensegrp['[ISIN No'][0]
    
    temp_nifty500 = nifty_500_sheet[risk_rating_fno['CO_CODE'][idx] == nifty_500_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_nifty500) > 0 :
        risk_rating_fno['Nifty 500'][idx] = temp_nifty500['Index'][0]
    
    temp_pet_chk = pet_check_code[pet_check_code['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_pet_chk)>0:
        risk_rating_fno['PET Check'][idx] = temp_pet_chk['Final'][0]
    
    temp_fscore = f_score_code[f_score_code['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_fscore)>0:
        risk_rating_fno['F-Score'][idx] = temp_fscore['Final Check'][0]
    
    temp_adto = adto_sheet[adto_sheet['CAPITALINE CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_adto)>0:
        risk_rating_fno['ADTO in Cr'][idx] = temp_adto['Amount Traded Rs.'][0]/pow(10,7)
    
    temp_promoter_pledge = promoter_pledge_sheet[promoter_pledge_sheet['CAPITALINE CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_promoter_pledge)>0:
        risk_rating_fno['Promoter Pledge%'][idx] = temp_promoter_pledge['[Total of Promoter and Group (Latest)]'][0]
    
    temp_impact_cost = impact_cost[impact_cost['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_impact_cost)>0:
        risk_rating_fno['Impact Cost'][idx] = temp_impact_cost['Min of NSE/BSE'][0]
    
    temp_sales_pat = sales_pat_sheet[sales_pat_sheet['CAPITALINE CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_sales_pat)>0:
        risk_rating_fno['Rev'][idx] = temp_sales_pat['[Net Sales (Latest)]'][0]
    
     #new added 24-09-2024
    temp_funding = funding_df[funding_df['ISIN'] == risk_rating_fno['[ISIN No'][idx]].reset_index()
    if len(temp_funding) > 0 :
        risk_rating_fno['Percent_Funding'][idx] = temp_funding['% Funding'][0]
    
    


# In[74]:


latest_3_months_adto = adto_months_req[3:]
latest_3_months_adto_cols = []
for col in latest_3_months_adto:
    latest_3_months_adto_cols.append(f'{col}_ADTO')

risk_rating_fno[latest_3_months_adto_cols] = 0
# risk_rating_fno['3M_ADTO_count'] = np.nan
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    temp_adto_df = adto_df[adto_df['CAPITALINE CODE'] ==risk_rating_fno['CO_CODE'][idx] ].reset_index(drop=True)
    if len(temp_adto_df) > 0 :
        for col in latest_3_months_adto:
            risk_rating_fno[f'{col}_ADTO'][idx] = temp_adto_df[col][0]/pow(10,7)


# In[ ]:





# In[ ]:





# In[ ]:





# In[75]:


def adto_count(risk_rating_fno , adto_in_cr):
#     flag = 0
    adto_count = 0
    risk_rating_fno_row = risk_rating_fno.reset_index(drop=True)
    for col in latest_3_months_adto_cols:
        if risk_rating_fno_row[col][0] > adto_in_cr:
            adto_count+=1
#     if adto_count == 3 :
#         flag = 1
    return adto_count


# In[76]:


def adto_count_poor(risk_rating_fno , adto_in_cr):
#     flag = 0
    adto_count = 0
    risk_rating_fno_row = risk_rating_fno.reset_index(drop=True)
    for col in latest_3_months_adto_cols:
        if risk_rating_fno_row[col][0] < adto_in_cr:
            adto_count+=1
#     if adto_count == 3 :
#         flag = 1
    return adto_count


# In[77]:
# added 27-09-2024
def restricted_reason_fno_non_fno(row_df , exclude = '') :
    row_ = row_df.reset_index(drop=True)
    row = row_.iloc[0]
    reasons = []
    restricted_count = 0 
    if row['Mcap']<100:
        reasons.append('Mcap')
        restricted_count+=1

    if row['[Total Shareholders Funds (Latest)]'] < 100:
        reasons.append('Shareholders Fund')
        restricted_count+=1
    
    if exclude.lower() == 'adto' :
        restricted_count+=1
    else:
        if ((row['ADTO in Cr']<0.2) | pd.isna(row['ADTO in Cr'])  ) : #| (row['3M_ADTO_COUNT']<3)
            reasons.append('ADTO in Cr')
            restricted_count+=1
    
    if row['Promoter Pledge%']>=50:
        reasons.append('Promoter Pledge%')
        restricted_count+=1
    
    if row['Impact Cost']>=1:
        reasons.append('Impact Cost')
        restricted_count+=1
    
    if row['Rev']<=200 :
        reasons.append('Rev')
        restricted_count+=1
    
    if row['PET Check']!='PET Good' :
        reasons.append('PET Check')
    
    if row['F-Score']!='F-Check Passed' :
        reasons.append('F-Score')
    
    if (('F-Score' in reasons) or ('PET Check' in reasons )):
            restricted_count+=1

    if exclude.lower() == 'symbol':
        restricted_count+=1
    else:
        if (row['BSE Series'] in ['Z','P','M'] ) :
            reasons.append('BSE Symbol')
        if (row['NSE Series'] == 'BZ' ) :
            reasons.append('NSE Symbol')
        if (('NSE Symbol' in reasons) or ('BSE Symbol' in reasons )):
            restricted_count+=1
    # reason_str = ''
    # for reason in reasons[:-1]:
    #     reason_str = reason_str+reason+ ' , '
    # if len(reasons)>0:
    #     reason_str = reason_str+reasons[-1]
    if restricted_count == 8:
        return 'Restricted'
    else:
        return ''
    # return reason_str


 #new added 24-09-2024
risk_rating_fno['Rating v1'] = ''
risk_rating_fno['Restricted'] = ''
risk_rating_fno['New Rules Restricted Reason'] = ''
risk_rating_fno['New Rules Rating (ex ADTO) Restricted'] = ''
risk_rating_fno['New Rules Rating (ex Symbol) Restricted'] = ''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 ) & ((risk_rating_fno['ADTO in Cr'][idx]>20) & ((adto_count(risk_rating_fno_row , 20))==3) )  & (risk_rating_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>5000) & (risk_rating_fno['PET Check'][idx]=='PET Good') & (risk_rating_fno['F-Score'][idx]=='F-Check Passed') & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ):
        risk_rating_fno['Rating v1'][idx] = 'Bluechip'
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ((risk_rating_fno['ADTO in Cr'][idx]>15) & ((adto_count(risk_rating_fno_row , 15))==3) )  & (risk_rating_fno['Promoter Pledge%'][idx]<40)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>1000) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['Rating v1'][idx] = 'Good'
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ((risk_rating_fno['ADTO in Cr'][idx]>2) & ((adto_count(risk_rating_fno_row , 2))==3) )  & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>200) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['Rating v1'][idx] = 'Average'
    else:
        risk_rating_fno['Rating v1'][idx] = 'Poor'
    if risk_rating_fno['Rating v1'][idx] == 'Poor' :
        if ((risk_rating_fno['BSE Series'][idx] in ['Z','P','M'] ) |  (risk_rating_fno['NSE Series'][idx] == 'BZ' ) | (risk_rating_fno['ADTO in Cr'][idx] < 0.2)  ): #| (pd.isna(risk_rating_fno['ADTO in Cr'][idx]))
            risk_rating_fno['Restricted'][idx] = 'Restricted'
            reasons = []
            if (risk_rating_fno['BSE Series'][idx] in ['Z','P','M'] ) :
                reasons.append('BSE Symbol')
            if (risk_rating_fno['NSE Series'][idx] == 'BZ' ) :
                reasons.append('NSE Symbol')
            if (risk_rating_fno['ADTO in Cr'][idx] < 0.2) :
                reasons.append('ADTO in Cr')
            reason_str = ''
            for reason in reasons[:-1]:
                reason_str = reason_str+reason+ ' , '
            if len(reasons)>0:
                reason_str = reason_str+reasons[-1]
            risk_rating_fno['New Rules Restricted Reason'][idx] = reason_str

#new added 20-09-2024
risk_rating_fno['New Rules Rating (ex Symbol)'] = ''
risk_rating_fno['New Rules Rating (ex Symbol) Restricted'] = ''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 ) & ((risk_rating_fno['ADTO in Cr'][idx]>20) & ((adto_count(risk_rating_fno_row , 20))==3) )  & (risk_rating_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>5000) & (risk_rating_fno['PET Check'][idx]=='PET Good') & (risk_rating_fno['F-Score'][idx]=='F-Check Passed') & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ):
        risk_rating_fno['New Rules Rating (ex Symbol)'][idx] = 'Bluechip'
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ((risk_rating_fno['ADTO in Cr'][idx]>15) & ((adto_count(risk_rating_fno_row , 15))==3) )  & (risk_rating_fno['Promoter Pledge%'][idx]<40)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>1000) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['New Rules Rating (ex Symbol)'][idx] = 'Good'
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ((risk_rating_fno['ADTO in Cr'][idx]>2) & ((adto_count(risk_rating_fno_row , 2))==3) )  & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>200) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['New Rules Rating (ex Symbol)'][idx] = 'Average'
    else:
        risk_rating_fno['New Rules Rating (ex Symbol)'][idx] = 'Poor'
        # added 27-09-2024
        risk_rating_fno['New Rules Rating (ex Symbol) Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_fno_row , exclude='symbol')
#new added 20-09-2024
risk_rating_fno['New Rules Rating (ex ADTO)'] = ''
risk_rating_fno['New Rules Rating (ex ADTO) Restricted'] = ''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 )  & (risk_rating_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>5000) & (risk_rating_fno['PET Check'][idx]=='PET Good') & (risk_rating_fno['F-Score'][idx]=='F-Check Passed') & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ):
        risk_rating_fno['New Rules Rating (ex ADTO)'][idx] = 'Bluechip'
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 )  & (risk_rating_fno['Promoter Pledge%'][idx]<40)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>1000) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['New Rules Rating (ex ADTO)'][idx] = 'Good'
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 )  & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>200) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['New Rules Rating (ex ADTO)'][idx] = 'Average'
    else:
        risk_rating_fno['New Rules Rating (ex ADTO)'][idx] = 'Poor'
        # added 27-09-2024
        risk_rating_fno['New Rules Rating (ex ADTO) Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_fno_row , exclude='adto')
#new added 20-09-2024
risk_rating_fno['New Rules Rating (ex Symbol, ex ADTO)'] = ''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 )  & (risk_rating_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>5000) & (risk_rating_fno['PET Check'][idx]=='PET Good') & (risk_rating_fno['F-Score'][idx]=='F-Check Passed') & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ):
        risk_rating_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Bluechip'
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 )  & (risk_rating_fno['Promoter Pledge%'][idx]<40)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>1000) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Good'
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 )  & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>200) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Average'
    else:
        risk_rating_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Poor'
#new added 20-09-2024
risk_rating_fno['New Rules - Quick Rating'] = ''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 )  & (risk_rating_fno['Rev'][idx]>5000)  ):
        risk_rating_fno['New Rules - Quick Rating'][idx] = 'Bluechip'
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 )  & (risk_rating_fno['Rev'][idx]>1000)  ): 
        risk_rating_fno['New Rules - Quick Rating'][idx] = 'Good'
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 )  & (risk_rating_fno['Rev'][idx]>200)   ): 
        risk_rating_fno['New Rules - Quick Rating'][idx] = 'Average'
    else:
        risk_rating_fno['New Rules - Quick Rating'][idx] = 'Poor'

# In[78]:


risk_rating_fno['3M_ADTO_COUNT'] = np.nan
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 ) & ((risk_rating_fno['ADTO in Cr'][idx]>20)  )  & (risk_rating_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>5000) & (risk_rating_fno['PET Check'][idx]=='PET Good') & (risk_rating_fno['F-Score'][idx]=='F-Check Passed') & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ):
        risk_rating_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_fno_row , 20)
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ((risk_rating_fno['ADTO in Cr'][idx]>15)  )  & (risk_rating_fno['Promoter Pledge%'][idx]<40)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>1000) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_fno_row , 15)
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ((risk_rating_fno['ADTO in Cr'][idx]>2)  )  & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>200) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
        risk_rating_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_fno_row , 2)
    else:
        risk_rating_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_fno_row , 2)


# In[79]:

 #new added 24-09-2024
risk_rating_fno['Poor_Reasons'] = ''
risk_rating_fno['Poor_Reasons_Count'] = 0
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    reasons = []
    poor_count = 0
    if risk_rating_fno['Rating v1'][idx] == 'Poor' :
        
        if risk_rating_fno['Mcap'][idx]<100:
            reasons.append('Mcap')
            poor_count+=1
        
        if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] < 100:
            reasons.append('Shareholders Fund')
            poor_count+=1
        
        if ((risk_rating_fno['ADTO in Cr'][idx]<2) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) ) :
            reasons.append('ADTO in Cr')
            poor_count+=1
        
        if risk_rating_fno['Promoter Pledge%'][idx]>=50:
            reasons.append('Promoter Pledge%')
            poor_count+=1
        
        if risk_rating_fno['Impact Cost'][idx]>=1:
            reasons.append('Impact Cost')
            poor_count+=1
        
        if risk_rating_fno['Rev'][idx]<=200 :
            reasons.append('Rev')
            poor_count+=1
        
        if risk_rating_fno['PET Check'][idx]!='PET Good' :
            reasons.append('PET Check')
            poor_count+=1
        
        if risk_rating_fno['F-Score'][idx]!='F-Check Passed' :
            reasons.append('F-Score')
            poor_count+=1
                           
        if risk_rating_fno['Nifty 500'][idx]!='Nifty 500':
            reasons.append('Nifty 500')
            poor_count+=1
        
            
        reason_str = ''
        for reason in reasons[:-1]:
            reason_str = reason_str+reason+ ' , '
        if len(reasons)>0:
            reason_str = reason_str+reasons[-1]
        risk_rating_fno['Poor_Reasons'][idx] = reason_str
        risk_rating_fno['Poor_Reasons_Count'][idx] = poor_count


# In[ ]:





# In[80]:


risk_rating_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'] = ''
risk_rating_fno['Meeting Size Criteria_Poor'] = ''
risk_rating_fno['Size Exception for Poor to Average'] = ''

risk_rating_fno['All Conditions of Good Met (Except "Good" Size)'] = ''
risk_rating_fno['Size Criteria_Average'] = ''
risk_rating_fno['Size Exception for Average to Good'] = ''

# risk_rating_fno['All conditions met Except ADTO'] = ''
# risk_rating_fno['Size Criteria for ADTO Upgrade'] = ''
# risk_rating_fno['ADTO > 0.5 Cr'] = ''
# risk_rating_fno['ADTO Upgrade (Poor to Average)'] = ''

for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    ####change below adto
    if risk_rating_fno['Rating v1'][idx] == 'Poor' :
        risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
        if ( (risk_rating_fno['Mcap'][idx]>=50 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=50 ) & (risk_rating_fno['ADTO in Cr'][idx]>2) & ((adto_count(risk_rating_fno_row , 2))==3)  & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>100) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ):
            
            risk_rating_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'][idx] = 'Yes'
        else:
            risk_rating_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'][idx] = 'No'
        
        count_p = 0
        
        if (risk_rating_fno['Mcap'][idx]>=150 ):
            count_p+=1
        if (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=150 ):
            count_p+=1
        if (risk_rating_fno['Rev'][idx]>300):
            count_p+=1
        risk_rating_fno['Meeting Size Criteria_Poor'][idx] = count_p
    
        if ((risk_rating_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'][idx]=='Yes') & (risk_rating_fno['Meeting Size Criteria_Poor'][idx]==2) ):
            risk_rating_fno['Size Exception for Poor to Average'][idx] = 'Upgrade'
    ####change below adto
    if risk_rating_fno['Rating v1'][idx] == 'Average' :
        risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
        if ( (risk_rating_fno['ADTO in Cr'][idx]>15) & ((adto_count(risk_rating_fno_row , 15))==3) & (risk_rating_fno['Promoter Pledge%'][idx] < 40)) :
            risk_rating_fno['All Conditions of Good Met (Except "Good" Size)'][idx] = 'Yes'
        else:
            risk_rating_fno['All Conditions of Good Met (Except "Good" Size)'][idx] = 'No'
        
        count_a = 0
        if (risk_rating_fno['Mcap'][idx]>=750 ):
            count_a+=1
        if (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=750 ):
            count_a+=1
        if (risk_rating_fno['Rev'][idx]>1500):
            count_a+=1
        risk_rating_fno['Size Criteria_Average'][idx] = count_a
        
        if ((risk_rating_fno['All Conditions of Good Met (Except "Good" Size)'][idx]=='Yes') & (risk_rating_fno['Size Criteria_Average'][idx]>=2) ):
            risk_rating_fno['Size Exception for Average to Good'][idx] = 'Upgrade'
    
    
#     if risk_rating_fno['Rating v1'][idx] == 'Poor' :
#         if ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & (risk_rating_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_fno['Impact Cost'][idx]<1) & (risk_rating_fno['Rev'][idx]>200) & ((risk_rating_fno['PET Check'][idx]=='PET Good') | (risk_rating_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_fno['Nifty 500'][idx]=='Nifty 500')   ): 
#             risk_rating_fno['All conditions met Except ADTO'][idx] = 'Yes'
#         else:
#             risk_rating_fno['All conditions met Except ADTO'][idx] = 'No'
        
        
#         count_p_adto = 0
        
#         if (risk_rating_fno['Mcap'][idx]>=300 ):
#             count_p_adto+=1
#         if (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=300 ):
#             count_p_adto+=1
#         if (risk_rating_fno['Rev'][idx]>600):
#             count_p_adto+=1
#         risk_rating_fno['Size Criteria for ADTO Upgrade'][idx] = count_p_adto
        
        
#         if (risk_rating_fno['ADTO in Cr'][idx]>0.5) :
#             risk_rating_fno['ADTO > 0.5 Cr'][idx] = 'Yes'
#         else:
#             risk_rating_fno['ADTO > 0.5 Cr'][idx] = 'No'
        
#         if ((risk_rating_fno['ADTO > 0.5 Cr'][idx]=='Yes') & (risk_rating_fno['Size Criteria for ADTO Upgrade'][idx]==3) & (risk_rating_fno['All conditions met Except ADTO'][idx] == 'Yes')):
#             risk_rating_fno['ADTO Upgrade (Poor to Average)'][idx] = 'Upgrade'


# In[81]:


risk_rating_fno['ADTO_Wipsaw_Upgrade'] = ''
risk_rating_fno['Rating v1_5'] = ''
risk_rating_fno['Rating v2'] = ''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    if risk_rating_fno['Rating v1'][idx] in ['Bluechip','Good'] :
        risk_rating_fno['Rating v2'][idx] = risk_rating_fno['Rating v1'][idx]
    elif risk_rating_fno['Rating v1'][idx]=='Poor' :
        if ((risk_rating_fno['Size Exception for Poor to Average'][idx]=='Upgrade')): # | (risk_rating_fno['ADTO Upgrade (Poor to Average)'][idx] == 'Upgrade')
            risk_rating_fno['Rating v2'][idx] = 'Average'
        else:
            risk_rating_fno['Rating v2'][idx] = risk_rating_fno['Rating v1'][idx]
    elif risk_rating_fno['Rating v1'][idx]=='Average' :
        if risk_rating_fno['Size Exception for Average to Good'][idx]=='Upgrade':
            risk_rating_fno['Rating v2'][idx] = 'Good'
        else:
            risk_rating_fno['Rating v2'][idx] = risk_rating_fno['Rating v1'][idx]


# In[82]:


risk_rating_fno['Rating v1_5'] = risk_rating_fno['Rating v2']


# In[ ]:





# In[83]:


#fno wipsaw


# In[84]:


def fno_average_reasons_wipsaw(risk_rating_fno):
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=100:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 100:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>2) ) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<50:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>200 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

    if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
        reasons.append('F-Score')

    if risk_rating_fno['Nifty 500'][idx]=='Nifty 500':
        reasons.append('Nifty 500')
    
    return reasons


# In[85]:


for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]]
    if risk_rating_fno['Rating v1'][idx] == 'Poor' :
        prev_month_row = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
        if len(prev_month_row)>0 :
            if prev_month_row['As per new Rules'][0] == 'Average':
                if risk_rating_fno['Poor_Reasons'][idx] == 'ADTO in Cr' :
                    adto_in_cr = 1
                    adto_thresh_count = adto_count(risk_rating_fno_row , adto_in_cr)
                    risk_rating_fno_row = risk_rating_fno_row.reset_index(drop=True)
                    adto_median_ = risk_rating_fno_row[latest_3_months_adto_cols]
                    adto_values = []
                    for col in adto_median_.columns:
                        adto_values.append(adto_median_[col][0])
                    adto_median = median(adto_values)
                    
                    
                    if ( (adto_thresh_count ==2 ) & (adto_median > adto_in_cr) ):
#                         print(risk_rating_fno['CO_NAME'][idx])
                        risk_rating_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                        risk_rating_fno['Rating v2'][idx] = 'Average'   
    
    if risk_rating_fno['Rating v1'][idx] == 'Poor' :
        prev_month_row = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
        if len(prev_month_row)>0 :
            if prev_month_row['As per new Rules'][0] == 'Good':
                if risk_rating_fno['Poor_Reasons'][idx] == 'ADTO in Cr' :
                    adto_in_cr = 2
                    adto_thresh_count = adto_count(risk_rating_fno_row , adto_in_cr)
                    risk_rating_fno_row = risk_rating_fno_row.reset_index(drop=True)
                    adto_median_ = risk_rating_fno_row[latest_3_months_adto_cols]
                    adto_values = []
                    for col in adto_median_.columns:
                        adto_values.append(adto_median_[col][0])
                    adto_median = median(adto_values)
                    
                    if ( (adto_thresh_count ==2 ) & (adto_median > adto_in_cr) ):
#                         print(risk_rating_fno['CO_NAME'][idx])
                        risk_rating_fno['Rating v2'][idx] = 'Good'
                        risk_rating_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                    else:
                        adto_in_cr = 1
                        adto_thresh_count = adto_count(risk_rating_fno_row , adto_in_cr)
                        risk_rating_fno_row = risk_rating_fno_row.reset_index(drop=True)
                        adto_median_ = risk_rating_fno_row[latest_3_months_adto_cols]
                        adto_values = []
                        for col in adto_median_.columns:
                            adto_values.append(adto_median_[col][0])
                        adto_median = median(adto_values)
                        if ( (adto_thresh_count ==2 ) & (adto_median > adto_in_cr) ):
#                             print(risk_rating_fno['CO_NAME'][idx])
                            risk_rating_fno['Rating v2'][idx] = 'Average'
                            risk_rating_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                        
            
    if risk_rating_fno['Rating v1'][idx] == 'Average' :
        prev_month_row = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
        if len(prev_month_row)>0 :
            if prev_month_row['As per new Rules'][0] == 'Good':
                risk_rating_fno_row_ = risk_rating_fno_row.reset_index(drop=True)
                if fno_average_reasons_wipsaw(risk_rating_fno_row_) == ['ADTO in Cr'] :
                    adto_in_cr = 2
                    adto_thresh_count = adto_count(risk_rating_fno_row , adto_in_cr)
                    risk_rating_fno_row = risk_rating_fno_row.reset_index(drop=True)
                    adto_median_ = risk_rating_fno_row[latest_3_months_adto_cols]
                    adto_values = []
                    for col in adto_median_.columns:
                        adto_values.append(adto_median_[col][0])
                    adto_median = median(adto_values)
                    
                    if ( (adto_thresh_count == 2 ) & (adto_median > adto_in_cr) ):
#                         print(risk_rating_fno['CO_NAME'][idx])
                        risk_rating_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                        risk_rating_fno['Rating v2'][idx] = 'Good'   


# In[ ]:





# In[86]:


# risk_rating_fno['Rating v2'] = risk_rating_fno['Rating v1']


# In[ ]:





# In[87]:


risk_rating_fno['Poor_Alz_Manual'] = ''
risk_rating_fno['Poor_Alz_Manual_Comments'] = ''
risk_rating_fno['Poor_to_Avg_Manual'] = ''
risk_rating_fno['Good_BC_to_Avg_Manual'] = ''
risk_rating_fno['Manual_Rating'] = risk_rating_fno['Rating v2']
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    temp_1 = poor_downgrd_rated_df[poor_downgrd_rated_df['CAPITALINE CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if len(temp_1) > 0 :
        risk_rating_fno['Manual_Rating'][idx] = temp_1['Rating'][0]
        risk_rating_fno['Poor_Alz_Manual'][idx] = 'Downgraded_to_Poor'
        risk_rating_fno['Poor_Alz_Manual_Comments'][idx] = temp_1['Comments'][0]
    
    if risk_rating_fno['Rating v2'][idx] == 'Poor' :
        temp_2 = poor_to_avg_rated_df[poor_to_avg_rated_df['CAPITALINE CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
        if len(temp_2) > 0 :
            risk_rating_fno['Manual_Rating'][idx] = temp_2['Rating'][0]
            risk_rating_fno['Poor_to_Avg_Manual'][idx] = 'Upgraded_to_Average'
    
    if risk_rating_fno['Rating v2'][idx] in ['Bluechip','Good'] :
        temp_3 = good_bc_to_avg_rated_df[good_bc_to_avg_rated_df['CAPITALINE CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
        if len(temp_3) > 0 :
            risk_rating_fno['Manual_Rating'][idx] = temp_3['Rating'][0]
            risk_rating_fno['Good_BC_to_Avg_Manual'][idx] = 'Downgraded_to_Average'
    


# In[88]:


def pet_check_old_rating(risk_rating_fno_row,pet_value):
    max_score = 0 
    
    co_code = risk_rating_fno_row['CO_CODE'][0]
    pet_data = pet_check_code[pet_check_code['CO_CODE'] == co_code].reset_index()
    
    if len(pet_data) > 0 :
        
        pow_y1 = pet_data['Power & Fuel Cost Y1'][0]
        pow_y2 = pet_data['Power & Fuel Cost Y2'][0]
        emp_y1 = pet_data['Employee Cost Y1'][0]
        emp_y2 = pet_data['Employee Cost Y2'][0]
        tax_y1 = pet_data['Total_Tax1'][0]
        tax_y2 = pet_data['Total_Tax2'][0]
        
        pow_y1_check = 0
        pow_y2_check = 0
        emp_y1_check = 0
        emp_y2_check = 0
        tax_y1_check = 0
        tax_y2_check = 0
        
        if pow_y1 >= pet_value :
            pow_y1_check = 1
        
        if pow_y2 >= pet_value :
            pow_y2_check = 1
        
        if emp_y1 >= pet_value :
            emp_y1_check = 1
        
        if emp_y2 >= pet_value :
            emp_y2_check = 1
        
        if tax_y1 >= pet_value :
            tax_y1_check = 1
        
        if tax_y2 >= pet_value :
            tax_y2_check = 1
        
        
        pow_emp_sum = pow_y1_check + pow_y2_check + emp_y1_check + emp_y2_check
        pow_tax_sum = pow_y1_check + pow_y2_check + tax_y1_check + tax_y2_check
        emp_tax_sum = emp_y1_check + emp_y2_check + tax_y1_check + tax_y2_check
        
        max_score = max([pow_emp_sum,pow_tax_sum,emp_tax_sum])
    
    
    else:
        max_score = 0
    
    return max_score


# In[89]:


def pet_check_old_rating_resticted(risk_rating_fno_row,pet_value):
    max_score = 0 
    
    co_code = risk_rating_fno_row['CO_CODE'][0]
    pet_data = pet_check_code[pet_check_code['CO_CODE'] == co_code].reset_index()
    
    if len(pet_data) > 0 :
        
        pow_y1 = pet_data['Power & Fuel Cost Y1'][0]
        pow_y2 = pet_data['Power & Fuel Cost Y2'][0]
        emp_y1 = pet_data['Employee Cost Y1'][0]
        emp_y2 = pet_data['Employee Cost Y2'][0]
        tax_y1 = pet_data['Total_Tax1'][0]
        tax_y2 = pet_data['Total_Tax2'][0]
        
        pow_y1_check = 0
        pow_y2_check = 0
        emp_y1_check = 0
        emp_y2_check = 0
        tax_y1_check = 0
        tax_y2_check = 0
        
        if pow_y1 <= pet_value :
            pow_y1_check = 1
        
        if pow_y2 <= pet_value :
            pow_y2_check = 1
        
        if emp_y1 <= pet_value :
            emp_y1_check = 1
        
        if emp_y2 <= pet_value :
            emp_y2_check = 1
        
        if tax_y1 <= pet_value :
            tax_y1_check = 1
        
        if tax_y2 <= pet_value :
            tax_y2_check = 1
        
        
        pow_emp_sum = pow_y1_check + pow_y2_check + emp_y1_check + emp_y2_check
        pow_tax_sum = pow_y1_check + pow_y2_check + tax_y1_check + tax_y2_check
        emp_tax_sum = emp_y1_check + emp_y2_check + tax_y1_check + tax_y2_check
        
        max_score = max([pow_emp_sum,pow_tax_sum,emp_tax_sum])
    
    
    else:
        max_score = 0
    
    return max_score


# In[90]:


risk_rating_fno['As per new Rules'] = risk_rating_fno['Manual_Rating']
risk_rating_fno['As per new Rules - Restricted'] = ''


# In[91]:


risk_rating_fno['As per old Rules'] =''
risk_rating_fno['As per old Rules - Restricted'] =''
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    
    risk_rating_fno_row = risk_rating_fno[risk_rating_fno['CO_CODE'] == risk_rating_fno['CO_CODE'][idx]].reset_index()
    if ( (risk_rating_fno['Mcap'][idx]>=5000 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=2000 ) & ( pet_check_old_rating(risk_rating_fno_row,50)==4 )   ):
        risk_rating_fno['As per old Rules'][idx] = 'Bluechip'
    elif ( (risk_rating_fno['Mcap'][idx]>=500 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ( pet_check_old_rating(risk_rating_fno_row,10)==4 )   ): 
        risk_rating_fno['As per old Rules'][idx] = 'Good'
    elif ( (risk_rating_fno['Mcap'][idx]>=100 ) & (risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ( pet_check_old_rating(risk_rating_fno_row,2)==4 )   ): 
        risk_rating_fno['As per old Rules'][idx] = 'Average'
    else:
        risk_rating_fno['As per old Rules'][idx] = 'Poor'
        if (pet_check_old_rating_resticted(risk_rating_fno_row,0.5) ==4) :
            risk_rating_fno['As per new Rules - Restricted'][idx] = 'Restricted'
        
    # added 27-09-2024 # commented 25-11-2024
    # if risk_rating_fno['As per new Rules'][idx] == 'Poor' : 
    #         risk_rating_fno['As per new Rules - Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_fno_row , exclude='')


# In[92]:


risk_rating_fno['Current Live Rating'] = ''
risk_rating_fno['Current Live Restricted'] = ''
 #new added 24-09-2024
risk_rating_fno[['BSE_VAR_pct','NSE_VAR_pct','ANGEL_VAR_pct','MTF_VAR_pct']] = np.nan
for idx in range(0,len(risk_rating_fno['CO_CODE'])):
    co_code = risk_rating_fno['CO_CODE'][idx]
    isin_t = bse_nse_group_sheet[bse_nse_group_sheet['CAPITALINE CODE'] == co_code].reset_index(drop=True)
    if len(isin_t) > 0:
        isin = isin_t['[ISIN No'][0]
        temp_curr_rating = angel_scrip_category_df[angel_scrip_category_df['ISIN No'] == isin].reset_index(drop=True)
        if len(temp_curr_rating) > 0 :
            risk_rating_fno['Current Live Rating'][idx] = temp_curr_rating['Angel scrip category'][0].strip()
            risk_rating_fno['Current Live Restricted'][idx] = temp_curr_rating['Restricted Scrips'][0].strip()
            #new added 24-09-2024
            risk_rating_fno['BSE_VAR_pct'][idx] = temp_curr_rating['BSE_VAR %'][0]#.strip()
            risk_rating_fno['NSE_VAR_pct'][idx] = temp_curr_rating['NSE_VAR %'][0]#.strip()
            risk_rating_fno['ANGEL_VAR_pct'][idx] = temp_curr_rating['ANGEL_VAR %'][0]#.strip()
            risk_rating_fno['MTF_VAR_pct'][idx] = temp_curr_rating['MTF_VAR %'][0]#.strip()


# In[ ]:





# In[93]:


on_premise_db = create_engine(
    "mysql+pymysql://{user}:{pw}@{host}:{port}/{db}".format(user="dharanee.patel",
                                                        pw="test$1234",
                                                        db="fundamental",
                                                        host="10.253.13.130",
                                                        port = '33066'))


# In[94]:


long_short_names = pd.read_sql_table('long_short_master', on_premise_db)


# In[95]:


close_price_df = pd.DataFrame()
close_price_df[['CO_CODE','CO_NAME']] = pet_check_code[['CO_CODE', 'CO_NAME']]
close_price_df = close_price_df[~close_price_df['CO_CODE'].isin(risk_rating_fno['CO_CODE'])].reset_index(drop = True)

close_price_df['NSE'] = np.datetime64('NAT')
close_price_df['BSE'] = np.datetime64('NAT')
close_price_df['Latest'] = np.datetime64('NAT')
for idx in range(0,len(close_price_df['CO_CODE'])):
    temp_bse_cp = close_price_sheet_bse[close_price_sheet_bse['CAPITALINE CODE'] == close_price_df['CO_CODE'][idx] ].reset_index()
    temp_nse_cp = close_price_sheet_nse[close_price_sheet_nse['CAPITALINE CODE'] == close_price_df['CO_CODE'][idx] ].reset_index()
    if len(temp_bse_cp) > 0 :
        close_price_df['BSE'][idx] = temp_bse_cp['[Date (Latest)]'][0]
    else:
        close_price_df['BSE'][idx] = def_date
    
    if len(temp_nse_cp) > 0 :
        close_price_df['NSE'][idx] = temp_nse_cp['[Date (Latest)]'][0]
    else:
        close_price_df['NSE'][idx] = def_date
    
    close_price_df['Latest'][idx] = max( [close_price_df['BSE'][idx] ,close_price_df['NSE'][idx]] )


# In[96]:


bse_series_symbols = ["T","X","XT","Z","P","M","MT","ZP","MS"]
nse_series_symbols = ["BE","BZ","SM","ST"]


# In[ ]:





# In[97]:


risk_rating_non_fno = pd.DataFrame()
# risk_rating_non_fno[['CO_CODE','CO_NAME']] = pet_check_code[['CO_CODE', 'CO_NAME']]
risk_rating_non_fno[['CO_CODE','CO_NAME']] = bse_nse_group_sheet[['CAPITALINE CODE', 'CO_NAME']]
risk_rating_non_fno = risk_rating_non_fno[~risk_rating_non_fno['CO_CODE'].isin(risk_rating_fno['CO_CODE'])].reset_index(drop = True)
risk_rating_non_fno['[Total Shareholders Funds (Latest)]'] = np.nan
risk_rating_non_fno['Mcap'] = np.nan
risk_rating_non_fno['BSE Series'] = ''
risk_rating_non_fno['NSE Series'] = ''
risk_rating_non_fno['Impact Cost'] = np.nan
risk_rating_non_fno['PET Check'] = ''
risk_rating_non_fno['F-Score'] = ''
risk_rating_non_fno['[ISIN No'] = ''
risk_rating_non_fno['BSE symbol downgrade?'] = ''
risk_rating_non_fno['NSE symbol downgrade?'] = ''
risk_rating_non_fno['Market Cap/ISIN'] = ''
risk_rating_non_fno['Promoter Pledge%'] = np.nan
risk_rating_non_fno['ADTO in Cr'] = np.nan
risk_rating_non_fno['Rev'] = np.nan
risk_rating_non_fno['Last Trading Date?'] = np.datetime64('NAT')
risk_rating_non_fno['Not traded?'] = 'Yes'

#new added 20-09-2024
risk_rating_non_fno['Percent_Funding'] = np.nan

# In[98]:


risk_rating_non_fno[latest_3_months_adto_cols] = 0
# risk_rating_non_fno['3M_ADTO_count'] = np.nan
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    temp_adto_df = adto_df[adto_df['CAPITALINE CODE'] ==risk_rating_non_fno['CO_CODE'][idx] ].reset_index(drop=True)
    if len(temp_adto_df) > 0 :
        for col in latest_3_months_adto:
            risk_rating_non_fno[f'{col}_ADTO'][idx] = temp_adto_df[col][0]/pow(10,7)


# In[99]:


for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    
    temp_nw = Networth_sheet[risk_rating_non_fno['CO_CODE'][idx] == Networth_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_nw)>0 :
        risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx] = temp_nw['[Networth (Latest)]'][0]
    
    temp_mcap = latest_mcap_sheet[risk_rating_non_fno['CO_CODE'][idx] == latest_mcap_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_mcap) > 0 :
        risk_rating_non_fno['Mcap'][idx] = temp_mcap['[Market Cap (Latest)]'][0]
    
    temp_bsensegrp = bse_nse_group_sheet[risk_rating_non_fno['CO_CODE'][idx] == bse_nse_group_sheet['CAPITALINE CODE']].reset_index()
    if len(temp_bsensegrp)>0:
        risk_rating_non_fno['BSE Series'][idx] = temp_bsensegrp['[BSE Group'][0]
        risk_rating_non_fno['NSE Series'][idx] = temp_bsensegrp['[NSE Series'][0]
    
    temp_impact_cost = impact_cost[impact_cost['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_impact_cost)>0:
        risk_rating_non_fno['Impact Cost'][idx] = temp_impact_cost['Min of NSE/BSE'][0]
    
    
    temp_pet_chk = pet_check_code[pet_check_code['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_pet_chk)>0:
        risk_rating_non_fno['PET Check'][idx] = temp_pet_chk['Final'][0]
    
    temp_fscore = f_score_code[f_score_code['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_fscore)>0:
        risk_rating_non_fno['F-Score'][idx] = temp_fscore['Final Check'][0]
    
    temp_lsnames_ = risk_rating_nonfno_isin[risk_rating_nonfno_isin['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_lsnames_)>0:
        risk_rating_non_fno['[ISIN No'][idx] = temp_lsnames_['[ISIN No'][0]
    
    if risk_rating_non_fno['[ISIN No'][idx] == '' :
        temp_lsnames = long_short_names[long_short_names['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
        if len(temp_lsnames)>0:
            risk_rating_non_fno['[ISIN No'][idx] = temp_lsnames['[ISIN No'][0]

    
    
    if risk_rating_non_fno['BSE Series'][idx] in bse_series_symbols:
        risk_rating_non_fno['BSE symbol downgrade?'][idx] = "Yes"
    else:
        risk_rating_non_fno['BSE symbol downgrade?'][idx] = "No"
    
    if risk_rating_non_fno['NSE Series'][idx] in nse_series_symbols:
        risk_rating_non_fno['NSE symbol downgrade?'][idx] = "Yes"
    else:
        risk_rating_non_fno['NSE symbol downgrade?'][idx] = "No"
    
    if ( (pd.isna(risk_rating_non_fno['Mcap'][idx])) | (pd.isna(risk_rating_non_fno['[ISIN No'][idx])) | (risk_rating_non_fno['Mcap'][idx] == 0) | (risk_rating_non_fno['[ISIN No'][idx] == '') ):
        risk_rating_non_fno['Market Cap/ISIN'][idx] = 'Yes'
    else:
        risk_rating_non_fno['Market Cap/ISIN'][idx] = 'No'
    
    temp_promoter_pledge = promoter_pledge_sheet[promoter_pledge_sheet['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_promoter_pledge)>0:
        risk_rating_non_fno['Promoter Pledge%'][idx] = temp_promoter_pledge['[Total of Promoter and Group (Latest)]'][0]
    
    temp_adto = adto_sheet[adto_sheet['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_adto)>0:
        risk_rating_non_fno['ADTO in Cr'][idx] = temp_adto['Amount Traded Rs.'][0]/pow(10,7)
    
    temp_sales_pat = sales_pat_sheet[sales_pat_sheet['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_sales_pat)>0:
        risk_rating_non_fno['Rev'][idx] = temp_sales_pat['[Net Sales (Latest)]'][0]
    
    temp_ltd = close_price_df[close_price_df['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_ltd)>0:
        risk_rating_non_fno['Last Trading Date?'][idx] = temp_ltd['Latest'][0]
        if risk_rating_non_fno['Last Trading Date?'][idx] < last_allowed_trading_date :
            risk_rating_non_fno['Not traded?'][idx] = 'Yes'
        else:
            risk_rating_non_fno['Not traded?'][idx] = 'No'
    #new added 20-09-2024
    temp_funding = funding_df[funding_df['ISIN'] == risk_rating_non_fno['[ISIN No'][idx]].reset_index()
    if len(temp_funding) > 0 :
        risk_rating_non_fno['Percent_Funding'][idx] = temp_funding['% Funding'][0]

# In[ ]:





# In[ ]:





# In[100]:


def adto_count(risk_rating_non_fno , adto_in_cr):
#     flag = 0
    adto_count = 0
    risk_rating_non_fno_row = risk_rating_non_fno.reset_index(drop=True)
#     print(risk_rating_non_fno_row)
    for col in latest_3_months_adto_cols:
        if risk_rating_non_fno_row[col][0] > adto_in_cr:
            adto_count+=1
#     if adto_count == 3 :
#         flag = 1
    return adto_count


# In[101]:


def adto_count_poor(risk_rating_non_fno , adto_in_cr):
#     flag = 0
    adto_count = 0
    risk_rating_non_fno_row = risk_rating_non_fno.reset_index(drop=True)
    for col in latest_3_months_adto_cols:
        if risk_rating_non_fno_row[col][0] < adto_in_cr:
            adto_count+=1
#     if adto_count == 3 :
#         flag = 1
    return adto_count


# In[102]:

#new added 20-09-2024 updated as well
# risk_rating_non_fno['3M_ADTO_Count'] = np.nan
risk_rating_non_fno['Rating v1'] = ''
risk_rating_non_fno['Restricted'] = ''
risk_rating_non_fno['New Rules Restricted Reason'] = ''
risk_rating_non_fno['New Rules Rating (ex ADTO) Restricted'] = ''
risk_rating_non_fno['New Rules Rating (ex Symbol) Restricted'] = ''
risk_rating_non_fno['As per new Rules - Restricted'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ( (risk_rating_non_fno['ADTO in Cr'][idx]>2) & ((adto_count(risk_rating_non_fno_row , 2))==3) )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>1000) & (risk_rating_non_fno['PET Check'][idx]=='PET Good') & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')): 
        risk_rating_non_fno['Rating v1'][idx] = 'Good'
    
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ( (risk_rating_non_fno['ADTO in Cr'][idx]>1) & ((adto_count(risk_rating_non_fno_row , 1))==3) )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>200) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
            risk_rating_non_fno['Rating v1'][idx] = 'Average'
    else:
        risk_rating_non_fno['Rating v1'][idx] = 'Poor'
    
    if risk_rating_non_fno['Rating v1'][idx] == 'Poor' :
        # added 27-09-2024
        # risk_rating_non_fno['As per new Rules - Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_non_fno_row , exclude='')
        # updated bse series , z,p,m to z,p # removed m 
        if ((risk_rating_non_fno['BSE Series'][idx] in ['Z','P'] ) |  (risk_rating_non_fno['NSE Series'][idx] == 'BZ' ) | (risk_rating_non_fno['ADTO in Cr'][idx] < 0.2)  ): #| (pd.isna(risk_rating_non_fno['ADTO in Cr'][idx]))
            risk_rating_non_fno['Restricted'][idx] = 'Restricted'
            reasons = []
            if (risk_rating_non_fno['BSE Series'][idx] in ['Z','P'] ) :
                reasons.append('BSE Symbol')
            if (risk_rating_non_fno['NSE Series'][idx] == 'BZ' ) :
                reasons.append('NSE Symbol')
            if (risk_rating_non_fno['ADTO in Cr'][idx] < 0.2) :
                reasons.append('ADTO in Cr')
            reason_str = ''
            for reason in reasons[:-1]:
                reason_str = reason_str+reason+ ' , '
            if len(reasons)>0:
                reason_str = reason_str+reasons[-1]
            risk_rating_non_fno['New Rules Restricted Reason'][idx] = reason_str
            

#new added 20-09-2024
risk_rating_non_fno['New Rules Rating (ex Symbol)'] = ''
risk_rating_non_fno['New Rules Rating (ex Symbol) Restricted'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ( (risk_rating_non_fno['ADTO in Cr'][idx]>2) & ((adto_count(risk_rating_non_fno_row , 2))==3) )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>1000) & (risk_rating_non_fno['PET Check'][idx]=='PET Good') & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')): 
        risk_rating_non_fno['New Rules Rating (ex Symbol)'][idx] = 'Good'
    
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ( (risk_rating_non_fno['ADTO in Cr'][idx]>1) & ((adto_count(risk_rating_non_fno_row , 1))==3) )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>200) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
            risk_rating_non_fno['New Rules Rating (ex Symbol)'][idx] = 'Average'
    else:
        risk_rating_non_fno['New Rules Rating (ex Symbol)'][idx] = 'Poor'
        # added 27-09-2024
        risk_rating_non_fno['New Rules Rating (ex Symbol) Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_non_fno_row , exclude='symbol')
#new added 20-09-2024
risk_rating_non_fno['New Rules Rating (ex ADTO)'] = ''
risk_rating_non_fno['New Rules Rating (ex ADTO) Restricted'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>1000) & (risk_rating_non_fno['PET Check'][idx]=='PET Good') & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')): 
        risk_rating_non_fno['New Rules Rating (ex ADTO)'][idx] = 'Good'
    
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>200) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
            risk_rating_non_fno['New Rules Rating (ex ADTO)'][idx] = 'Average'
    else:
        risk_rating_non_fno['New Rules Rating (ex ADTO)'][idx] = 'Poor'
        # added 27-09-2024
        risk_rating_non_fno['New Rules Rating (ex ADTO) Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_non_fno_row , exclude='adto')
#new added 20-09-2024
risk_rating_non_fno['New Rules Rating (ex Symbol, ex ADTO)'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>1000) & (risk_rating_non_fno['PET Check'][idx]=='PET Good') &  (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')): 
        risk_rating_non_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Good'
    
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>200) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) &  (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
            risk_rating_non_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Average'
    else:
        risk_rating_non_fno['New Rules Rating (ex Symbol, ex ADTO)'][idx] = 'Poor'
#new added 20-09-2024
risk_rating_non_fno['New Rules - Quick Rating'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & (  (risk_rating_non_fno['Rev'][idx]>1000) ) ): 
        risk_rating_non_fno['New Rules - Quick Rating'][idx] = 'Good'
    
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & (  (risk_rating_non_fno['Rev'][idx]>200) ) ): 
            risk_rating_non_fno['New Rules - Quick Rating'][idx] = 'Average'
    else:
        risk_rating_non_fno['New Rules - Quick Rating'][idx] = 'Poor'

# In[103]:


risk_rating_non_fno['3M_ADTO_COUNT'] = np.nan
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ( (risk_rating_non_fno['ADTO in Cr'][idx]>2) )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<25)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>1000) & (risk_rating_non_fno['PET Check'][idx]=='PET Good') & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')): 
        risk_rating_non_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_non_fno_row , 2)
    
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ( (risk_rating_non_fno['ADTO in Cr'][idx]>1)  )  & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>200) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
            risk_rating_non_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_non_fno_row , 1)
    else:
        risk_rating_non_fno['3M_ADTO_COUNT'][idx] = adto_count(risk_rating_non_fno_row , 1)


# In[104]:


# risk_rating_non_fno[risk_rating_non_fno['3M_ADTO_COUNT']!=3]


# In[ ]:





# In[ ]:





# In[105]:


risk_rating_non_fno['Poor_Reasons'] = ''
risk_rating_non_fno['Poor_Reasons_Count'] = 0
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    reasons = []
    poor_count = 0
    if risk_rating_non_fno['Rating v1'][idx] == 'Poor' :
        
        if risk_rating_non_fno['Mcap'][idx]<100:
            reasons.append('Mcap')
            poor_count+=1

        if risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx] < 100:
            reasons.append('Shareholders Fund')
            poor_count+=1
        
        if ((risk_rating_non_fno['ADTO in Cr'][idx]<1) | pd.isna(risk_rating_non_fno['ADTO in Cr'][idx]) | (risk_rating_non_fno['3M_ADTO_COUNT'][idx]<3)) :
            reasons.append('ADTO in Cr')
            poor_count+=1
        
        if risk_rating_non_fno['Promoter Pledge%'][idx]>=50:
            reasons.append('Promoter Pledge%')
            poor_count+=1
        
        if risk_rating_non_fno['Impact Cost'][idx]>=1:
            reasons.append('Impact Cost')
            poor_count+=1
        
        if risk_rating_non_fno['Rev'][idx]<=200 :
            reasons.append('Rev')
            poor_count+=1
        
        if risk_rating_non_fno['PET Check'][idx]!='PET Good' :
            reasons.append('PET Check')
            poor_count+=1
        
        if risk_rating_non_fno['F-Score'][idx]!='F-Check Passed' :
            reasons.append('F-Score')
            poor_count+=1
                           
        if (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "Yes"):
            reasons.append('NSE symbol downgrade')
            poor_count+=1
        
        if (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "Yes"):
            reasons.append('BSE symbol downgrade')
            poor_count+=1
        
        if (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'Yes'):
            reasons.append('Market Cap/ISIN')
            poor_count+=1
        
        if (risk_rating_non_fno['Not traded?'][idx] == 'Yes'):
            reasons.append('Not traded')
            poor_count+=1
            
        reason_str = ''
        for reason in reasons[:-1]:
            reason_str = reason_str+reason+ ' , '
        if len(reasons)>0:
            reason_str = reason_str+reasons[-1]
        risk_rating_non_fno['Poor_Reasons'][idx] = reason_str
        risk_rating_non_fno['Poor_Reasons_Count'][idx] = poor_count

# In[106]:


#risk rating exceptions


# In[107]:


risk_rating_non_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'] = ''
risk_rating_non_fno['Meeting Size Criteria_Poor'] = ''
risk_rating_non_fno['Size Exception for Poor to Average'] = ''

risk_rating_non_fno['All Conditions of Good Met (Except "Good" Size)'] = ''
risk_rating_non_fno['Size Criteria_Average'] = ''
risk_rating_non_fno['Size Exception for Average to Good'] = ''

risk_rating_non_fno['All conditions met Except ADTO'] = ''
risk_rating_non_fno['Size Criteria for ADTO Upgrade'] = ''
risk_rating_non_fno['ADTO > 0.5 Cr'] = ''
risk_rating_non_fno['ADTO Upgrade (Poor to Average)'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    if risk_rating_non_fno['Rating v1'][idx] == 'Poor' :
        ####change below adto
        risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
        if ( (risk_rating_non_fno['Mcap'][idx]>=50 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=50 ) & (risk_rating_non_fno['ADTO in Cr'][idx]>1)  & ((adto_count(risk_rating_non_fno_row , 1))==3)   & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>100) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
           
            risk_rating_non_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'][idx] = 'Yes'
        else:
            risk_rating_non_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'][idx] = 'No'
        
        count_p = 0
        
        if (risk_rating_non_fno['Mcap'][idx]>=150 ):
            count_p+=1
        if (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=150 ):
            count_p+=1
        if (risk_rating_non_fno['Rev'][idx]>300):
            count_p+=1
        risk_rating_non_fno['Meeting Size Criteria_Poor'][idx] = count_p
    
        if ((risk_rating_non_fno['All Conditions Met (Except Size) + 50% of Average Size Criteria Met'][idx]=='Yes') & (risk_rating_non_fno['Meeting Size Criteria_Poor'][idx]>=2) ):
            risk_rating_non_fno['Size Exception for Poor to Average'][idx] = 'Upgrade'
    
    ####change below adto
    if risk_rating_non_fno['Rating v1'][idx] == 'Average' :
        risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
        if ((risk_rating_non_fno['PET Check'][idx]=='PET Good') &(risk_rating_non_fno['ADTO in Cr'][idx]>2) & ((adto_count(risk_rating_non_fno_row , 2))==3)  & (risk_rating_non_fno['Promoter Pledge%'][idx] < 25)) :
            risk_rating_non_fno['All Conditions of Good Met (Except "Good" Size)'][idx] = 'Yes'
        else:
            risk_rating_non_fno['All Conditions of Good Met (Except "Good" Size)'][idx] = 'No'
        
        count_a = 0
        if (risk_rating_non_fno['Mcap'][idx]>=750 ):
            count_a+=1
        if (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=750 ):
            count_a+=1
        if (risk_rating_non_fno['Rev'][idx]>1500):
            count_a+=1
        risk_rating_non_fno['Size Criteria_Average'][idx] = count_a
        
        if ((risk_rating_non_fno['All Conditions of Good Met (Except "Good" Size)'][idx]=='Yes') & (risk_rating_non_fno['Size Criteria_Average'][idx]>=2) ):
            risk_rating_non_fno['Size Exception for Average to Good'][idx] = 'Upgrade'
    
    
    
    if risk_rating_non_fno['Rating v1'][idx] == 'Poor' :
        if ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & (risk_rating_non_fno['Promoter Pledge%'][idx]<50)  & (risk_rating_non_fno['Impact Cost'][idx]<1) & (risk_rating_non_fno['Rev'][idx]>200) & ((risk_rating_non_fno['PET Check'][idx]=='PET Good') | (risk_rating_non_fno['F-Score'][idx]=='F-Check Passed')) & (risk_rating_non_fno['NSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['BSE symbol downgrade?'][idx] == "No") & (risk_rating_non_fno['Market Cap/ISIN'][idx] == 'No') & (risk_rating_non_fno['Not traded?'][idx] == 'No')  ): 
            risk_rating_non_fno['All conditions met Except ADTO'][idx] = 'Yes'
        else:
            risk_rating_non_fno['All conditions met Except ADTO'][idx] = 'No'
        
        
        count_p_adto = 0
        
        if (risk_rating_non_fno['Mcap'][idx]>=300 ):
            count_p_adto+=1
        if (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=300 ):
            count_p_adto+=1
        if (risk_rating_non_fno['Rev'][idx]>600):
            count_p_adto+=1
        risk_rating_non_fno['Size Criteria for ADTO Upgrade'][idx] = count_p_adto
        
        
        if (risk_rating_non_fno['ADTO in Cr'][idx]>0.5) :
            risk_rating_non_fno['ADTO > 0.5 Cr'][idx] = 'Yes'
        else:
            risk_rating_non_fno['ADTO > 0.5 Cr'][idx] = 'No'
        
        if ((risk_rating_non_fno['ADTO > 0.5 Cr'][idx]=='Yes') & (risk_rating_non_fno['Size Criteria for ADTO Upgrade'][idx]==3) & (risk_rating_non_fno['All conditions met Except ADTO'][idx] == 'Yes')):
            risk_rating_non_fno['ADTO Upgrade (Poor to Average)'][idx] = 'Upgrade'


# In[ ]:





# In[108]:


risk_rating_non_fno['ADTO_Wipsaw_Upgrade'] = ''
risk_rating_non_fno['Rating v1_5'] = ''
risk_rating_non_fno['Rating v2'] = ''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    if risk_rating_non_fno['Rating v1'][idx] == 'Good' :
        risk_rating_non_fno['Rating v2'][idx] = risk_rating_non_fno['Rating v1'][idx]
    elif risk_rating_non_fno['Rating v1'][idx]=='Poor' :
        if ((risk_rating_non_fno['Size Exception for Poor to Average'][idx]=='Upgrade') | (risk_rating_non_fno['ADTO Upgrade (Poor to Average)'][idx]=='Upgrade')):
            risk_rating_non_fno['Rating v2'][idx] = 'Average'
        else:
            risk_rating_non_fno['Rating v2'][idx] = risk_rating_non_fno['Rating v1'][idx]
    elif risk_rating_non_fno['Rating v1'][idx]=='Average' :
        if risk_rating_non_fno['Size Exception for Average to Good'][idx]=='Upgrade':
            risk_rating_non_fno['Rating v2'][idx] = 'Good'
        else:
            risk_rating_non_fno['Rating v2'][idx] = risk_rating_non_fno['Rating v1'][idx]


# In[109]:


risk_rating_non_fno['Rating v1_5'] = risk_rating_non_fno['Rating v2']


# In[110]:


#non fno wipsaw


# In[111]:


def non_fno_average_reasons_wipsaw(risk_rating_fno):
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=100:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 100:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>1)) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<50:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>200 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

    if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
        reasons.append('F-Score')

    if (risk_rating_fno['NSE symbol downgrade?'][idx] != "Yes"):
        reasons.append('NSE symbol downgrade')

    if (risk_rating_fno['BSE symbol downgrade?'][idx] != "Yes"):
        reasons.append('BSE symbol downgrade')

    if (risk_rating_fno['Market Cap/ISIN'][idx] != 'Yes'):
        reasons.append('Market Cap/ISIN')

    if (risk_rating_fno['Not traded?'][idx] != 'Yes'):
        reasons.append('Not traded')
    
    return reasons


# In[112]:


for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]]
    if risk_rating_non_fno['Rating v1'][idx] == 'Poor' :
        prev_month_row = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
        if len(prev_month_row)>0 :
            if prev_month_row['As per new Rules'][0] == 'Average':
                if risk_rating_non_fno['Poor_Reasons'][idx] == 'ADTO in Cr' :
                    adto_in_cr = 1
                    adto_thresh_count = adto_count(risk_rating_non_fno_row , adto_in_cr)
                    risk_rating_non_fno_row = risk_rating_non_fno_row.reset_index(drop=True)
                    adto_median_ = risk_rating_non_fno_row[latest_3_months_adto_cols]
                    adto_values = []
                    for col in adto_median_.columns:
                        adto_values.append(adto_median_[col][0])
                    adto_median = median(adto_values)
                    
                    if ( (adto_thresh_count ==2 ) & (adto_median > adto_in_cr) ):
#                         print(risk_rating_non_fno['CO_NAME'][idx])
                        risk_rating_non_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                        risk_rating_non_fno['Rating v2'][idx] = 'Average'   
    
    if risk_rating_non_fno['Rating v1'][idx] == 'Poor' :
        prev_month_row = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
        if len(prev_month_row)>0 :
            if prev_month_row['As per new Rules'][0] == 'Good':
                if risk_rating_non_fno['Poor_Reasons'][idx] == 'ADTO in Cr' :
                    adto_in_cr = 2
                    adto_thresh_count = adto_count(risk_rating_non_fno_row , adto_in_cr)
                    risk_rating_non_fno_row = risk_rating_non_fno_row.reset_index(drop=True)
                    adto_median_ = risk_rating_non_fno_row[latest_3_months_adto_cols]
                    adto_values = []
                    for col in adto_median_.columns:
                        adto_values.append(adto_median_[col][0])
                    adto_median = median(adto_values)
                    
                    if ( (adto_thresh_count ==2 ) & (adto_median > adto_in_cr) ):
#                         print(risk_rating_non_fno['CO_NAME'][idx])
                        risk_rating_non_fno['Rating v2'][idx] = 'Good'
                        risk_rating_non_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                    else:
                        adto_in_cr = 1
                        adto_thresh_count = adto_count(risk_rating_non_fno_row , adto_in_cr)
                        risk_rating_non_fno_row = risk_rating_non_fno_row.reset_index(drop=True)
                        adto_median_ = risk_rating_non_fno_row[latest_3_months_adto_cols]
                        adto_values = []
                        for col in adto_median_.columns:
                            adto_values.append(adto_median_[col][0])
                        adto_median = median(adto_values)
                        if ( (adto_thresh_count ==2 ) & (adto_median > adto_in_cr) ):
#                             print(risk_rating_non_fno['CO_NAME'][idx])
                            risk_rating_non_fno['Rating v2'][idx] = 'Average'
                            risk_rating_non_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                        
            
    if risk_rating_non_fno['Rating v1'][idx] == 'Average' :
        prev_month_row = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
        if len(prev_month_row)>0 :
            if prev_month_row['As per new Rules'][0] == 'Good':
                risk_rating_non_fno_row_ = risk_rating_non_fno_row.reset_index(drop=True)
                if non_fno_average_reasons_wipsaw(risk_rating_non_fno_row_) == ['ADTO in Cr'] :
                    adto_in_cr = 2
                    adto_thresh_count = adto_count(risk_rating_non_fno_row , adto_in_cr)
                    risk_rating_non_fno_row = risk_rating_non_fno_row.reset_index(drop=True)
                    adto_median_ = risk_rating_non_fno_row[latest_3_months_adto_cols]
                    adto_values = []
                    for col in adto_median_.columns:
                        adto_values.append(adto_median_[col][0])
                    adto_median = median(adto_values)
                    
                    if ( (adto_thresh_count == 2 ) & (adto_median > adto_in_cr) ):
#                         print(risk_rating_non_fno['CO_NAME'][idx])
                        risk_rating_non_fno['ADTO_Wipsaw_Upgrade'][idx] = 'Yes'
                        risk_rating_non_fno['Rating v2'][idx] = 'Good'   


# In[ ]:





# In[ ]:





# In[ ]:





# In[113]:


# risk_rating_non_fno.to_csv('test_nonfno.csv')


# In[ ]:





# In[ ]:





# In[ ]:





# In[114]:


risk_rating_non_fno['Poor_Alz_Manual'] = ''
risk_rating_non_fno['Poor_Alz_Manual_Comments'] = ''
risk_rating_non_fno['Poor_to_Avg_Manual'] = ''
risk_rating_non_fno['Good_BC_to_Avg_Manual'] = ''
risk_rating_non_fno['Manual_Rating'] = risk_rating_non_fno['Rating v2']
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    temp_1 = poor_downgrd_rated_df[poor_downgrd_rated_df['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if len(temp_1) > 0 :
        risk_rating_non_fno['Manual_Rating'][idx] = temp_1['Rating'][0]
        risk_rating_non_fno['Poor_Alz_Manual'][idx] = 'Downgraded_to_Poor'
        risk_rating_non_fno['Poor_Alz_Manual_Comments'][idx] = temp_1['Comments'][0]
    
    if risk_rating_non_fno['Rating v2'][idx] == 'Poor' :
        temp_2 = poor_to_avg_rated_df[poor_to_avg_rated_df['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
        if len(temp_2) > 0 :
            risk_rating_non_fno['Manual_Rating'][idx] = temp_2['Rating'][0]
            risk_rating_non_fno['Poor_to_Avg_Manual'][idx] = 'Upgraded_to_Average'
    
    if risk_rating_non_fno['Rating v2'][idx] in ['Bluechip','Good'] :
        temp_3 = good_bc_to_avg_rated_df[good_bc_to_avg_rated_df['CAPITALINE CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
        if len(temp_3) > 0 :
            risk_rating_non_fno['Manual_Rating'][idx] = temp_3['Rating'][0]
            risk_rating_non_fno['Good_BC_to_Avg_Manual'][idx] = 'Downgraded_to_Average'


# In[115]:


risk_rating_non_fno['As per new Rules'] = risk_rating_non_fno['Manual_Rating']
# added 27-09-2024 updated 25-11-2024
risk_rating_non_fno['As per new Rules - Restricted'] = risk_rating_non_fno['Restricted']
# for idx in range(0,len(risk_rating_non_fno['As per new Rules - Restricted'])):
#     if risk_rating_non_fno['As per new Rules'][idx] == 'Poor' :
#         risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
#         risk_rating_non_fno['As per new Rules - Restricted'][idx] = restricted_reason_fno_non_fno(row_df = risk_rating_non_fno_row , exclude='')

# In[116]:


risk_rating_non_fno['As per old Rules'] =''
risk_rating_non_fno['As per old Rules - Restricted'] =''
for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    risk_rating_non_fno_row = risk_rating_non_fno[risk_rating_non_fno['CO_CODE'] == risk_rating_non_fno['CO_CODE'][idx]].reset_index()
    if ( (risk_rating_non_fno['Mcap'][idx]>=500 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=500 ) & ( pet_check_old_rating(risk_rating_non_fno_row,10)==4 )   ): 
        risk_rating_non_fno['As per old Rules'][idx] = 'Good'
    elif ( (risk_rating_non_fno['Mcap'][idx]>=100 ) & (risk_rating_non_fno['[Total Shareholders Funds (Latest)]'][idx]>=100 ) & ( pet_check_old_rating(risk_rating_non_fno_row,2)==4 )   ): 
        risk_rating_non_fno['As per old Rules'][idx] = 'Average'
    else:
        risk_rating_non_fno['As per old Rules'][idx] = 'Poor'
        if (pet_check_old_rating_resticted(risk_rating_non_fno_row,0.5) ==4) :
            risk_rating_non_fno['As per old Rules - Restricted'][idx] = 'Restricted'


# In[117]:


risk_rating_non_fno['Current Live Rating'] = ''
risk_rating_non_fno['Current Live Restricted'] = ''
#new added 20-09-2024
risk_rating_non_fno[['BSE_VAR_pct','NSE_VAR_pct','ANGEL_VAR_pct','MTF_VAR_pct']] = np.nan

for idx in range(0,len(risk_rating_non_fno['CO_CODE'])):
    co_code = risk_rating_non_fno['CO_CODE'][idx]
    isin_t = bse_nse_group_sheet[bse_nse_group_sheet['CAPITALINE CODE'] == co_code].reset_index(drop=True)
    if len(isin_t) > 0:
        isin = isin_t['[ISIN No'][0]
        temp_curr_rating = angel_scrip_category_df[angel_scrip_category_df['ISIN No'] == isin].reset_index(drop=True)
        if len(temp_curr_rating) > 0 :
            risk_rating_non_fno['Current Live Rating'][idx] = temp_curr_rating['Angel scrip category'][0].strip()
            risk_rating_non_fno['Current Live Restricted'][idx] = temp_curr_rating['Restricted Scrips'][0].strip()
            #new added 20-09-2024
            risk_rating_non_fno['BSE_VAR_pct'][idx] = temp_curr_rating['BSE_VAR %'][0]#.strip()
            risk_rating_non_fno['NSE_VAR_pct'][idx] = temp_curr_rating['NSE_VAR %'][0]#.strip()
            risk_rating_non_fno['ANGEL_VAR_pct'][idx] = temp_curr_rating['ANGEL_VAR %'][0]#.strip()
            risk_rating_non_fno['MTF_VAR_pct'][idx] = temp_curr_rating['MTF_VAR %'][0]#.strip()


# In[118]:


#sheet 1


# In[119]:


close_price_df_nonfno = close_price_df.copy()


# In[120]:


close_price_df = pd.DataFrame()
close_price_df[['CO_CODE','CO_NAME']] = pet_check_code[['CO_CODE', 'CO_NAME']]
# close_price_df = close_price_df[~close_price_df['CO_CODE'].isin(risk_rating_fno['CO_CODE'])].reset_index(drop = True)

close_price_df['NSE'] = np.datetime64('NAT')
close_price_df['BSE'] = np.datetime64('NAT')
close_price_df['Latest'] = np.datetime64('NAT')
for idx in range(0,len(close_price_df['CO_CODE'])):
    temp_bse_cp = close_price_sheet_bse[close_price_sheet_bse['CAPITALINE CODE'] == close_price_df['CO_CODE'][idx] ].reset_index()
    temp_nse_cp = close_price_sheet_nse[close_price_sheet_nse['CAPITALINE CODE'] == close_price_df['CO_CODE'][idx] ].reset_index()
    if len(temp_bse_cp) > 0 :
        close_price_df['BSE'][idx] = temp_bse_cp['[Date (Latest)]'][0]
    else:
        close_price_df['BSE'][idx] = def_date
    
    if len(temp_nse_cp) > 0 :
        close_price_df['NSE'][idx] = temp_nse_cp['[Date (Latest)]'][0]
    else:
        close_price_df['NSE'][idx] = def_date
    
    close_price_df['Latest'][idx] = max( [close_price_df['BSE'][idx] ,close_price_df['NSE'][idx]] )


# In[121]:


financial_data_columns = ['[Year End (Latest)]',
       '[Year End (Latest1)]', '[Year End (Latest2)]', '[Year (Latest)]',
       '[Year (Latest1)]', '[Year (Latest2)]', '[Net Sales (Latest)]',
       '[Net Sales (Latest1)]', '[Net Sales (Latest2)]',
       '[PowerampFuel Cost (Latest)]', '[PowerampFuel Cost (Latest1)]',
       '[PowerampFuel Cost (Latest2)]', '[Employee Cost (Latest)]',
       '[Employee Cost (Latest1)]', '[Employee Cost (Latest2)]',
       '[Profit Before Tax (Latest)]', '[Profit Before Tax (Latest1)]',
       '[Profit Before Tax (Latest2)]', '[Reported Net Profit (Latest)]',
       '[Reported Net Profit (Latest1)]', '[Reported Net Profit (Latest2)]',
       '[Tax (Latest)]', '[Tax (Latest1)]', '[Tax (Latest2)]',
       '[Fringe Benefit tax (Latest)]', '[Fringe Benefit tax (Latest1)]',
       '[Fringe Benefit tax (Latest2)]', '[Deferred Tax (Latest)]',
       '[Deferred Tax (Latest1)]', '[Deferred Tax (Latest2)]',
       '[Operating Profit (Latest)]', '[Operating Profit (Latest1)]',
       '[Operating Profit (Latest2)]', '[Interest (Latest)]',
       '[Interest (Latest1)]', '[Interest (Latest2)]',
       '[Adjusted Net Profit (Latest)]', '[Adjusted Net Profit (Latest1)]',
       '[Adjusted Net Profit (Latest2)]', '[MODE (Latest)]',
       '[MODE (Latest1)]', '[MODE (Latest2)]',
       '[Total Shareholders Funds (Latest)]',
       '[Total Shareholders Funds (Latest1)]',
       '[Total Shareholders Funds (Latest2)]',
       '[Total Debt / Loan Funds (Latest)]',
       '[Total Debt / Loan Funds (Latest1)]',
       '[Total Debt / Loan Funds (Latest2)]',
       '[Cash and Bank Balance (Latest)]', '[Cash and Bank Balance (Latest1)]',
       '[Cash and Bank Balance (Latest2)]',
       '[Balance at Bank and Call Money (Latest)]',
       '[Balance at Bank and Call Money (Latest1)]',
       '[Balance at Bank and Call Money (Latest2)]', '[Total Assets (Latest)]',
       '[Total Assets (Latest1)]', '[Total Assets (Latest2)]',
       '[Net Cash from Operating Activities (Latest)]',
       '[Net Cash from Operating Activities (Latest1)]',
       '[Net Cash from Operating Activities (Latest2)]',
       '[Purchased of Fixed Assets (Latest)]',
       '[Purchased of Fixed Assets (Latest1)]',
       '[Purchased of Fixed Assets (Latest2)]',
       '[Sale of Fixed Assets (Latest)]', '[Sale of Fixed Assets (Latest1)]',
       '[Sale of Fixed Assets (Latest2)]', '[Capital Expenditure (Latest)]',
       '[Capital Expenditure (Latest1)]', '[Capital Expenditure (Latest2)]',
       '[capital WIP (Latest)]', '[capital WIP (Latest1)]',
       '[capital WIP (Latest2)]']


# In[122]:


adto_df_columns = adto_months_req + ['Median']


# In[123]:


sheet_1_raw_data = bse_nse_group_sheet.copy()


# In[124]:


sheet_1_raw_data['FnO'] = ''
sheet_1_raw_data[financial_data_columns] = np.nan
for col in adto_df_columns :
    sheet_1_raw_data[f'{col}_ADTO'] = np.nan
sheet_1_raw_data['ImpactCost_Min of NSE/BSE'] = np.nan
sheet_1_raw_data['Market Cap'] = np.nan
sheet_1_raw_data['Last_traded_date'] = np.datetime64('NAT')
sheet_1_raw_data['Promoter Pledge%'] = np.nan
sheet_1_raw_data['BSE_VAR %'] = np.nan
sheet_1_raw_data['NSE_VAR %'] = np.nan
sheet_1_raw_data['ANGEL_VAR %'] = np.nan
sheet_1_raw_data['MTF_VAR %'] = np.nan

for idx in range(0,len(sheet_1_raw_data['CAPITALINE CODE'])):
    temp_fin_data = financial_data[financial_data['CAPITALINE CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_fin_data) > 0 :
        for col in financial_data_columns :
            sheet_1_raw_data[col][idx] = temp_fin_data[col][0]
    temp_adto_data = adto_df[adto_df['CAPITALINE CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_adto_data) > 0 :
        for col in adto_df_columns :
            sheet_1_raw_data[f'{col}_ADTO'][idx] = temp_adto_data[col][0]
    
    temp_impact_cost = impact_cost[impact_cost['CO_CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_impact_cost) > 0 :
        sheet_1_raw_data['ImpactCost_Min of NSE/BSE'][idx] = temp_impact_cost['Min of NSE/BSE'][0]
    
    temp_fno_list_sheet = fno_list_sheet[fno_list_sheet['CAPITALINE CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_fno_list_sheet)>0 :
        sheet_1_raw_data['FnO'][idx] = 'FnO'
    else:
        sheet_1_raw_data['FnO'][idx] = ''
    
    temp_mcap_sheet = latest_mcap_sheet[latest_mcap_sheet['CAPITALINE CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_mcap_sheet) >0 : 
        sheet_1_raw_data['Market Cap'][idx] = temp_mcap_sheet['[Market Cap (Latest)]'][0]
    
    temp_close_price_df = close_price_df[close_price_df['CO_CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_close_price_df) >0 :
        sheet_1_raw_data['Last_traded_date'][idx] = temp_close_price_df['Latest'][0]
    
    temp_promoter_pledge = promoter_pledge_sheet[promoter_pledge_sheet['CAPITALINE CODE'] == sheet_1_raw_data['CAPITALINE CODE'][idx]].reset_index()
    if len(temp_promoter_pledge)>0:
        sheet_1_raw_data['Promoter Pledge%'][idx] = temp_promoter_pledge['[Total of Promoter and Group (Latest)]'][0]
    
    temp_angel_scrip_category = angel_scrip_category_df[angel_scrip_category_df['ISIN No'] == sheet_1_raw_data['[ISIN No'][idx]].reset_index()
    if len(temp_angel_scrip_category)>0:
        sheet_1_raw_data['BSE_VAR %'][idx] = temp_angel_scrip_category['BSE_VAR %'][0]
        sheet_1_raw_data['NSE_VAR %'][idx] = temp_angel_scrip_category['NSE_VAR %'][0]
        sheet_1_raw_data['ANGEL_VAR %'][idx] = temp_angel_scrip_category['ANGEL_VAR %'][0]
        sheet_1_raw_data['MTF_VAR %'][idx] = temp_angel_scrip_category['MTF_VAR %'][0]
    
#     if idx == 10:
#         break


# In[125]:


# asm_long_term_df_nse  asm_long_term_df_bse asm_short_term_df_bse


# In[126]:


sheet_1_raw_data['ASM_Stage'] = ''
for idx in range(0,len(sheet_1_raw_data['CAPITALINE CODE'])):
    asm_stage = ''
    isin_no = sheet_1_raw_data['[ISIN No'][idx]
    temp_nse_asm = asm_long_term_df_nse[asm_long_term_df_nse['ISIN \n'] == isin_no ].reset_index()
    temp_bse_asm_l = asm_long_term_df_bse[asm_long_term_df_bse['ISIN'] == isin_no ].reset_index()
    temp_bse_asm_s = asm_short_term_df_bse[asm_short_term_df_bse['ISIN'] == isin_no ].reset_index()
    
    if len(temp_nse_asm) > 0 :
        asm_stage = temp_nse_asm['ASM STAGE \n'][0].split(' (')[0]
        sheet_1_raw_data['ASM_Stage'][idx] = asm_stage
    elif len(temp_bse_asm_l) > 0 :
        asm_stage = 'LTASM - '+temp_bse_asm_l['ASM Stage'][0]
        sheet_1_raw_data['ASM_Stage'][idx] = asm_stage
    elif len(temp_bse_asm_s) > 0 :
        asm_stage = 'STASM - '+temp_bse_asm_s['Short Term 5/15 Days ASM Stage'][0]
        sheet_1_raw_data['ASM_Stage'][idx] = asm_stage
    else:
        ...


# In[127]:


# sheet_2_curr_prev_rating


# In[128]:


# prev_month_data_fno_sheet = prev_month_data_fno_sheet.rename(columns= {"CO_CODE" :"CAPITALINE CODE"})
# prev_month_data_non_fno_sheet = prev_month_data_non_fno_sheet.rename(columns= {"CO_CODE" :"CAPITALINE CODE"})


# In[129]:


sheet_2_data = bse_nse_group_sheet[['CAPITALINE CODE', 'CO_NAME']]


# In[130]:


sheet_2_data['Current_Rating'] = np.nan
sheet_2_data['Previous_Rating'] = np.nan
sheet_2_data['Current Live Rating'] = np.nan
sheet_2_data['As per old Rules'] = np.nan

for idx in range(0,len(sheet_2_data['CAPITALINE CODE'])):
    co_code = sheet_2_data['CAPITALINE CODE'][idx]
    prev_rating_fno = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE']==co_code].reset_index()
    prev_rating_non_fno = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE']==co_code].reset_index()
    if len(prev_rating_fno) > 0 :
        sheet_2_data['Previous_Rating'][idx] = prev_rating_fno['Manual_Rating'][0]
    elif len(prev_rating_non_fno) > 0 :
        sheet_2_data['Previous_Rating'][idx] = prev_rating_non_fno['Manual_Rating'][0]
    
    curr_risk_rating_fno = risk_rating_fno[risk_rating_fno['CO_CODE']==co_code].reset_index()
    curr_risk_rating_non_fno = risk_rating_non_fno[risk_rating_non_fno['CO_CODE']==co_code].reset_index()
    
    if len(curr_risk_rating_fno)>0 :
        sheet_2_data['Current_Rating'][idx] = curr_risk_rating_fno['Manual_Rating'][0]
        sheet_2_data['Current Live Rating'][idx] = curr_risk_rating_fno['Current Live Rating'][0]
        sheet_2_data['As per old Rules'][idx] = curr_risk_rating_fno['As per old Rules'][0]
    elif len(curr_risk_rating_non_fno) > 0 :
        sheet_2_data['Current_Rating'][idx] = curr_risk_rating_non_fno['Manual_Rating'][0]
        sheet_2_data['Current Live Rating'][idx] = curr_risk_rating_non_fno['Current Live Rating'][0]
        sheet_2_data['As per old Rules'][idx] = curr_risk_rating_non_fno['As per old Rules'][0]


# In[131]:


sheet_3_data = sheet_2_data[sheet_2_data['Previous_Rating'].isna()].reset_index(drop=True)


# In[132]:


sheet_45_data = sheet_2_data[~sheet_2_data['Previous_Rating'].isna()].reset_index(drop=True)


# In[133]:


sheet_45_data = sheet_45_data[sheet_45_data['Current_Rating']!=sheet_45_data['Previous_Rating']].reset_index(drop=True)


# In[134]:


# ['Bluechip','Good','Average','Poor']


# In[135]:


sheet_45_data['curr_rank'] = np.nan
sheet_45_data['prev_rank'] = np.nan
for idx in range(0,len(sheet_45_data)):
    if sheet_45_data['Current_Rating'][idx] == 'Bluechip':
        sheet_45_data['curr_rank'][idx] = 1
    elif sheet_45_data['Current_Rating'][idx] == 'Good':
        sheet_45_data['curr_rank'][idx] = 2 
    elif sheet_45_data['Current_Rating'][idx] == 'Average':
        sheet_45_data['curr_rank'][idx] = 3
    elif sheet_45_data['Current_Rating'][idx] == 'Poor':
        sheet_45_data['curr_rank'][idx] = 4
    
    
    
    if sheet_45_data['Previous_Rating'][idx] == 'Bluechip':
        sheet_45_data['prev_rank'][idx] = 1
    elif sheet_45_data['Previous_Rating'][idx] == 'Good':
        sheet_45_data['prev_rank'][idx] = 2 
    elif sheet_45_data['Previous_Rating'][idx] == 'Average':
        sheet_45_data['prev_rank'][idx] = 3
    elif sheet_45_data['Previous_Rating'][idx] == 'Poor':
        sheet_45_data['prev_rank'][idx] = 4


# In[136]:


sheet_4_data_upgrade = sheet_45_data[sheet_45_data['curr_rank']<sheet_45_data['prev_rank']][['CAPITALINE CODE', 'CO_NAME', 'Current_Rating', 'Previous_Rating','Current Live Rating','As per old Rules']].reset_index(drop=True)


# In[137]:


sheet_5_data_downgrade = sheet_45_data[sheet_45_data['curr_rank']>sheet_45_data['prev_rank']][['CAPITALINE CODE', 'CO_NAME', 'Current_Rating', 'Previous_Rating','Current Live Rating','As per old Rules']].reset_index(drop=True)


# In[138]:


##


# In[139]:


def fno_bluechip_reasons(risk_rating_fno):
    risk_rating_fno = risk_rating_fno.rename(columns = { '6M Median ADTO in Rs crs' : 'ADTO in Cr' ,
                                                       'Latest_3M_ADTO_above_thresh':'3M_ADTO_COUNT',
                                                        
                                                    'Revenue- Latest FY': 'Rev' ,
                                                    'Rating v1 (as per rules)' : 'Rating v1',
                                                    'Rating v2 (V1 adj for whipsaw)': 'Rating v2',
                                                    'Rating v2 (V1 wo adj for whipsaw)' : 'Rating v1_5'  ,
                                                       })
    
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=5000:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 2000:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>20) ) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<25:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>5000 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

    try :
        if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
            reasons.append('F-Score')

        if risk_rating_fno['Nifty 500'][idx]=='Nifty 500':
            reasons.append('Nifty 500')
    except:
        ...
    try :
        if (risk_rating_fno['NSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('NSE symbol downgrade')

        if (risk_rating_fno['BSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('BSE symbol downgrade')

        if (risk_rating_fno['Market Cap/ISIN'][idx] != 'Yes'):
            reasons.append('Market Cap/ISIN')

        if (risk_rating_fno['Not traded?'][idx] != 'Yes'):
            reasons.append('Not traded')
    except:
        ...
    
    return reasons


# In[140]:


def fno_good_reasons(risk_rating_fno):
    risk_rating_fno = risk_rating_fno.rename(columns = { '6M Median ADTO in Rs crs' : 'ADTO in Cr' ,
                                                       'Latest_3M_ADTO_above_thresh':'3M_ADTO_COUNT',
                                                        
                                                    'Revenue- Latest FY': 'Rev' ,
                                                    'Rating v1 (as per rules)' : 'Rating v1',
                                                    'Rating v2 (V1 adj for whipsaw)': 'Rating v2',
                                                    'Rating v2 (V1 wo adj for whipsaw)' : 'Rating v1_5'  ,
                                                       })
    
    
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=500:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 500:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>15)) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<40:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>100 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

    try :
        if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
            reasons.append('F-Score')

        if risk_rating_fno['Nifty 500'][idx]=='Nifty 500':
            reasons.append('Nifty 500')
    except:
        ...
    try :
        if (risk_rating_fno['NSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('NSE symbol downgrade')

        if (risk_rating_fno['BSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('BSE symbol downgrade')

        if (risk_rating_fno['Market Cap/ISIN'][idx] != 'Yes'):
            reasons.append('Market Cap/ISIN')

        if (risk_rating_fno['Not traded?'][idx] != 'Yes'):
            reasons.append('Not traded')
    except:
        ...
    
    return reasons


# In[141]:


def fno_average_reasons(risk_rating_fno):
    risk_rating_fno = risk_rating_fno.rename(columns = { '6M Median ADTO in Rs crs' : 'ADTO in Cr' ,
                                                       'Latest_3M_ADTO_above_thresh':'3M_ADTO_COUNT',
                                                        
                                                    'Revenue- Latest FY': 'Rev' ,
                                                    'Rating v1 (as per rules)' : 'Rating v1',
                                                    'Rating v2 (V1 adj for whipsaw)': 'Rating v2',
                                                    'Rating v2 (V1 wo adj for whipsaw)' : 'Rating v1_5'  ,
                                                       })
    
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=100:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 100:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>2) ) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<50:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>200 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

    try :
        if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
            reasons.append('F-Score')

        if risk_rating_fno['Nifty 500'][idx]=='Nifty 500':
            reasons.append('Nifty 500')
    except:
        ...
    try :
        if (risk_rating_fno['NSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('NSE symbol downgrade')

        if (risk_rating_fno['BSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('BSE symbol downgrade')

        if (risk_rating_fno['Market Cap/ISIN'][idx] != 'Yes'):
            reasons.append('Market Cap/ISIN')

        if (risk_rating_fno['Not traded?'][idx] != 'Yes'):
            reasons.append('Not traded')
    except:
        ...
    return reasons


# In[142]:


#non fno


# In[143]:


def non_fno_good_reasons(risk_rating_fno):
    risk_rating_fno = risk_rating_fno.rename(columns = { '6M Median ADTO in Rs crs' : 'ADTO in Cr' ,
                                                       'Latest_3M_ADTO_above_thresh':'3M_ADTO_COUNT',
                                                        
                                                    'Revenue- Latest FY': 'Rev' ,
                                                    'Rating v1 (as per rules)' : 'Rating v1',
                                                    'Rating v2 (V1 adj for whipsaw)': 'Rating v2',
                                                    'Rating v2 (V1 wo adj for whipsaw)' : 'Rating v1_5'  ,
                                                       })
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=500:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 500:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>2)) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<25:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>1000 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

#     if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
#         reasons.append('F-Score')
    try :
        if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
            reasons.append('F-Score')

        if risk_rating_fno['Nifty 500'][idx]=='Nifty 500':
            reasons.append('Nifty 500')
    except:
        ...
    try :
        if (risk_rating_fno['NSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('NSE symbol downgrade')

        if (risk_rating_fno['BSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('BSE symbol downgrade')

        if (risk_rating_fno['Market Cap/ISIN'][idx] != 'Yes'):
            reasons.append('Market Cap/ISIN')

        if (risk_rating_fno['Not traded?'][idx] != 'Yes'):
            reasons.append('Not traded')
    except:
        ...

    return reasons


# In[144]:


def non_fno_average_reasons(risk_rating_fno):
    risk_rating_fno = risk_rating_fno.rename(columns = { '6M Median ADTO in Rs crs' : 'ADTO in Cr' ,
                                                       'Latest_3M_ADTO_above_thresh':'3M_ADTO_COUNT',
                                                        
                                                    'Revenue- Latest FY': 'Rev' ,
                                                    'Rating v1 (as per rules)' : 'Rating v1',
                                                    'Rating v2 (V1 adj for whipsaw)': 'Rating v2',
                                                    'Rating v2 (V1 wo adj for whipsaw)' : 'Rating v1_5'  ,
                                                       })
    
    idx = 0
    reasons = []
       
    if risk_rating_fno['Mcap'][idx]>=100:
        reasons.append('Mcap')

    if risk_rating_fno['[Total Shareholders Funds (Latest)]'][idx] >= 100:
        reasons.append('Shareholders Fund')

    if ((risk_rating_fno['ADTO in Cr'][idx]>1)) | pd.isna(risk_rating_fno['ADTO in Cr'][idx]) | (risk_rating_fno['3M_ADTO_COUNT'][idx]<3) :
        reasons.append('ADTO in Cr')

    if risk_rating_fno['Promoter Pledge%'][idx]<50:
        reasons.append('Promoter Pledge%')

    if risk_rating_fno['Impact Cost'][idx]<1:
        reasons.append('Impact Cost')

    if risk_rating_fno['Rev'][idx]>200 :
        reasons.append('Rev')

    if risk_rating_fno['PET Check'][idx]=='PET Good' :
        reasons.append('PET Check')

    if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
        reasons.append('F-Score')
    
    try :
        if risk_rating_fno['F-Score'][idx]=='F-Check Passed' :
            reasons.append('F-Score')

        if risk_rating_fno['Nifty 500'][idx]=='Nifty 500':
            reasons.append('Nifty 500')
    except:
        ...
    try :
        if (risk_rating_fno['NSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('NSE symbol downgrade')

        if (risk_rating_fno['BSE symbol downgrade?'][idx] != "Yes"):
            reasons.append('BSE symbol downgrade')

        if (risk_rating_fno['Market Cap/ISIN'][idx] != 'Yes'):
            reasons.append('Market Cap/ISIN')

        if (risk_rating_fno['Not traded?'][idx] != 'Yes'):
            reasons.append('Not traded')
    except:
        ...
    return reasons


# In[145]:


sheet_4_data_upgrade['Reasons'] = ''
for idx in range(0,len(sheet_4_data_upgrade['CAPITALINE CODE'])):
    co_code = sheet_4_data_upgrade['CAPITALINE CODE'][idx]
    fno_flag = False
    prev_rating_fno = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE']==co_code].reset_index(drop=True)
    prev_rating_non_fno = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE']==co_code].reset_index(drop=True)
    if len(prev_rating_fno) > 0 :
        prev_rating = prev_rating_fno
        fno_flag = True
    elif len(prev_rating_non_fno) > 0 :
        prev_rating = prev_rating_non_fno
        fno_flag = False
    
    curr_risk_rating_fno = risk_rating_fno[risk_rating_fno['CO_CODE']==co_code].reset_index(drop=True)
    curr_risk_rating_non_fno = risk_rating_non_fno[risk_rating_non_fno['CO_CODE']==co_code].reset_index(drop=True)
    
    if len(curr_risk_rating_fno)>0 :
        curr_rating = curr_risk_rating_fno
        fno_flag = True
    elif len(curr_risk_rating_non_fno) > 0 :
        curr_rating = curr_risk_rating_non_fno
        fno_flag = False
    
    reason = []
    #for fno stocks
    if fno_flag == True:
        if curr_rating['Manual_Rating'][0] == 'Bluechip':
            curr_reasons = fno_bluechip_reasons(curr_rating)
            prev_reasons = fno_bluechip_reasons(prev_rating)
            
            reasons = [r for r in curr_reasons if r not in prev_reasons]
        
        if curr_rating['Manual_Rating'][0] == 'Good':
            curr_reasons = fno_good_reasons(curr_rating)
            prev_reasons = fno_good_reasons(prev_rating)
            
            reasons = [r for r in curr_reasons if r not in prev_reasons]
        
        if curr_rating['Manual_Rating'][0] == 'Average':
            curr_reasons = fno_average_reasons(curr_rating)
            prev_reasons = fno_average_reasons(prev_rating)
            
            reasons = [r for r in curr_reasons if r not in prev_reasons]
        
        if reasons == [] :
            
            if curr_rating['Size Exception for Poor to Average'][0] == 'Upgrade': 
                reasons = ['Size Exception for Poor to Average']
            if curr_rating['Size Exception for Average to Good'][0] == 'Upgrade':
                reasons = ['Size Exception for Average to Good']
#             if curr_rating['ADTO Upgrade (Poor to Average)'][0] == 'Upgrade':
#                 reasons = ['ADTO Upgrade (Poor to Average)']
        if reasons == [] :
            reasons = ['ADTO in Cr']
    
    #for nonfno stocks
    else:
     
        if curr_rating['Manual_Rating'][0] == 'Good':
            curr_reasons = non_fno_good_reasons(curr_rating)
            prev_reasons = non_fno_good_reasons(prev_rating)
            
            reasons = [r for r in curr_reasons if r not in prev_reasons]
        
        if curr_rating['Manual_Rating'][0] == 'Average':
            curr_reasons = non_fno_average_reasons(curr_rating)
            prev_reasons = non_fno_average_reasons(prev_rating)
            
            reasons = [r for r in curr_reasons if r not in prev_reasons]
    
        if reasons == [] :
            
            if curr_rating['Size Exception for Poor to Average'][0] == 'Upgrade': 
                reasons = ['Size Exception for Poor to Average']
            if curr_rating['Size Exception for Average to Good'][0] == 'Upgrade':
                reasons = ['Size Exception for Average to Good']
            if curr_rating['ADTO Upgrade (Poor to Average)'][0] == 'Upgrade':
                reasons = ['ADTO Upgrade (Poor to Average)']
        if reasons == [] :
            reasons = ['ADTO in Cr']
    
    reason_str = ''
    for reason in reasons[:-1]:
        reason_str = reason_str+reason+ ' , '
    if len(reasons)>0:
        reason_str = reason_str+reasons[-1]
    sheet_4_data_upgrade['Reasons'][idx] = reason_str
#     if idx == 0:
#         break


# In[146]:


# risk_rating_fno.columns


# In[147]:


# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)


# In[148]:


sheet_5_data_downgrade['Reasons'] = ''
for idx in range(0,len(sheet_5_data_downgrade['CAPITALINE CODE'])):
    co_code = sheet_5_data_downgrade['CAPITALINE CODE'][idx]
    fno_flag = False
    prev_rating_fno = prev_month_data_fno_sheet[prev_month_data_fno_sheet['CO_CODE']==co_code].reset_index(drop=True)
    prev_rating_non_fno = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE']==co_code].reset_index(drop=True)
    if len(prev_rating_fno) > 0 :
        prev_rating = prev_rating_fno
        fno_flag = True
    elif len(prev_rating_non_fno) > 0 :
        prev_rating = prev_rating_non_fno
        fno_flag = False
    
    curr_risk_rating_fno = risk_rating_fno[risk_rating_fno['CO_CODE']==co_code].reset_index(drop=True)
    curr_risk_rating_non_fno = risk_rating_non_fno[risk_rating_non_fno['CO_CODE']==co_code].reset_index(drop=True)
    
    if len(curr_risk_rating_fno)>0 :
        curr_rating = curr_risk_rating_fno
        fno_flag = True
    elif len(curr_risk_rating_non_fno) > 0 :
        curr_rating = curr_risk_rating_non_fno
        fno_flag = False
    
    reason = []
    #for fno stocks
    if fno_flag == True:
        if prev_rating['Manual_Rating'][0] == 'Bluechip':
            curr_reasons = fno_bluechip_reasons(curr_rating)
            prev_reasons = fno_bluechip_reasons(prev_rating)
            
            reasons = [r for r in prev_reasons if r not in curr_reasons]
        
        if prev_rating['Manual_Rating'][0] == 'Good':
            curr_reasons = fno_good_reasons(curr_rating)
            prev_reasons = fno_good_reasons(prev_rating)
            
            reasons = [r for r in prev_reasons if r not in curr_reasons]
        
        if prev_rating['Manual_Rating'][0] == 'Average':
            curr_reasons = fno_average_reasons(curr_rating)
            prev_reasons = fno_average_reasons(prev_rating)
            
            reasons = [r for r in prev_reasons if r not in curr_reasons]
        
        if reasons == [] :
            try :
                if prev_rating['Size Exception for Poor to Average'][0] == 'Upgrade': 
                    reasons = ['Prev_Rating_Size Exception for Poor to Average']
            except:
                ...
            try:
                if prev_rating['Size Exception for Average to Good'][0] == 'Upgrade':
                    reasons = ['Prev_Rating_Size Exception for Average to Good']
            except:
                ...
            # if prev_rating['ADTO Upgrade (Poor to Average)'][0] == 'Upgrade':
            #     reasons = ['Prev_Rating_ADTO Upgrade (Poor to Average)']
    
        if reasons == [] :
            reasons = ['ADTO in Cr']
    
    #for nonfno stocks
    else:
     
        if prev_rating['Manual_Rating'][0] == 'Good':
            curr_reasons = non_fno_good_reasons(curr_rating)
            prev_reasons = non_fno_good_reasons(prev_rating)
            
            reasons = [r for r in prev_reasons if r not in curr_reasons]
        
        if prev_rating['Manual_Rating'][0] == 'Average':
            curr_reasons = non_fno_average_reasons(curr_rating)
            prev_reasons = non_fno_average_reasons(prev_rating)
            
            reasons = [r for r in prev_reasons if r not in curr_reasons]
    
        if reasons == [] :
            try:
                if prev_rating['Size Exception for Poor to Average'][0] == 'Upgrade': 
                    reasons = ['Prev_Rating_Size Exception for Poor to Average']
            except:
                ...
            try:
                if prev_rating['Size Exception for Average to Good'][0] == 'Upgrade':
                    reasons = ['Prev_Rating_Size Exception for Average to Good']
            except:
                ...
            try:
                if prev_rating['ADTO Upgrade (Poor to Average)'][0] == 'Upgrade':
                    reasons = ['Prev_Rating_ADTO Upgrade (Poor to Average)']
            except:
                ...
        if reasons == [] :
            reasons = ['ADTO in Cr']
    
    reason_str = ''
    for reason in reasons[:-1]:
        reason_str = reason_str+reason+ ' , '
    if len(reasons)>0:
        reason_str = reason_str+reasons[-1]
    sheet_5_data_downgrade['Reasons'][idx] = reason_str
#     if idx == 0:
#         break


# In[149]:


restricted_data = bse_nse_group_sheet[['CAPITALINE CODE', 'CO_NAME']]
restricted_data['Current_Restricted'] = ''
restricted_data['Previous_Restricted'] = ''
restricted_data['Current Live Rating'] = ''
restricted_data['As per old Rules'] = ''
for idx in range(0,len(restricted_data['CAPITALINE CODE'])):
    co_code = restricted_data['CAPITALINE CODE'][idx]
    prev_rating_non_fno = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE']==co_code].reset_index()

    if len(prev_rating_non_fno) > 0 :
        if pd.isna(prev_rating_non_fno['Restricted'][0]) :
            ...
        else:
            restricted_data['Previous_Restricted'][idx] = prev_rating_non_fno['Restricted'][0]
    
    curr_risk_rating_non_fno = risk_rating_non_fno[risk_rating_non_fno['CO_CODE']==co_code].reset_index()
    
    if len(curr_risk_rating_non_fno) > 0 :
        restricted_data['Current_Restricted'][idx] = curr_risk_rating_non_fno['Restricted'][0]
        restricted_data['Current Live Rating'][idx] = curr_risk_rating_non_fno['Current Live Rating'][0]
        restricted_data['As per old Rules'][idx] = curr_risk_rating_non_fno['As per old Rules'][0]


# In[150]:


# newly restricted in curr month
up_restricted_data = restricted_data[restricted_data['Previous_Restricted']=='']
up_restricted_data = up_restricted_data[up_restricted_data["Current_Restricted"]!=''].reset_index(drop=True)


# In[151]:


# restricted prev month but now removed in curr month.
down_restricted_data = restricted_data[restricted_data['Current_Restricted']=='']
down_restricted_data = down_restricted_data[down_restricted_data["Previous_Restricted"]!=''].reset_index(drop=True)


# In[152]:


def restricted_reasons(risk_rating_fno):
    risk_rating_fno = risk_rating_fno.rename(columns = { '6M Median ADTO in Rs crs' : 'ADTO in Cr' ,
                                                       'Latest_3M_ADTO_above_thresh':'3M_ADTO_COUNT',
                                                        
                                                    'Revenue- Latest FY': 'Rev' ,
                                                    'Rating v1 (as per rules)' : 'Rating v1',
                                                    'Rating v2 (V1 adj for whipsaw)': 'Rating v2',
                                                    'Rating v2 (V1 wo adj for whipsaw)' : 'Rating v1_5'  ,
                                                       })
    idx = 0
    reasons = []
       
    if risk_rating_fno['BSE Series'][idx] in ['Z','P','M']:
        reasons.append('BSE Series')
    
    if (risk_rating_fno['NSE Series'][idx] == 'BZ'):
        reasons.append('NSE Series')
    
    if ((risk_rating_fno['ADTO in Cr'][idx] < 0.2) ) | (pd.isna(risk_rating_fno['ADTO in Cr'][idx])) :
        reasons.append('ADTO in Cr')
    
    return reasons


# In[153]:


#         if ((risk_rating_non_fno['BSE Series'][idx] in ['Z','P','M'] ) |  (risk_rating_non_fno['NSE Series'][idx] == 'BZ' ) | (risk_rating_non_fno['ADTO in Cr'][idx] < 0.2)


# In[154]:


down_restricted_data['Reasons'] = ''
for idx in range(0,len(down_restricted_data['CAPITALINE CODE'])):
    
    co_code = down_restricted_data['CAPITALINE CODE'][idx]
    
    curr_rating = risk_rating_non_fno[risk_rating_non_fno['CO_CODE']==co_code].reset_index(drop=True)
    prev_rating = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE']==co_code].reset_index(drop=True)
    

    curr_reasons = restricted_reasons(curr_rating)
    prev_reasons = restricted_reasons(prev_rating)
            
    reasons = [r for r in prev_reasons if r not in curr_reasons]
    reason_str = ''
    for reason in reasons[:-1]:
        reason_str = reason_str+reason+ ' , '
    if len(reasons)>0:
        reason_str = reason_str+reasons[-1]
    down_restricted_data['Reasons'][idx] = reason_str


# In[ ]:





# In[155]:


up_restricted_data['Reasons'] = ''
for idx in range(0,len(up_restricted_data['CAPITALINE CODE'])):
    reasons = []
    co_code = up_restricted_data['CAPITALINE CODE'][idx]
    
    curr_rating = risk_rating_non_fno[risk_rating_non_fno['CO_CODE']==co_code].reset_index(drop=True)
    prev_rating = prev_month_data_non_fno_sheet[prev_month_data_non_fno_sheet['CO_CODE']==co_code].reset_index(drop=True)
    

    curr_reasons = restricted_reasons(curr_rating)
    if len(prev_rating) > 0 :
        prev_reasons = restricted_reasons(prev_rating)
    
    if len(prev_rating) > 0 :
        reasons = [r for r in curr_reasons if r not in prev_reasons]
    else:
        reasons = ['Prev_Month_Missing']
    
    reason_str = ''
    for reason in reasons[:-1]:
        reason_str = reason_str+reason+ ' , '
    if len(reasons)>0:
        reason_str = reason_str+reasons[-1]
    up_restricted_data['Reasons'][idx] = reason_str


# In[ ]:





# In[ ]:





# In[ ]:





# In[156]:


sheet6_manual_intervention = pd.DataFrame()


# In[157]:


temp_fno_manual = risk_rating_fno[risk_rating_fno['Rating v2']!=risk_rating_fno['Manual_Rating']][['CO_CODE', 'CO_NAME','Rating v2','Poor_Alz_Manual', 'Poor_Alz_Manual_Comments', 'Poor_to_Avg_Manual','Good_BC_to_Avg_Manual', 'Manual_Rating']].reset_index(drop=True)


# In[158]:


temp_nonfno_manual =risk_rating_non_fno[risk_rating_non_fno['Rating v2']!=risk_rating_non_fno['Manual_Rating']][['CO_CODE', 'CO_NAME','Rating v2','Poor_Alz_Manual', 'Poor_Alz_Manual_Comments', 'Poor_to_Avg_Manual','Good_BC_to_Avg_Manual', 'Manual_Rating']].reset_index(drop=True)


# In[159]:


sheet6_manual_intervention = pd.concat([temp_fno_manual,temp_nonfno_manual],ignore_index=True)


# In[ ]:





# In[160]:


# output_dir = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\Output\Risk_Rating_{month}.xlsx'


# In[161]:


#final columns rename before writing to excel.


# In[162]:


adto_df_columns = adto_months_req


# In[163]:


risk_rating_fno = risk_rating_fno.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                     'Rating v1_5' :'Rating v2 (V1 wo adj for whipsaw)' ,
                                                   })

for col in adto_df_columns :
    risk_rating_fno = risk_rating_fno.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[164]:


risk_rating_non_fno = risk_rating_non_fno.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                    'Rating v1_5' :'Rating v2 (V1 wo adj for whipsaw)' ,
                                                        })

for col in adto_df_columns :
    risk_rating_non_fno = risk_rating_non_fno.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[165]:


pet_check_code = pet_check_code.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    pet_check_code = pet_check_code.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[166]:


f_score_code = f_score_code.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    f_score_code = f_score_code.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[167]:


sheet_1_raw_data = sheet_1_raw_data.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    sheet_1_raw_data = sheet_1_raw_data.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[168]:


adto_df = adto_df.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    adto_df = adto_df.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[169]:


sheet_2_data = sheet_2_data.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    sheet_2_data = sheet_2_data.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[170]:


sheet_3_data = sheet_3_data.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    sheet_3_data = sheet_3_data.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[171]:


sheet_4_data_upgrade = sheet_4_data_upgrade.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    sheet_4_data_upgrade = sheet_4_data_upgrade.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[172]:


sheet_5_data_downgrade = sheet_5_data_downgrade.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    sheet_5_data_downgrade = sheet_5_data_downgrade.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[173]:


up_restricted_data = up_restricted_data.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    up_restricted_data = up_restricted_data.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[174]:


down_restricted_data = down_restricted_data.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    down_restricted_data = down_restricted_data.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[175]:


sheet6_manual_intervention = sheet6_manual_intervention.rename(columns = { 'ADTO in Cr':'6M Median ADTO in Rs crs',
                                                    'Rev' :'Revenue- Latest FY',
                                                    'Rating v1':'Rating v1 (as per rules)',
                                                    '3M_ADTO_COUNT':'Latest_3M_ADTO_above_thresh' ,
                                                    'Rating v2':'Rating v2 (V1 adj for whipsaw)',
                                                        })

for col in adto_df_columns :
    sheet6_manual_intervention = sheet6_manual_intervention.rename(columns = {f'{col}_ADTO': f'{col}_ADTO in Rs crs' })


# In[176]:

#new added 20-09-2024
# 
latest_3_months_adto_cols_req = [ column+' in Rs crs' for column in latest_3_months_adto_cols ]
risk_rating_fno_columns_sequence =  ['CO_CODE', 'CO_NAME', 'F&O','[ISIN No', 'Current Live Rating',
       'Current Live Restricted', 'As per new Rules',
       'As per new Rules - Restricted','New Rules Rating (ex Symbol)','New Rules Rating (ex Symbol) Restricted',
      'New Rules Rating (ex ADTO)', 'New Rules Rating (ex ADTO) Restricted','New Rules Rating (ex Symbol, ex ADTO)',
      'New Rules - Quick Rating','Poor_Reasons_Count','Percent_Funding' ,'BSE_VAR_pct', 'NSE_VAR_pct',
       'ANGEL_VAR_pct','Poor_Reasons', 'New Rules Restricted Reason',
      'Mcap','[Total Shareholders Funds (Latest)]','Revenue- Latest FY' , '6M Median ADTO in Rs crs','PET Check', 'F-Score',
      'Promoter Pledge%','Impact Cost',
         'BSE Series','NSE Series', 'Nifty 500',
          ] + latest_3_months_adto_cols_req + [
           'Poor_Alz_Manual',
       'Poor_Alz_Manual_Comments', 'Poor_to_Avg_Manual',
       'Good_BC_to_Avg_Manual', 'Manual_Rating',
          'Rating v1 (as per rules)', 'Latest_3M_ADTO_above_thresh',
           'All Conditions Met (Except Size) + 50% of Average Size Criteria Met',
       'Meeting Size Criteria_Poor', 'Size Exception for Poor to Average',
       'All Conditions of Good Met (Except "Good" Size)',
       'Size Criteria_Average', 'Size Exception for Average to Good',
       'ADTO_Wipsaw_Upgrade',  'Rating v2 (V1 wo adj for whipsaw)',
 'Rating v2 (V1 adj for whipsaw)' ,  'MTF_VAR_pct', 'As per old Rules',
       'As per old Rules - Restricted']


#rearrange non fno sheets : 

risk_rating_non_fno_columns_new_sequence = ['CO_CODE', 'CO_NAME','[ISIN No', 'Current Live Rating',
       'Current Live Restricted', 'As per new Rules',
       'As per new Rules - Restricted','New Rules Rating (ex Symbol)', 'New Rules Rating (ex Symbol) Restricted' ,'New Rules Rating (ex ADTO)', 'New Rules Rating (ex ADTO) Restricted',
       'New Rules Rating (ex Symbol, ex ADTO)', 'New Rules - Quick Rating', 'Poor_Reasons_Count' ,'Percent_Funding'
         , 'BSE_VAR_pct', 'NSE_VAR_pct','ANGEL_VAR_pct','Poor_Reasons', 'New Rules Restricted Reason'
         , 'Mcap','[Total Shareholders Funds (Latest)]','Revenue- Latest FY', '6M Median ADTO in Rs crs', 'PET Check', 'F-Score',
     'Promoter Pledge%',  'Impact Cost','BSE Series', 'NSE Series' ] + latest_3_months_adto_cols_req + [ 'Poor_Alz_Manual','Poor_Alz_Manual_Comments', 'Poor_to_Avg_Manual',
       'Good_BC_to_Avg_Manual', 'Manual_Rating',
        'BSE symbol downgrade?', 'NSE symbol downgrade?',
       'Market Cap/ISIN',  
       'Rating v1 (as per rules)', 'Restricted',
       'Latest_3M_ADTO_above_thresh', 
       'All Conditions Met (Except Size) + 50% of Average Size Criteria Met',
       'Meeting Size Criteria_Poor', 'Size Exception for Poor to Average',
       'All Conditions of Good Met (Except "Good" Size)',
       'Size Criteria_Average', 'Size Exception for Average to Good',
       'All conditions met Except ADTO', 'Size Criteria for ADTO Upgrade',
       'ADTO > 0.5 Cr', 'ADTO Upgrade (Poor to Average)',
       'ADTO_Wipsaw_Upgrade', 'Rating v2 (V1 wo adj for whipsaw)',
       'Rating v2 (V1 adj for whipsaw)', 'Last Trading Date?', 'Not traded?', 'MTF_VAR_pct', 'As per old Rules',
       'As per old Rules - Restricted' ]
risk_rating_non_fno_to_wr = risk_rating_non_fno[risk_rating_non_fno_columns_new_sequence] 
risk_rating_fno_to_wr = risk_rating_fno[risk_rating_fno_columns_sequence]

out_path = output_dir
writer = pd.ExcelWriter(out_path , engine='xlsxwriter', datetime_format='dd/mm/yyyy')
risk_rating_fno_to_wr.to_excel(writer, sheet_name='Risk Rating_F&O', index=False)
risk_rating_non_fno_to_wr.to_excel(writer, sheet_name='Risk Rating_Non F&O', index=False)
pet_check_code.to_excel(writer, sheet_name='PET_CHECK',index=False)
f_score_code.to_excel(writer, sheet_name='F_SCORE',index=False)
sheet_1_raw_data.to_excel(writer, sheet_name='Data_Points_Curr',index=False)
prev_month_data_data_points.to_excel(writer, sheet_name='Data_Points_Prev',index=False)
adto_df.to_excel(writer, sheet_name='ADTO_Data',index=False)
sheet_2_data.to_excel(writer, sheet_name='Curr_Prev_Rating',index=False)
sheet_3_data.to_excel(writer, sheet_name='New_Curr_Rating',index=False)
sheet_4_data_upgrade.to_excel(writer, sheet_name='Rating_Upgrade',index=False)
sheet_5_data_downgrade.to_excel(writer, sheet_name='Rating_Downgrade',index=False)
up_restricted_data.to_excel(writer, sheet_name='Restriction_added',index=False)
down_restricted_data.to_excel(writer, sheet_name='Restriction_removed',index=False)
sheet6_manual_intervention.to_excel(writer, sheet_name='Manual_Rating',index=False)
writer.close()


# In[177]:


print( (time.time() - cstart)/60 )


# In[178]:


# 10 mins run time


# In[ ]:




