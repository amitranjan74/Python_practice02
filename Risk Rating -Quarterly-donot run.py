import time
cstart = time.time()

#update to current month
month = 'May-2025'

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
from scipy import stats 
from statistics import mean,median
import math
import time
import pickle
import os
print('Start time: ',dt.datetime.now())

##update 6 months only - remove oldest month and add latest excel name
# adto_months_req = ['March 2024','April 2024','May 2024','June 2024','July 2024','August 2024']


# latest_3_months_adto = adto_months_req[3:]
# latest_3_months_adto_cols = []
# for col in latest_3_months_adto:
#     latest_3_months_adto_cols.append(f'{col}_ADTO')
# latest_3_months_adto_cols_req = [ column+' in Rs crs' for column in latest_3_months_adto_cols ]
input_excel_risk_rating = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\Output\Risk_Rating_{month}.xlsx'

output_dir = rf'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Code\{month}\Output\Quarterly Rating Review {month}.xlsx'


email_rating_excel = r'W:\Knowledge repositiory\Master folder\Fundamental Work\Risk rating\Stock Risk Rating\Risk_FY24.xlsx'
email_rating_df = pd.read_excel(email_rating_excel, sheet_name = 'Log')

date_filter  = datetime.now() - relativedelta.relativedelta(days=20)
email_rating_df_filtered = email_rating_df[email_rating_df['Date'] >=date_filter].reset_index(drop=True)
# New Category

risk_rating_fno = pd.read_excel(input_excel_risk_rating , sheet_name='Risk Rating_F&O')

risk_rating_non_fno = pd.read_excel(input_excel_risk_rating , sheet_name='Risk Rating_Non F&O')

score_dict = {
    'Bluechip' : 1 ,
    'Good':2 ,
    'Average':3 ,
    'Poor' : 4,
    'Restricted' : 1,
    '-' : 0,
    '' : 0 ,
}

risk_rating_fno['Data Check'] = 1
risk_rating_fno['Downgrade/Upgrade/Maintain'] = ''
risk_rating_fno['Email Rating'] = ''
risk_rating_fno['Email Rating - Restricted'] = ''
for idx, row in risk_rating_fno.iterrows():
    checker_flag = 1
    if (  (pd.isna(row['[Total Shareholders Funds (Latest)]'])) or (row['[Total Shareholders Funds (Latest)]'] == '') ) :
        checker_flag = 0
    if (  (pd.isna(row['Revenue- Latest FY'])) or (row['Revenue- Latest FY'] == '') ) :
        checker_flag = 0
    if (  (pd.isna(row['Mcap'])) or (row['Mcap'] == '') ) :
        checker_flag = 0
    if ( (row['PET Check'] == 'PET Good') or  (row['F-Score'] == 'F-Check Passed')) :
        ...
    else:
        checker_flag = 0
    
    risk_rating_fno['Data Check'][idx] = checker_flag

    email_rating_df_filtered_req = email_rating_df_filtered[email_rating_df_filtered['ISIN NO'] == row['[ISIN No']].reset_index(drop=True)
    if len(email_rating_df_filtered_req) > 0 :
        risk_rating_fno['Email Rating'][idx] = email_rating_df_filtered_req['New Category'][0]
        risk_rating_fno['Email Rating - Restricted'][idx] = email_rating_df_filtered_req['New Status'][0]


    decision = ''
    row[['Current Live Rating','Current Live Restricted','As per new Rules','As per new Rules - Restricted']] = row[['Current Live Rating','Current Live Restricted','As per new Rules','As per new Rules - Restricted']].fillna('')
    CLR_score = score_dict[row['Current Live Rating']]
    CLRR_score = score_dict[row['Current Live Restricted']]
    APNR_score = score_dict[row['As per new Rules']]
    APNRR_score = score_dict[row['As per new Rules - Restricted']]

    if (CLR_score+ CLRR_score ) == (APNR_score+APNRR_score) :
        decision = 'Maintain'
    elif (CLR_score+ CLRR_score ) > (APNR_score+APNRR_score) :
        decision = 'Upgrade'
    elif (CLR_score+ CLRR_score ) < (APNR_score+APNRR_score) :
        decision = 'Downgrade'

    risk_rating_fno['Downgrade/Upgrade/Maintain'][idx] = decision


risk_rating_non_fno['Data Check'] = 1
risk_rating_non_fno['Downgrade/Upgrade/Maintain'] = ''
risk_rating_non_fno['Email Rating'] = ''
risk_rating_non_fno['Email Rating - Restricted'] = ''
for idx, row in risk_rating_non_fno.iterrows():
    checker_flag = 1
    if (  (pd.isna(row['[Total Shareholders Funds (Latest)]'])) or (row['[Total Shareholders Funds (Latest)]'] == '') ) :
        checker_flag = 0
    if (  (pd.isna(row['Revenue- Latest FY'])) or (row['Revenue- Latest FY'] == '') ) :
        checker_flag = 0
    if (  (pd.isna(row['Mcap'])) or (row['Mcap'] == '') ) :
        checker_flag = 0
    if ( (row['PET Check'] == 'PET Good') or  (row['F-Score'] == 'F-Check Passed')) :
        ...
    else:
        checker_flag = 0
    
    risk_rating_non_fno['Data Check'][idx] = checker_flag

    email_rating_df_filtered_req = email_rating_df_filtered[email_rating_df_filtered['ISIN NO'] == row['[ISIN No']].reset_index(drop=True)
    if len(email_rating_df_filtered_req) > 0 :
        risk_rating_non_fno['Email Rating'][idx] = email_rating_df_filtered_req['New Category'][0]
        risk_rating_non_fno['Email Rating - Restricted'][idx] = email_rating_df_filtered_req['New Status'][0]


    decision = ''
    row[['Current Live Rating','Current Live Restricted','As per new Rules','As per new Rules - Restricted']] = row[['Current Live Rating','Current Live Restricted','As per new Rules','As per new Rules - Restricted']].fillna('')
    CLR_score = score_dict[row['Current Live Rating']]
    CLRR_score = score_dict[row['Current Live Restricted']]
    APNR_score = score_dict[row['As per new Rules']]
    APNRR_score = score_dict[row['As per new Rules - Restricted']]

    if (CLR_score+ CLRR_score ) == (APNR_score+APNRR_score) :
        decision = 'Maintain'
    elif (CLR_score+ CLRR_score ) > (APNR_score+APNRR_score) :
        decision = 'Upgrade'
    elif (CLR_score+ CLRR_score ) < (APNR_score+APNRR_score) :
        decision = 'Downgrade'

    risk_rating_non_fno['Downgrade/Upgrade/Maintain'][idx] = decision

# print(risk_rating_fno.columns)

# print(risk_rating_non_fno.columns)


risk_rating_fno_columns_sequence = [
    "CO_NAME",
    "[ISIN No",
    "Data Check",
    "Downgrade/Upgrade/Maintain",
    "Current Live Rating",
    "Current Live Restricted",
    "As per new Rules",
    "As per new Rules - Restricted",
    "New Rules Rating (ex Symbol)",
    "New Rules Rating (ex Symbol) Restricted",
    "New Rules Rating (ex ADTO)",
    "New Rules Rating (ex ADTO) Restricted",
    "New Rules Rating (ex Symbol, ex ADTO)",
    "New Rules - Quick Rating",
    "Email Rating",
    "Email Rating - Restricted",
    "Mcap",
    "[Total Shareholders Funds (Latest)]",
    "Revenue- Latest FY",
    "PET Check",
    "F-Score",
    "Promoter Pledge%",
    "Percent_Funding",
    "BSE_VAR_pct",
    "NSE_VAR_pct",
    "ANGEL_VAR_pct",
    "6M Median ADTO in Rs crs",
    "Impact Cost",
    "BSE Series",
    "NSE Series",
    "Poor_Alz_Manual",
    "Poor_to_Avg_Manual",
    "Good_BC_to_Avg_Manual"
]



#rearrange non fno sheets : 

risk_rating_non_fno_columns_new_sequence = [
    "CO_NAME",
    "[ISIN No",
    "Data Check",
    "Downgrade/Upgrade/Maintain",
    "Current Live Rating",
    "Current Live Restricted",
    "As per new Rules",
    "As per new Rules - Restricted",
    "New Rules Rating (ex Symbol)",
    "New Rules Rating (ex Symbol) Restricted",
    "New Rules Rating (ex ADTO)",
    "New Rules Rating (ex ADTO) Restricted",
    "New Rules Rating (ex Symbol, ex ADTO)",
    "New Rules - Quick Rating",
    "Email Rating",
    "Email Rating - Restricted",
    "Mcap",
    "[Total Shareholders Funds (Latest)]",
    "Revenue- Latest FY",
    "PET Check",
    "F-Score",
    "Promoter Pledge%",
    "Percent_Funding",
    "BSE_VAR_pct",
    "NSE_VAR_pct",
    "ANGEL_VAR_pct",
    "6M Median ADTO in Rs crs",
    "Impact Cost",
    "BSE Series",
    "NSE Series",
    "Poor_Alz_Manual",
    "Poor_to_Avg_Manual",
    "Good_BC_to_Avg_Manual"
]

risk_rating_non_fno_to_wr = risk_rating_non_fno[risk_rating_non_fno_columns_new_sequence] 
risk_rating_fno_to_wr = risk_rating_fno[risk_rating_fno_columns_sequence]

out_path = output_dir
writer = pd.ExcelWriter(out_path , engine='xlsxwriter', datetime_format='dd/mm/yyyy')
risk_rating_fno_to_wr.to_excel(writer, sheet_name='F&O', index=False)
risk_rating_non_fno_to_wr.to_excel(writer, sheet_name='Non F&O', index=False)
writer.close()

print( (time.time() - cstart)/60 )