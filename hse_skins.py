import pandas as pd
import os
import re

def build_13_month_dates():
    """builds an array of 13 months going back from the previous month"""
    dates = []
    month = pd.Timestamp.now() - pd.DateOffset(months=1)
    dates.append(month.strftime('%b-%y'))
    for i in range(12):
        month = month - pd.DateOffset(months=1)
        dates.append(month.strftime('%b-%y'))
    return dates


def clean_reg_files(files):
    # regex match to remove all weird different files
    print(f'Length of files before regex: {len(files)}')
    regex = re.compile(r'[\d]{8} - HSE Skins.xlsx')
    files = [i for i in files if regex.match(i)]
    print(f'Length of files after regex: {len(files)}')
    return files


def find_hse_file(date):
    """Finds the earliest date in a given month that the report was run and returns that filename"""
    reformatted_date = pd.to_datetime(date, format='%b-%y').strftime('%Y%m')
    regex = re.compile(f'{reformatted_date}[\d]{{2}} - HSE Skins.xlsx')
    relevant_files = [pd.to_datetime(i[0:8]) for i in files if regex.match(i)]
    min_date = min(relevant_files).strftime('%Y%m%d')
    filename = f'W:/Learnpro/HSE Sharps and Skins/{min_date} - HSE Skins.xlsx'
    return filename

def get_data(df):
    respers = len(df[df['GGC Managing Skin Care for Responsible Persons'].notnull()])
    managers = len(df[df['GGC: Managing Skin Care at Work for Managers'].notnull()])
    return respers, managers

dates = build_13_month_dates()

files = clean_reg_files(os.listdir('W:/Learnpro/HSE Sharps and Skins'))




#
abs_13mo = pd.read_excel(
    'W:/workforce monthly reports/monthly_reports/May-20 Snapshot/GG&C_Balanced_Scorecard_13m - May-20.xlsx')
depts = abs_13mo['Department'].unique().tolist()
master_skins = pd.DataFrame()
for date in dates:
    print(date)
    try:
        file = find_hse_file(date)
        df = pd.read_excel(file, sheet_name='Export')
        print(len(df))
        for dept in depts:
            curr_df = df[df['department'] == dept]
            output = get_data(curr_df)
            dataframe_line = {'Department': dept, 'Report Date': pd.to_datetime(date, format='%b-%y'),
                              'HSE Skins - Responsible Persons compliant': output[0],
                              'HSE Skins - Managers compliant': output[1]
                              }
            print(dataframe_line)
            master_skins = master_skins.append(dataframe_line, ignore_index=True)
    except ValueError:
        print("no file found")
today = pd.Timestamp.now().strftime('%Y%m%d')
master_skins.to_excel('W:/Storyboards/hse-skins-' + today + '.xlsx', index=False)
