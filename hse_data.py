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
    regex = re.compile(r'[\d]{8} - HSE Sharps.xlsx')
    files = [i for i in files if regex.match(i)]
    print(f'Length of files after regex: {len(files)}')
    return files


def find_HSE_file(date):
    """Finds the earliest date in a given month that the report was run and returns that filename"""
    reformatted_date = pd.to_datetime(date, format='%b-%y').strftime('%Y%m')
    regex = re.compile(f'{reformatted_date}[\d]{{2}} - HSE Sharps.xlsx')
    relevant_files = [pd.to_datetime(i[0:8]) for i in files if regex.match(i)]
    print(f'Month: {date} - dates of report = {relevant_files}')
    min_date = min(relevant_files).strftime('%Y%m%d')
    filename = f'/media/wdrive/Learnpro/HSE Sharps and Skins/{min_date} - HSE Sharps.xlsx'
    print(f'Most relevant file - {filename}')
    return filename


def open_HSE_file(filename):
    """Opens the relevant file and selects relevant portion of staff"""
    df = pd.read_excel(filename, sheet_name='Export')
    return df


def GGC_compliance(df):
    """Select the relevant in-scope staff, then work out compliance rate among those staff"""
    initial_len = len(df)
    df = df[df['GGC Module'].isin(['Complete', 'Expired', 'Not Undertaken', 'No Account'])]
    scope_len = len(df)
    print(f'Initial length: {initial_len}, After removal of out of scope: {scope_len}')
    compliant = df[df['GGC Module'] == 'Complete']
    compliance_percentage = round(len(compliant) / scope_len * 100, 1)
    print(f'GGC compliance percentage : {compliance_percentage}')


def NES_compliance(df):
    """Same structure as GGC compliance"""
    initial_len = len(df)
    df = df[df['NES Module'].isin(['Complete', 'Expired', 'Not Undertaken', 'No Account'])]
    scope_len = len(df)
    print(f'Initial length: {initial_len}, After removal of out of scope: {scope_len}')
    compliant = df[df['NES Module'] == 'Complete']
    compliance_percentage = round(len(compliant) / scope_len * 100, 1)
    print(f'NES compliance percentage : {compliance_percentage}')


dates = build_13_month_dates()

# get all files
files = clean_reg_files(os.listdir('/media/wdrive/Learnpro/HSE Sharps and Skins'))

# loop through dates and produce compliance percentages
for date in dates:
    filename = find_HSE_file(date)
    df = open_HSE_file(filename)
    GGC_compliance(df)
    NES_compliance(df)
