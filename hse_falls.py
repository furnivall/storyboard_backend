import pandas as pd
import os
import re

def build_13_month_dates():
    """builds an array of 13 months going back from the previous month"""
    dates = []
    month = pd.Timestamp.now() - pd.DateOffset(months=2)
    dates.append(month.strftime('%Y-%m'))
    for i in range(12):
        month = month - pd.DateOffset(months=1)
        dates.append(month.strftime('%Y-%m'))
    return dates


def find_HSE_file(date):
    """Finds the earliest date in a given month that the report was run and returns that filename"""
    reformatted_date = pd.to_datetime(date, format='%Y-%m').strftime('%Y%m')
    regex = re.compile(f'{reformatted_date}[\d]{{2}} - HSE Falls.xlsx')
    relevant_files = [pd.to_datetime(i[0:8]) for i in all_files if regex.match(i)]
    min_date = min(relevant_files).strftime('%Y%m%d')
    filename = f'/media/wdrive/Learnpro/HSE Falls/{min_date} - HSE Falls.xlsx'
    return filename

def get_data(df):
    inscope = len(df[df['Falls Compliant'].isin(['Not Undertaken', 'Complete', 'No Account', 'Not Complete'])])
    print(df['Falls Compliant'].value_counts())
    complete = len(df[df['Falls Compliant'] == 'Complete'])
    return inscope, complete

# list of all files in HR CMS folder
all_files = os.listdir('/media/wdrive/Learnpro/HSE Falls')

# regex match to remove all weird different files
regex = re.compile(r'[\d]{8} - HSE Falls.xlsx')

print(len(all_files))
# list comprehension to implement regex
all_files = [i for i in all_files if regex.match(i)]

print(all_files)

print(len(all_files))

abs_13mo = pd.read_excel(
    '/media/wdrive/workforce monthly reports/monthly_reports/May-20 Snapshot/GG&C_Balanced_Scorecard_13m - May-20.xlsx')
depts = abs_13mo['Department'].unique().tolist()
master_falls = pd.DataFrame()
dates = build_13_month_dates()
today = pd.Timestamp.now().strftime('%Y%m%d')
for date in dates:
    file = find_HSE_file(date)
    print(f'{date} - {file}')
    df = pd.read_excel(file, sheet_name='Export')
    print(len(df))
    print(df['Falls Compliant'].value_counts())
    for dept in depts:
        curr_df = df[df['department'] == dept]
        output = get_data(curr_df)
        dataframe_line = {'Department': dept, 'Report Date': pd.to_datetime(date, format='%Y-%m'),
                          'Falls Inscope': output[0],
                          'Falls Compliant': output[1]
                          }
        master_falls = master_falls.append(dataframe_line, ignore_index=True)
master_falls.to_excel('/media/wdrive/storyboards)

