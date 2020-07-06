import pandas as pd
import re
import os


def build_13_month_dates():
    """builds an array of 13 months going back from the previous month"""
    dates = []
    month = pd.Timestamp.now() - pd.DateOffset(months=2)
    dates.append(month.strftime('%Y-%m'))
    for i in range(12):
        month = month - pd.DateOffset(months=1)
        dates.append(month.strftime('%Y-%m'))
    return dates


# list of all files in HR CMS folder
all_files = os.listdir('/media/wdrive/Learnpro/M&H')

# regex match to remove all weird different files
regex = re.compile(r'[\d]{4}-[\d]{2} - Staff Download - GGC.xlsx')

print(len(all_files))
# list comprehension to implement regex
all_files = [i for i in all_files if regex.match(i)]

print(all_files)

df = pd.read_excel('/media/wdrive/Learnpro/M&H/' + all_files[0], sheet_name='GGC')
abs_13mo = pd.read_excel(
    '/media/wdrive/workforce monthly reports/monthly_reports/May-20 Snapshot/GG&C_Balanced_Scorecard_13m - May-20.xlsx')
depts = abs_13mo['Department'].unique().tolist()
print(df.columns)

print(df['Scope'].value_counts())
print(df['Assessed-Total'].value_counts())


def get_data(df):
    inscope = df[df['Scope'] == 1]

    threeplus = len(inscope[inscope['Assessed-Total'] > 2])
    onetwo = len(inscope[(inscope['Assessed-Total'] > 0) & (inscope['Assessed-Total'] < 3)])
    zero = len(inscope[inscope['Assessed-Total'] == 0])
    # debug
    # print(f"Number inscope - {len(inscope)}, 3+: {threeplus}, 1-2: {onetwo}, zero:{zero}")
    return [len(inscope), threeplus, onetwo, zero]


dates = build_13_month_dates()
mh_master = pd.DataFrame()
today = pd.Timestamp.now().strftime('%Y%m%d')
for date in dates:
    count_inscope = 0
    count_threeplus = 0
    count_onetwo = 0
    count_zero = 0
    print(date)
    try:
        df = pd.read_excel('/media/wdrive/learnpro/M&H/' + date + ' - Staff Download - GGC.xlsx', sheet_name='GGC')
        for dept in depts:
            curr_df = df[df['department'] == dept]

            output = get_data(curr_df)
            count_inscope = count_inscope + output[0]
            count_threeplus = count_threeplus + output[1]
            count_onetwo = count_onetwo + output[2]
            count_zero = count_zero + output[3]
            dataframe_line = {'Department': dept, 'Report Date': pd.to_datetime(date, format='%Y-%m'),
                              'Moving & Handling - In scope': output[0],
                              'Moving & Handling - 3+ completions': output[1],
                              'Moving & Handling - 1-2 completions': output[2],
                              'Moving & Handling - 0 completions': output[3],
                              }
            mh_master = mh_master.append(dataframe_line, ignore_index=True)

        print(f'Month: {date} - Total inscope - {count_inscope}, total three+ - {count_threeplus}, '
              f'total 1-2 - {count_onetwo}, total zeroes {count_zero}')
    except FileNotFoundError:
        mh_master.to_excel('/media/wdrive/storyboards/hse-mh-' + today + '.xlsx', index=False)
        exit()
mh_master.to_excel('/media/wdrive/storyboards/hse-mh-' + today + '.xlsx', index=False)
