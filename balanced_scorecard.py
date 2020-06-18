import os
import pandas as pd

master = pd.DataFrame()

def build_13_month_dates():
    # builds an array of 13 months going back from the previous month
    dates = []
    month = pd.Timestamp.now() - pd.DateOffset(months=1)
    dates.append(month.strftime('%b-%y'))
    for i in range(12):
        month = month - pd.DateOffset(months = 1)
        dates.append(month.strftime('%b-%y'))
    return dates

def find_file(date):
    # finds the file from the relevant snapshot folder as appropriate
    files = os.listdir('/media/wdrive/Workforce Monthly Reports/Monthly_Reports/' + date +
                       ' Snapshot/')

    # deals with all the different variations of the balance scorecard names
    for i in files:
        if "GG&C_Balanced_Scorecard - "+date in i:
            print(date + " - found - method 1")
            return '/media/wdrive/Workforce Monthly Reports/Monthly_Reports/' + date + ' Snapshot/' + i, date
        elif "GGC_Balanced_Scorecard - "+date in i:
            print(date + " - found - method 2")
            return '/media/wdrive/Workforce Monthly Reports/Monthly_Reports/' + date + ' Snapshot/' + i, date
        elif "GGC Balanced_Scorecard - "+date in i:
            print(date + " - found - method 3")
            return '/media/wdrive/Workforce Monthly Reports/Monthly_Reports/' + date + ' Snapshot/' + i, date
        elif "GGC Balanced Scorecard - "+date in i:
            print(date + " - found - method 4")
            return '/media/wdrive/Workforce Monthly Reports/Monthly_Reports/' + date + ' Snapshot/' + i, date
        elif "GGC Area Balanced Scorecard - "+date in i:
            print(date + " - found - method 5")
            return '/media/wdrive/Workforce Monthly Reports/Monthly_Reports/' + date + ' Snapshot/' + i, date
    print(files)

def open_and_pull(file, date):
    report_date = pd.to_datetime(date, format='%b-%y').strftime('%Y/%m')
    print(report_date)

    df = pd.read_excel(file)
    df = df[df['Sector/Directorate/HSCP'] != 'Non Paid Employees']
    df = df[df['Sector/Directorate/HSCP'] != 'GP Trainees']
    if "Report Date" in df.columns:
        df = df[df['Report Date'] == report_date]
    else:
        df = df[df['Report date'] == report_date]
    print(len(df))
    return df

def pivot_maker(df, value):
    pivot = pd.pivot_table(df, index='Sector/Directorate/HSCP', values=value, aggfunc=np.sum)
    return pivot






master = pd.DataFrame()
dates = build_13_month_dates()
for month in dates:
    file, date = find_file(month)
    # date = find_file(month)[1]
    df = open_and_pull(file, date)
    master = master.append(df, ignore_index=True)
    print(len(master))
print(master.columns)




master.to_csv('/home/danny/workspace/abs-full.csv', index=False)