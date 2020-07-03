import os
import pandas as pd
import re

# list of all files in HR CMS folder
all_files = os.listdir('/media/HR-CMS/HR CMS/Administration/Extracts')

# regex match to remove all weird different files
regex = re.compile(r'HRCMS_Data_Extract_[\d]{8}.xlsx')

# list comprehension to implement regex
all_files = [i for i in all_files if regex.match(i)]


def build_13_month_dates():
    """builds an array of 13 months going back from the previous month"""
    dates = []
    month = pd.Timestamp.now() - pd.DateOffset(months=1)
    dates.append(month.strftime('%b-%y'))
    for i in range(12):
        month = month - pd.DateOffset(months=1)
        dates.append(month.strftime('%b-%y'))
    return dates


def find_relevant_file(date):
    """Get list of data extracts for given month, then select earliest one"""
    relevant_files = []
    date = pd.to_datetime(date, format='%b-%y').strftime('%Y%m')
    for i in all_files:
        # drop the nonsense files
        if "_Culled" in i:
            continue
        elif "Extract_" + date not in i:
            continue
        else:
            # add files to a list
            relevant_files.append(pd.to_datetime(i[19:27]))
    # find the earliest reporting date in the month in question
    min_date = min(relevant_files).strftime('%Y%m%d')
    final_filename = '/media/HR-CMS/HR CMS/Administration/Extracts/HRCMS_Data_Extract_' + min_date + '.xlsx'
    return final_filename


def open_file_read_data(filename):
    """Opens file and replaces garbage data"""
    df = pd.read_excel(filename)
    df.replace({'EHealth': 'eHealth',
                "Women & Children'S": "Women & Children's",
                'Diagnostic Services': 'Diagnostics Directorate',
                'Human Resources & Organisational Development': 'HR and OD',
                'Glasgow City Hscp': 'Glasgow City HSCP',
                'Support Services - Partnership': 'Estates and Facilities',
                'Estates And Facilities': 'Estates and Facilities',
                'Ppfm': 'Estates and Facilities',
                'Support Services - Acute': 'Estates and Facilities',
                'Support Services': 'Estates and Facilities',
                'Domestic Services': 'Estates and Facilities',
                'Finance Services - Partnership': 'Finance',
                'East Dunbartonshire Hscp': 'East Dunbartonshire HSCP',
                'Hr And Od': 'HR and OD',
                'Nursing Director': 'Board Nurse Director',
                'Human Resources - Corporate': 'HR and OD',
                'North': 'North Sector'
                }, inplace=True)
    return df


def get_bullying_cases(df):
    """Gets bullying cases"""
    df = df[df['Active Case?'] == 'Yes']
    return (df[df['Category'] == 'Dignity At Work'])


def get_suspended_8w(df):
    """Pulls active cases where colleague has been suspended for more than 8 weeks"""
    df = df[df['Active Case?'] == 'Yes']
    df = df[df['No. Of Weeks Suspended (To-Date)'].astype(float) >= 8]
    return df


def get_dismissals(df, date):
    """Pull all dismissals in the given month"""
    date = pd.to_datetime(date, format='%b-%y')
    df['Date Concluded'] = pd.to_datetime(df['Date Concluded'], errors='coerce')
    dateindex = date.month
    df = df[df['Date Concluded'].dt.month == dateindex]
    df = df[df['Er Outcome'] == 'Dismissal']
    return df


def get_grievances(df):
    """Gets df full of grievances"""
    df = df[df['Active Case?'] == 'Yes']
    df = df[df['Category'] == 'Grievance']
    return df


def get_disciplinaries(df):
    """Get df full of disciplinaries"""
    df = df[df['Active Case?'] == 'Yes']
    df = df[df['Category'] == 'Disciplinary']
    return df


dates = build_13_month_dates()

cms_master = pd.DataFrame()

# for sector in ['Clyde Sector', 'Diagnostics Directorate', 'North Sector', 'Regional Services', 'South Sector',
#                "Women & Children's", 'Acute Corporate', 'Board Administration', 'Board Medical Director',
#                'Board Nurse Director', 'Centre For Population Health', 'Corporate Communications', 'eHealth',
#                'Estates and Facilities', 'Finance', 'HR and OD', 'Public Health', 'East Dunbartonshire - Oral Health',
#                'East Dunbartonshire HSCP', 'East Renfrewshire HSCP', 'Glasgow City HSCP', 'Inverclyde HSCP',
#                'Renfrewshire HSCP', 'West Dunbartonshire HSCP', 'Diagnostic Services', 'Acute Directors']:

abs_13mo = pd.read_excel('/media/wdrive/workforce monthly reports/monthly_reports/May-20 Snapshot/GG&C_Balanced_Scorecard_13m - May-20.xlsx')
depts = abs_13mo['Department'].unique().tolist()


for date in dates:
    filename = find_relevant_file(date)
    df = open_file_read_data(filename)
    print(filename)
    for dept in depts:
        print(dept)
        curr_df = df[df['Ward/ Department'] == dept]
        dataframe_line = {'Department': dept, 'Report Date': pd.to_datetime(date, format='%b-%y'),
                          'Dismissals': len(get_dismissals(curr_df, date)),
                          'Suspensions > 8 weeks': len(get_suspended_8w(curr_df)),
                          'Grievances': len(get_grievances(curr_df)), 'Disciplinaries': len(get_disciplinaries(curr_df)),
                          'Bullying Cases': len(get_bullying_cases(curr_df))}

        # build single-line dataframe for given date/sector combination
        current_df = pd.DataFrame({k: [v] for k, v in dataframe_line.items()})

    # add to master file
        cms_master = cms_master.append(current_df, ignore_index=True)

    # print completion dialog
        print(f'Department: {dept}, date: {date} - Complete')


# output data
today = pd.Timestamp.now().strftime('%Y%m%d')
cms_master.to_excel('/media/wdrive/storyboards/cms_data'+today+'.xlsx', index=False)
