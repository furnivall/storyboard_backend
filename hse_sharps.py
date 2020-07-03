import pandas as pd
import os
import re

#TODO : amend HSE Sectors below:
# sectors = {'Non Paid Employees', 'Board Medical Director', 'Finance', 'eHealth', 'Clyde Sector', 'Renfrewshire HSCP',
#            'East Renfrewshire HSCP', 'GP Trainees', 'South Sector', 'Acute Directors', 'Inverclyde HSCP',
#            "Women & Children's", 'Board Nurse Director', 'East Dunbartonshire Hscp', 'Centre For Population Health',
#            'Board Administration', 'HR and OD', 'Glasgow City HSCP', 'Acute Corporate', 'Diagnostics Directorate',
#            'Estates and Facilities', 'Public Health', 'West Dunbartonshire HSCP', 'East Dunbartonshire - Oral Health',
#            'North Sector', 'Corporate Communications', 'Regional Services', 'East Dunbartonshire HSCP'}
# abs_sectors =  ['Clyde Sector', 'Diagnostics Directorate', 'North Sector', 'Regional Services', 'South Sector',
#                "Women & Children's", 'Acute Corporate', 'Board Administration', 'Board Medical Director',
#                'Board Nurse Director', 'Centre For Population Health', 'Corporate Communications', 'eHealth',
#                'Estates and Facilities', 'Finance', 'HR and OD', 'Public Health', 'East Dunbartonshire - Oral Health',
#                'East Dunbartonshire HSCP', 'East Renfrewshire HSCP', 'Glasgow City HSCP', 'Inverclyde HSCP',
#                'Renfrewshire HSCP', 'West Dunbartonshire HSCP', 'Diagnostic Services', 'Acute Directors']

set_of_directorates = set()
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
    min_date = min(relevant_files).strftime('%Y%m%d')
    filename = f'/media/wdrive/Learnpro/HSE Sharps and Skins/{min_date} - HSE Sharps.xlsx'
    return filename


def open_HSE_file(filename):
    """Opens the relevant file and selects relevant portion of staff"""
    df = pd.read_excel(filename, sheet_name='Export')
    df.replace({'Diagnostic Services':'Diagnostics Directorate'}, inplace=True)
    for i in df['Sector/Directorate/HSCP']:
        set_of_directorates.add(i)
    return df


def GGC_compliance(df):
    """Select the relevant in-scope staff, then work out compliance rate among those staff"""
    initial_len = len(df)
    df = df[df['GGC Module'].isin(['Complete', 'Expired', 'Not Undertaken', 'No Account'])]
    scope_len = len(df)
    compliant = df[df['GGC Module'] == 'Complete']
    try:
        compliance_percentage = round(len(compliant) / scope_len * 100, 1)
        return [compliance_percentage, len(compliant), scope_len]
    except ZeroDivisionError:
        return ["No staff eligible", len(compliant), scope_len]


def NES_compliance(df):
    """Same structure as GGC compliance"""
    initial_len = len(df)
    df = df[df['NES Module'].isin(['Complete', 'Expired', 'Not Undertaken', 'No Account'])]
    scope_len = len(df)
    compliant = df[df['NES Module'] == 'Complete']

    try:
        compliance_percentage = round(len(compliant) / scope_len * 100, 1)
        return [compliance_percentage, len(compliant), scope_len]
    except ZeroDivisionError:
        return ["No staff eligible", len(compliant), scope_len]


dates = build_13_month_dates()

# get all files
files = clean_reg_files(os.listdir('/media/wdrive/Learnpro/HSE Sharps and Skins'))


data = pd.DataFrame()
# loop through dates and produce compliance percentages
# for sector in ['Clyde Sector', 'Diagnostics Directorate', 'North Sector', 'Regional Services', 'South Sector',
#                "Women & Children's", 'Acute Corporate', 'Board Administration', 'Board Medical Director',
#                'Board Nurse Director', 'Centre For Population Health', 'Corporate Communications', 'eHealth',
#                'Estates and Facilities', 'Finance', 'HR and OD', 'Public Health', 'East Dunbartonshire - Oral Health',
#                'East Dunbartonshire HSCP', 'East Renfrewshire HSCP', 'Glasgow City HSCP', 'Inverclyde HSCP',
#                'Renfrewshire HSCP', 'West Dunbartonshire HSCP', 'Diagnostic Services', 'Acute Directors']:

abs_13mo = pd.read_excel(
    '/media/wdrive/workforce monthly reports/monthly_reports/May-20 Snapshot/GG&C_Balanced_Scorecard_13m - May-20.xlsx')
depts = abs_13mo['Department'].unique().tolist()

for date in dates:
    filename = find_HSE_file(date)
    df = open_HSE_file(filename)
    for dept in depts:
        curr_df = df[df['department'] == dept]
        ggc = GGC_compliance(curr_df)
        print(ggc)
        nes = NES_compliance(curr_df)
        dataframe_line = {'Department': dept, 'Report Date': pd.to_datetime(date, format='%b-%y'),
                          'Sharps - GGC course %': ggc[0], 'Sharps - GGC inscope': ggc[2],
                          'Sharps - GGC compliant': ggc[1], 'Sharps - NES percentage': nes[0],
                          'Sharps - NES inscope': nes[2], 'Sharps - NES compliant': nes[1]}
        # build single-line dataframe for given date/sector combination
        current_df = pd.DataFrame({k: [v] for k, v in dataframe_line.items()})
        data = data.append(current_df, ignore_index=True)
        print(f'{dept} - {date} - GGC percentage = {ggc[0]}%, GGC Inscope & Compliant = ({ggc[2]}, {ggc[1]}) '
              f'NES percentage = {nes}, NES Inscope & Compliant = ({nes[2]}, {nes[1]})',)

print(set_of_directorates)
today = pd.Timestamp.now().strftime('%Y%m%d')
data.to_excel('/media/wdrive/storyboards/hse-sharps'+today+'.xlsx', index=False)