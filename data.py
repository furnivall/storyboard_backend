"""This file merges the abs file and the """

import pandas as pd
today = pd.Timestamp.now().strftime('%Y%m%d')
df1 = pd.read_excel('/media/wdrive/storyboards/abs-full'+today+'.xlsx')
df2 = pd.read_excel('/media/wdrive/storyboards/cms_data'+today+'.xlsx')
df3 = pd.read_excel('/media/wdrive/storyboards/hse-full'+today+'.xlsx')


df_merged = df1.merge(df2, on=['Sector/Directorate/HSCP', 'Report Date'], how='left')
df_merged = df_merged.merge(df3, on=['Sector/Directorate/HSCP', 'Report Date'])
df_merged.to_excel('/media/wdrive/storyboards/merged_file'+today+'.xlsx', index=False)
