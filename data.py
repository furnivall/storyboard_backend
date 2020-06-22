"""This file merges the abs file and the """

import pandas as pd

df1 = pd.read_excel('/media/wdrive/storyboards/abs-full.xlsx')
df2 = pd.read_excel('/home/danny/workspace/cms_data.xlsx')

df_merged = df1.merge(df2, on=['Sector/Directorate/HSCP', 'Report Date'], how='left')

df_merged.to_excel('/media/wdrive/storyboards/merged_cms.xlsx', index=False)
