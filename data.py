"""This file merges the abs file and the """

import pandas as pd
today = pd.Timestamp.now().strftime('%Y%m%d')
df1 = pd.read_excel('/media/wdrive/storyboards/abs-full'+today+'.xlsx')
df2 = pd.read_excel('/media/wdrive/storyboards/cms_data'+today+'.xlsx')
df3 = pd.read_excel('/media/wdrive/storyboards/hse-sharps'+today+'.xlsx')


df_merged = df1.merge(df2, on=['Department', 'Report Date'], how='left')
#reduce memory usage
df1 = []
df2 = []
print("Merged CMS data")
print("Merging in HSE data")
df_merged = df_merged.merge(df3, on=['Department', 'Report Date'], how='left')
print("Building merged excel")
df_merged.to_csv('/media/wdrive/storyboards/merged_file'+today+'.csv', index=False)
