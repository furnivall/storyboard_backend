import pandas as pd

df = pd.read_excel('/media/wdrive/workforce monthly reports/monthly_reports/May-20 Snapshot/GG&C_Balanced_Scorecard_13m - May-20.xlsx')
print(df.shape)
print(len(df['Department'].unique()))