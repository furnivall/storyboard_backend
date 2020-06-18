import os
import pandas as pd

def build_13_month_dates():
    month = pd.Timestamp.now()
    print(month)

build_13_month_dates()