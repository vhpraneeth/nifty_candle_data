#!/usr/bin/env python3
import pandas as pd

print('Process started')
df = pd.read_excel('NIFTY25JUN2010000PE.xlsx', parse_dates= [ ['Date', 'Time']])

df.drop(['Ticker'], axis=1, inplace=True)
# df['timestamp'] = pd.to_datetime(df['Date_Time'])
df.set_index(['Date_Time'], inplace=True)
df = df.resample('15Min').agg({'Open ': 'first', 'High ': 'max', 'Low ': 'min', 'Close ': 'last'})

df.to_excel('Result.xlsx', index=False)
df.to_csv('Result.csv', index=False)
