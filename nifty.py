#!/usr/bin/env python3
import pandas as pd

print('Process started')
df = pd.read_excel('NIFTY25JUN2010000PE.xlsx', parse_dates=[['Date', 'Time']])

df.set_index(['Date_Time'], inplace=True)
df = df.resample('15Min').agg({'Open ': 'first', 'High ': 'max', 'Low ': 'min', 'Close ': 'last'})

df['Ticker'] = 'NIFTY25JUN2010000PE'
df.reset_index(inplace=True)
df['Date'] = df['Date_Time'].dt.date
df['Time'] = df['Date_Time'].dt.time

df['Date'] = df['Date'].apply(lambda x: x.strftime('%d/%m/%Y'))
df['Time'] = df['Time'].apply(lambda x: x.strftime('%H:%M'))

df = pd.DataFrame(df, columns=['Ticker', 'Date', 'Time', 'Open ', 'High ', 'Low ', 'Close ', ])  # 'Volume'
df.to_excel('Result.xlsx', index=False)
df.to_csv('Result.csv', index=False)
