# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 18:01:47 2020

@author: Alex
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 16:52:25 2020

@author: Alex
"""

import requests as req
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns
import configparser

config = configparser.ConfigParser()
config.read("config.ini")

api_key = config['AEMET-opendata']['api_key']
print(api_key)


start_date, end_date = '2016-01-01T00:00:00UTC', '2020-01-02T00:00:00UTC'
station = '8019'

url = 'https://opendata.aemet.es/opendata/api/valores/climatologicos/diarios/datos/fechaini/'+start_date+'/fechafin/'+end_date+'/estacion/'+station

querystring = {"api_key":api_key}

headers = {
    'cache-control': "no-cache"
    }

response = req.request("GET", url, headers=headers, params=querystring)

print(response.text)

json_response =response.json()

# Pull json data from url response as pandas dataframe
df = pd.read_json (json_response['datos'])

print(df.dtypes)

# Select the columns we want to work with
cols = ['tmed', 'tmin', 'tmax','fecha']
df = df[cols]


# Convert values in order
df['fecha'] = pd.to_datetime(df['fecha'])
df['tmed'] = df['tmed'].str.replace(',', '.').astype(float)
df['tmin'] = df['tmin'].str.replace(',', '.').astype(float)
df['tmax'] = df['tmax'].str.replace(',', '.').astype(float)

# Generate new column with Max - Min temperature
df['tdif'] = df['tmax'] - df['tmin']

# Check out the data types of each column.
print(df.dtypes)
print(df.head())

# Set date as the index of the dataframe
aemet_daily = df.set_index('fecha')

print(aemet_daily.head())

# Add columns with year, month, and weekday name
aemet_daily['Year'] = aemet_daily.index.year
aemet_daily['Month'] = aemet_daily.index.month
aemet_daily['Weekday Name'] = aemet_daily.index.day_name()


# Display a random sampling of 5 rows
print(aemet_daily.sample(5, random_state=0))

# Use seaborn style defaults and set the default figure size
sns.set(rc={'figure.figsize':(11, 4)})
aemet_daily['tmed'].plot(marker='.', alpha=0.5, linestyle='None',linewidth=0.5);

cols_plot = ['tmed', 'tmin', 'tmax']
    
 
fig, axes = plt.subplots(len(cols_plot), 1, figsize=(11, 10), sharex=True)

for name, ax in zip(['tmed', 'tmin', 'tmax'], axes):    
    sns.boxplot(data=aemet_daily, x='Month', y=name, ax=ax)
    ax.set_ylabel('Temperature (ºC)')    
    ax.set_title(name)
# Remove the automatic x-axis label from all but the bottom subplot
if ax != axes[-1]:
    ax.set_xlabel('')    
    
fig, ax = plt.subplots()    
sns.boxplot(data=aemet_daily, x='Month', y='tdif', ax=ax)
ax.set_ylabel('Difference (Max-Min)')
ax.set_title('tdif')
    
    
# Specify the data columns we want to include
data_columns = ['tmed', 'tmin', 'tmax','tdif']
# Resample to weekly frequency, aggregating with mean
aemet_weekly_mean = aemet_daily[data_columns].resample('W').mean()
aemet_weekly_mean.head(5)

# Start and end of the date range to extract
start, end = '2019-01', '2019-6'
# Plot daily and weekly resampled time series together
fig, ax = plt.subplots()
ax.plot(aemet_daily.loc[start:end, 'tmed'],
marker='.', linestyle='-', linewidth=0.5, label='Daily')
ax.plot(aemet_weekly_mean.loc[start:end, 'tmed'],
marker='o', markersize=8, linestyle='-', label='Weekly Mean Resample')
ax.set_ylabel('Temperature (ºC)')
ax.legend();

# Compute the monthly means
aemet_monthly_mean = aemet_daily[data_columns].resample('M').mean()
aemet_monthly_mean.head(5)

fig, ax = plt.subplots()
ax.plot(aemet_monthly_mean['tmed'], color='black', label='tmed')
aemet_monthly_mean[['tmin', 'tmax']].plot.area(ax=ax, linewidth=0)
ax.xaxis.set_major_locator(mdates.YearLocator())
ax.legend()
ax.set_ylabel('Temperature (ºC)');

