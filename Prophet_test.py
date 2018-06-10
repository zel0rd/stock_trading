# -*- coding: utf-8 -*-
"""
Created on Mon Jun 11 04:45:31 2018

@author: zelord

new_dates
g_closes


d = {'ds':new_dates,'y':g_closes}

df = pd.DataFrame(data=d)

df

"""

import pandas as pd
from fbprophet import Prophet
df = pd.read_csv('example_wp_log_peyton_manning.csv')
df.head()

m = Prophet()
m.fit(df)

future = m.make_future_dataframe(periods=365)
future.tail()

forecast = m.predict(future)
forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail()

 
fig1 = m.plot(forecast)
fig1
fig2 = m.plot_components(forecast)
fig2
 
