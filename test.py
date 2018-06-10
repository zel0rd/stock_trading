import numpy as np
from pandas import DataFrame
import matplotlib.pyplot as plt



g_dates = [1,2,3,4,5,6,7,8,9,10]
g_closes = [11,22,33,44,55,66,77,88,99,10]

data = [g_dates,g_closes]
newData = np.transpose(data)
df = DataFrame(newData,columns=['dates','closes'])
print(df)