import pandas as pd
import plotly.express as px

file1 = 'plot_demo2.csv'
path1 = './999_Data/'

f1 = path1 + file1
df1 = pd.read_csv(f1)

fig1 = px.line(df1, x = df1['Date'], y = df1['Rating'], title = 'Testing 123')
fig1.show()