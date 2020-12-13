import pandas as pd
import numpy as np

df = pd.read_excel(r'C:\Users\coron\OneDrive\Atom\Python\For_Python.xlsx')
print(df.columns[1])
print(df.columns[:])
webinar_brc = df["BRC"][0]
print(webinar_brc)
