import pandas as pd

excel1 = r"C:\Users\coron\OneDrive\Atom\Python\Workbook1.xlsx"
excel2 = r"C:\Users\coron\OneDrive\Atom\Python\Workbook2.xlsx"

df1 = pd.read_excel(excel1)
df2 = pd.read_excel(excel2)

merge = pd.merge(df1, df2, on="Locations")
print(merge)

merge.to_excel(r"C:\Users\coron\OneDrive\Atom\Python\Output.xlsx")
