"""
This is a code to read excel file and calculate sum (Matrix) values
"""
import pandas as pd
import PySimpleGUI as sg

path = sg.popup_get_file("Please Select The File")
df = pd.read_excel(path, index_col=0)

#Total sum per column:
df.loc["Total", :] = df.sum(axis=0)

#Total sum per row:
df.loc[:, "Sum"] = df.sum(axis=1)

print(df)

# saving the dataframe
df.to_excel(path_to_save_file)
