import pandas as pd
import matplotlib.pyplot as plt


marks_df = pd.read_excel("Marks\Class1_Marks.xlsx")
stdnames = marks_df['Name']
marks_df = marks_df.set_index('Name')

##print(marks_df.iloc[1].isnull()['HW2'] )
print(stdnames)