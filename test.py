import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

marks_df = pd.read_excel("Marks\Class1_Marks.xlsx")
stdnames = marks_df['Name']
'''
marks_df = marks_df.set_index('Name')
dict1 = dict(marks_df.iloc[4].isnull())
del_dict1 = ('Email','Rubric', 'Total Grade')

for d in del_dict1:
    dict1.pop(d)

print(dict1 )



nulls = []
for item in dict1:
    if dict1.get(item):
        nulls.append(item)

print(nulls)

'''
ru = marks_df.columns.values.tolist()
del ru[0:3]
ru.remove("Total Grade")


rubrics = marks_df.iloc[0]
unwated_elm =  ('Email', 'Name','Rubric', 'Total Grade')
for elm in unwated_elm:
    rubrics.pop(elm)
print(rubrics)

plt.pie(rubrics, radius=1,labels=rubrics.keys(), autopct="%1.2f%%", wedgeprops=dict(width=1, edgecolor='white'))
plt.title("Weight of Course Activities")
plt.savefig('plot.png', dpi=300, bbox_inches='tight')
plt.show()