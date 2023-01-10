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
plt.savefig('Plots/Rubric_weight.png', dpi=300, bbox_inches='tight')
#plt.show()


m_df = marks_df
m_df = m_df.set_index('Name')
std_marks = dict(m_df.loc['Student 1'])
unwated_elm =  ('Email','Rubric', "Total Grade")
for elm in unwated_elm:
    std_marks.pop(elm)

print(std_marks)

rubrics = marks_df.iloc[0]
unwated_elm =  ('Email', 'Name','Rubric', 'Total Grade')
for elm in unwated_elm:
    rubrics.pop(elm)

print(dict(rubrics))
rubrics=dict(rubrics)

fig, ax = plt.subplots()
ax.bar(range(len(rubrics.keys())),list(rubrics.values()),label= 'Activities Weight',width=0.3)
ax.bar(range(len(std_marks.keys())),list(std_marks.values()),label= 'Student 1' + ' Marks',width=0.25)

ax.set_ylabel('Grades')
ax.set_title('Student 1'+ ' Activities Marks')
ax.set_xticks(range(len(rubrics.keys())))
ax.set_xticklabels(list(rubrics.keys()),rotation='vertical')
ax.legend()
fig.tight_layout()
plt.show()
fig.savefig('Plots/' + 'Student 1' + '_bar.png', dpi=300, bbox_inches='tight')

'''

whole_class = marks_df[['Name','Total Grade']].sort_values(by='Total Grade').drop(0)
whole_class1 = list(whole_class['Total Grade']).index(95)
new = []
new1 = []
for i in range(len(list(whole_class['Total Grade']))):
    if i == whole_class1 :
        new.append(95)
        new1.append('You') 
    else: 
        new.append(0) 
        new1.append('')
print(new)

fig, ax = plt.subplots()
ax.bar(whole_class['Name'],whole_class['Total Grade'],width=0.3)
ax.bar(whole_class['Name'],new,width=0.3)

ax.set_ylabel('Grades')
ax.set_title('Whole Class')
ax.set_xticks(whole_class['Name'])
ax.set_xticklabels(new1,rotation='vertical')
fig.tight_layout()
plt.show()
'''


ax.bar(range(len(std_marks.keys())),list(std_marks.values()),label= 'Student 1' + ' Marks',width=0.25)

ax.set_ylabel('Grades')
ax.set_title('Student 1'+ ' Activities Marks')
ax.set_xticks(range(len(rubrics.keys())))
ax.set_xticklabels(list(rubrics.keys()),rotation='vertical')
ax.legend()
fig.tight_layout()
'''