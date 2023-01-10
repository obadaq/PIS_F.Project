# final project : generating student report from total grades excel sheet

import pandas as pd
import matplotlib.pyplot as plt


class student_data:

    
    def __init__(self,marks_df_path):
        

        self.marks_df = pd.read_excel(marks_df_path)
        self.std_names = self.marks_df['Name'] 
        self.rubrics = dict(self.marks_df.iloc[0])
        unwated_elm =  ('Email', 'Name','Rubric', 'Total Grade')
        for elm in unwated_elm:
            self.rubrics.pop(elm)
        plt.pie(self.rubrics.values(), radius=1,labels=self.rubrics.keys(), autopct="%1.2f%%", wedgeprops=dict(width=1, edgecolor='white'))
        plt.title("Weight of Course Activities")
        plt.savefig('Plots/Rubric_weight.png', dpi=300, bbox_inches='tight')
        self.rubrics.update({'Total Grade':100})
    def get_names(self):
        return self.std_names.drop(0)


    def get_rubrics(self):
        return self.rubrics


    def get_std_marks(self,std_name):
        m_df = self.marks_df
        m_df = m_df.set_index('Name')
        std_marks = dict(m_df.loc[std_name])
        unwated_elm =  ('Email','Rubric')
        for elm in unwated_elm:
            std_marks.pop(elm)
        return std_marks


    def unsupmitted(self,std_name):
        m_df = self.marks_df
        m_df = m_df.set_index('Name')
        rubric = dict(m_df.loc[std_name].isnull())
        unwated_elm =  ('Email','Rubric', 'Total Grade')
        for elm in unwated_elm:
            rubric.pop(elm)
        nulls = []
        for item in rubric:
            if rubric.get(item):
                nulls.append(item)
        return nulls
        
    def plot_std_marks(self,std_name):
        
        std_marks = self.get_std_marks(std_name=std_name)
        
        fig, ax = plt.subplots()
        
        ax.bar(range(len(self.rubrics.keys())),list(self.rubrics.values()),label= 'Activities Weight',width=0.3)
        ax.bar(range(len(std_marks.keys())),list(std_marks.values()),label= std_name + ' Marks',width=0.25)

        ax.set_ylabel('Grades')
        ax.set_title(std_name+ ' Activities Marks')
        ax.set_xticks(range(len(self.rubrics.keys())))
        ax.set_xticklabels(list(self.rubrics.keys()),rotation='vertical')
        ax.legend()
        fig.tight_layout()
        fig.savefig('Plots/' + std_name + '_bar.png', dpi=300, bbox_inches='tight')

        
    def plot_std_rank(self,std_name):
        whole_class_marks = self.marks_df[['Name','Total Grade']].sort_values(by='Total Grade').drop(0)
        std_mark = self.get_std_marks(std_name=std_name)['Total Grade']
        std_rank = list(whole_class_marks['Total Grade']).index(std_mark)

        std_mark_list = []
        std_name_list = []
        for i in range(len(list(whole_class_marks['Total Grade']))):
            if i == std_rank :
                std_mark_list.append(std_mark)
                std_name_list.append('You') 
            else: 
                std_mark_list.append(0) 
                std_name_list.append('')
    
        fig, ax = plt.subplots()

        ax.bar(whole_class_marks['Name'],whole_class_marks['Total Grade'],width=0.3)
        ax.bar(whole_class_marks['Name'],std_mark_list,width=0.3)

        ax.set_ylabel('Grades')
        ax.set_title('Whole Class')
        ax.set_xticks(whole_class_marks['Name'])
        ax.set_xticklabels(std_name_list,rotation='vertical')
        fig.tight_layout()
        fig.savefig('Plots/' + std_name + '_Rank.png', dpi=300, bbox_inches='tight')

class as_pdf: