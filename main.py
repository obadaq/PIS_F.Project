# final project : generating student report from total grades excel sheet

import pandas as pd
import matplotlib.pyplot as plt


class student_report:

    
    def __init__(self,marks_df_path):
        

        self.marks_df = pd.read_excel(marks_df_path)
        self.std_names = self.marks_df['Name'] 
        self.rubrics = self.marks_df.iloc[0]
        unwated_elm =  ('Email', 'Name','Rubric', 'Total Grade')
        for elm in unwated_elm:
            self.rubrics.pop(elm)
        
       
    def get_names(self):
        return self.std_names

    def fet_rubrics(self):
        return self.rubrics

    def get_std_marks(self,std_name):
        m_df = self.marks_df
        m_df = m_df.set_index('Name')
        std_marks = dict(m_df.loc[std_name])
        unwated_elm =  ('Email','Rubric')
        for elm in m_df:
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
        
