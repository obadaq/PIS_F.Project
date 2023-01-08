# final project : generating student report from total grades excel sheet

import pandas as pd
import matplotlib.pyplot as plt


class student_report:

    
    def __init__(self,marks_df_path):
        
        self.marks_df = pd.read_excel(marks_df_path)
        self.std_names = self.marks_df['Name']


    def get_names(self):
        return self.std_names


    def get_std_marks(self,std_name):
        self.marks_df = self.marks_df.set_index('Name')
        return self.marks_df.loc[std_name]

    def get_nulls(self):
        