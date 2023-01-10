from main import student_data
from fpdf import FPDF
from math import isnan

std_r = student_data('Marks\Class1_Marks.xlsx')

std_names = std_r.get_names()
rubrics = std_r.get_rubrics()
student_Name = 'Student 1'
std1_marks = std_r.get_std_marks(student_Name)


std1_miss = std_r.unsupmitted(student_Name)
std_r.plot_std_marks(student_Name)
std_r.plot_std_rank(student_Name)


std_r.report('Student 5')