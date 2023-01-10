from main import student_report

std_r = student_report('Marks\Class1_Marks.xlsx')

std_names = std_r.get_names()
rubrics = std_r.get_rubrics()
std1_marks = std_r.get_std_marks('Student 1')
std1_miss = std_r.unsupmitted('Student 1')
std_r.plot_std_marks('Student 1')
std_r.plot_std_rank('Student 1')
