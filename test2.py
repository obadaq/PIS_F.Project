from main import CourseExcelAnalysis

import smtplib
import ssl
 
std_r = CourseExcelAnalysis('Marks\\Class1_Marks.xlsx')

print(std_r.report("Student 1"))
