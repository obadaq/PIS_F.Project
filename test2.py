from main import CourseExcelAnalysis

import smtplib
import ssl
 
std_r = CourseExcelAnalysis('Marks\\Class2_Marks.xlsx')

#std_r.report("Rafat")
std_r.mail_students()