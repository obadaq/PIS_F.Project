# final project : generating student report from total grades excel sheet

import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


class PDF(FPDF):

    def header(self):
        self.image('ppu_logo.png', 10,8,30)
        self.set_font('times', 'B', 18)
        self.cell(0,10,'Palestine Plytechnic University', border=False, ln=True, align='C')
        self.set_font('times', 'B', 14)
        self.cell(0,6,'Final Student Evaluation', border=False, ln=True, align='C')
        self.ln(20)


class course_exel_analysis:

    
    def __init__(self,marks_df_path):
        
        self.teachers_mail = 'eng.obadaq@gmail.com'
        self.email_pswd = ''
        self.marks_df = pd.read_excel(marks_df_path)
        self.std_names = self.marks_df['Name']
        self.std_emails = self.marks_df['Email'] 
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


    def get_std_emails(self):
        return self.std_emails.drop(0)

    def get_std_marks(self,std_name):
        m_df = self.marks_df
        m_df = m_df.set_index('Name')
        std_marks = dict(m_df.loc[std_name])
        unwated_elm =  ('Email','Rubric')
        for elm in unwated_elm:
            std_marks.pop(elm)
        
        return std_marks

    def get_pswd(self):
        return self.email_pswd

    def set_teachers_mail(self,new_mail):
        self.teachers_mail = new_mail

    def set_pswd(self,new_pswd):
        self.email_pswd = new_pswd

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

    def report(self,std_name):
        bar_chart_path = 'Plots/'+std_name+'_bar.png'
        rank_chart_path = 'Plots/'+std_name+'_Rank.png'
        pdf_name = std_name + '.pdf'
        std_marks = self.get_std_marks(std_name=std_name)
        std_miss = self.unsupmitted(std_name=std_name)
        self.plot_std_marks(std_name=std_name)
        self.plot_std_rank(std_name=std_name)

        report = PDF('P','mm','A4')
        report.set_auto_page_break(auto=True, margin = 15)
        report.add_page()
        report.set_font('times','B', 14)
        report.cell(0,10,"Course Activity Grades _ "+ std_name ,ln=True,border=True)
        report.cell(0,10,"* The Weight of course activities was as follows : ",ln=True,border=False)
        report.ln(100)
        report.image('Plots\Rubric_weight.png', 50,70, 100)
        report.cell(0,10,"* You score the following grades in this course: ",ln=True,border=False)
        report.cell(0,10,str(std_marks),ln=True,border=True,align='C')
        report.image(bar_chart_path, 40,195, 120)
        report.ln(90)
        if std_miss:
            report.cell(0,10,"* You did not submitt the following  : ",ln=True,border=False)
            report.cell(0,10,str(std_miss),ln=True,border=True,align='C')
        report.cell(0,10,"* Your rank in the whole class was as the following : ",ln=True,border=False)
        report.image(rank_chart_path, 40,80, 120)
        report.output(pdf_name)

    def mail_students(self):
        
        smtp_port = 587                 
        smtp_server = "smtp.gmail.com"  


        sender = self.teachers_mail
        recivers_emails = list(self.get_std_emails())
        recivers_names = self.get_names()

        for student in recivers_names:
            self.report(std_name=student)

        self.set_pswd(input(f'Please Enter {sender} Password >>::  '))
        pswd = self.get_pswd()

        email_subject = "Course Evaluation Report"

        i = 0
        for person in recivers_emails:

            body = f"""
            Dear {recivers_names[i]}
        
            Please find your evaluation report in the attachments , for any question you can
            ask me by email or vist me in the office hours.
        
            Regards
        
            Eng.Obada Qawasmi
            Industial Automation labs supervisor
            Palestine Polytechnic University

            """

            # make a MIME object to define parts of the email
            msg = MIMEMultipart()
            msg['From'] = sender
            msg['To'] = person
            msg['Subject'] = email_subject

            # Attach the body of the message
            msg.attach(MIMEText(body, 'plain'))

            # Define the file to attach
            filename = f"{recivers_names}.pdf"

            # Open the file in python as a binary
            attachment= open(filename, 'rb')  # r for read and b for binary

            # Encode as base 64
            attachment_package = MIMEBase('application', 'octet-stream')
            attachment_package.set_payload((attachment).read())
            encoders.encode_base64(attachment_package)
            attachment_package.add_header('Content-Disposition', "attachment; filename= " + filename)
            msg.attach(attachment_package)

            # Cast as string
            text = msg.as_string()

            # Connect with the server
            print("Connecting to server...")
            TIE_server = smtplib.SMTP(smtp_server, smtp_port)
            TIE_server.starttls()
            TIE_server.login(sender, pswd)
            print("Succesfully connected to server")
            print()


            # Send emails to "person" as list is iterated
            print(f"Sending email to: {person}...")
            TIE_server.sendmail(sender, person, text)
            print(f"Email sent to: {person}")
            print()
            i += 1
        TIE_server.quit()