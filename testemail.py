

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from main import student_data

std_r = student_data('Marks\Class1_Marks.xlsx')
std_emails = list(std_r.get_std_emails())

# Setup port number and server name

smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp.gmail.com"  # Google SMTP Server

# Set up the email lists
email_from = "eng.obadaq@gmail.com"
std_names = list(std_r.get_names())
email_list = std_emails

# Define the password (better to reference externally)
pswd = "lpqvazcmpjllcpwu" # As shown in the video this password is now dead, left in as example only


# name the email subject
subject = "Course Evaluation Report"



# Define the email function (dont call it email!)
def send_emails(email_list):
    i = 0
    for person in email_list:

        # Make the body of the email
        body = f"""
        Dear {std_names[i]}
        
        Please find your evaluation report in the attachments , for any question you can
        ask me by email or vist me in the office hours.
        
        Regards
        
        Eng.Obada Qawasmi
        Industial Automation labs supervisor
        Palestine Polytechnic University

        """

        # make a MIME object to define parts of the email
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = person
        msg['Subject'] = subject

        # Attach the body of the message
        msg.attach(MIMEText(body, 'plain'))

        # Define the file to attach
        filename = "Student 5.pdf"

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
        TIE_server.login(email_from, pswd)
        print("Succesfully connected to server")
        print()


        # Send emails to "person" as list is iterated
        print(f"Sending email to: {person}...")
        TIE_server.sendmail(email_from, person, text)
        print(f"Email sent to: {person}")
        print()
        i += 1
    # Close the port
    TIE_server.quit()


# Run the function
send_emails(email_list)
