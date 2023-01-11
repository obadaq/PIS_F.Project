import smtplib
import ssl

smtp_port = 587
smtp_server = 'smtp.gmail.com'

message = 'Testing SMTP lib For the First Time bro !!'

email_from = 'eng.obadaq@gmail.com'
email_to = 'rnaserdeen@ppu.edu'

pswd = "lpqvazcmpjllcpwu"
simple_email_context = ssl.create_default_context()

try: 
    print("Connecting to server ... ")
    TIE_server = smtplib.SMTP(smtp_server,smtp_port)
    TIE_server.starttls(context=simple_email_context)
    TIE_server.login(email_from, pswd)
    print("Connecting to server ... ")

    print()
    print(f"Sending email to - {email_to}")
    TIE_server.sendmail(email_from, email_to, message)
    print(f"Email successfully sent to - {email_to}")

except Exception as e:
    print(e)

finally:
    TIE_server.quit()