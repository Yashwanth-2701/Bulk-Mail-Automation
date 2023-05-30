from tkinter import *
from tkinter import filedialog

import pandas as pd
import smtplib
from tkinter import messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys,os

application_path = os.path.dirname(sys.executable)

mailid=""
password=""

# Code
def send_attachment():
    
    filename_info = filename.get()
    sem_info = sem.get()
    date1_info = date1.get()
    date2_info = date2.get()
    internals_info = internals.get()

    global email
    e = pd.read_excel("mail.xlsx")
    emails = e['Mails'].values
    print('mails sent to:')
    print(emails)

    msg = MIMEMultipart()

    msg['Subject'] = "Regarding Internal Assessment"
    msg['From'] = mailid
    email_body_info = '''Greetings All,

    As ''' + internals_info + ''' Internal Assessment for ''' + sem_info + ''' Semester Students is Scheduled on ''' + date1_info + ''', All Faculty Members who are Handling Subjects for this Class are Hereby informed to Prepare 2 Sets of Question Paper and Send it to HOD sir on or before ''' + date2_info + '''

    Note:
    1. Please use the question paper template provided in the format. 
    2. Send the question papers in pdf format only.
    3. Make sure all the questions are in the font style preferably TimesNew Roman 


    Thank you,
    Regards '''

    msg.attach(MIMEText(email_body_info, 'plain'))

    filename_info = filedialog.askopenfilename(initialdir="D:", title='open file')
    attachment = open(filename_info, 'rb')

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    p.set_payload(attachment.read())

    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename_info)

    msg.attach(p)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(mailid, password)
    text = msg.as_string()

    print("Login successful")

    for email in emails:
        server.sendmail(mailid, email, text)

    print("message sent")

    messagebox.showinfo('Success', 'Mail Sent Successfully')


def send_message():
    sem_info = sem.get()
    date1_info = date1.get()
    date2_info = date2.get()
    internals_info = internals.get()

    e = pd.read_excel("mail.xlsx")
    emails = e['Mails'].values
    print('mails sent to:')
    print(emails)

    # Server Establishment

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(mailid, password)

    # Mail Contents
    subject = "Regarding Internal Assessment"
    msg = '''Greetings All,

    As ''' + internals_info + ''' Internal Assessment for ''' + sem_info + ''' Semester Students is Scheduled on ''' + date1_info + ''', All Faculty Members who are Handling Subjects for this Class are Hereby informed to Prepare 2 Sets of Question Paper and Send it to HOD sir on or before ''' + date2_info + '''

    Note:
    1. Please use the question paper template provided in the format. 
    2. Send the question papers in pdf format only.
    3. Make sure all the questions are in the font style preferably TimesNew Roman 


    Thank you,
    Regards '''

    body = "Subject:{}\n\n{}".format(subject, msg)

    for mail in emails:
        server.sendmail(mailid, mail, body)

        print("message sent")

    messagebox.showinfo('Success', 'Mail Sent Successfully')

# GUI
app = Tk()

app.geometry("500x250")
app.title("MAP")

heading = Label(text="MAP", bg="black", fg="white", font="10", width="500", height="2")
heading.pack()

internals = Label(text="Internals")
internals.place(x=85, y=100)

sem = Label(text="Sem")
sem.place(x=195, y=100)

date1 = Label(text="Internals Date")
date1.place(x=275, y=100)

date2 = Label(text="Last Date")
date2.place(x=380, y=100)

filename = StringVar()
internals = StringVar()
sem = StringVar()
date1 = StringVar()
date2 = StringVar()


internals_entry = Entry(textvariable=internals, width=10)
internals_entry.place(x=75, y=130)

sem_entry = Entry(textvariable=sem, width=10)
sem_entry.place(x=175, y=130)

date1_entry = Entry(textvariable=date1, width=10)
date1_entry.place(x=275, y=130)

date2_entry = Entry(textvariable=date2, width=10)
date2_entry.place(x=375, y=130)

button = Button(app, text="Send Message", command=send_message, width="20", height="2", bg="grey")
button.place(x=70, y=180)

button = Button(app, text="Send Attachment", command=send_attachment, width="20", height="2", bg="grey")
button.place(x=275, y=180)

mainloop()
