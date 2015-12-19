"""
Handles sending of email reports.
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
import time
import easygui


# def Confirm(cnxn, to, msg):
#     if easygui.ynbox('Send message to {}'.format(to)):
#         mail_cnxn.sendmail('stoebelj@berea.edu', [to, 'stoebelj@berea.edu'], msg.as_string())
#     print '-> sent!'

def TestCnxnOpen(cnxn):
    """
    tests that a connection is open.
    :param cnxn: a connection to an email server
    :return: if the connection is open
    """
    try:
        status = cnxn.noop()[0]
    except smtplib.SMTPServerDisconnected:  # smtplib.SMTPServerDisconnected
        status = -1
    return True if status == 250 else False

def EmailConnect(password):
    """
    creates an email connection to your mail server
    :param password: email password
    :return: mail connection object
    """

    mail_cnxn=smtplib.SMTP('mail.somwhere.com',25)
    mail_cnxn.starttls()
    mail_cnxn.ehlo()
    mail_cnxn.login('username', password)
    return mail_cnxn

def EmailAdvisor(student, go, password):
    """
    emails an advisor about any no passing scores.
    :param student: a student object
    :param go: if this email should send
    :param password: email password
    :return: email sent to advisor unless go is False
    """

    mail_cnxn = EmailConnect(password)

    tests=''
    #attach tests
    for test in student.tests:
        tests+=test.text
        tests+='<p>'

    with open('advisor_email.txt', 'r') as text:
        email_body=text.read().format(advisor_sal=student.advisor_salutation, stu_first=student.stu_first, stu_last=student.stu_last, tests=tests)

    html_text=MIMEText(email_body, 'html')      #prepares the email body to go into the message
    msg=MIMEMultipart('alternative')
    msg.attach(html_text)
    sender = formataddr((str(Header(u'Your Name', 'utf-8')), "someone@somewhere.com"))     #use this to insert what ever name you like into the from field.
    msg['Subject'] = 'Praxis Results for {}'.format(student.name)
    msg['From'] = sender
    msg['To'] = student.advisor_email
    msg['bcc']='someone@somewhere.com'

    print 'Sending results to {0} re: {1}...'.format(student.advisor_email, student.name),

    if go:

        while True:
            if not TestCnxnOpen(mail_cnxn):     #create a new connection if the current one is closed.
                mail_cnxn = EmailConnect(password)

            try:
                mail_cnxn.sendmail('someone@somewhere.com', [student.advisor_email, 'someone@somewhere.com'], msg.as_string())
                print '-> Sent!'
                break
                #wait 60 seconds if disconnected
            except smtplib.SMTPServerDisconnected:
                print '-> Disconnected. Waiting 60 seconds.'
                time.sleep(60)

    else:
        print 'Send not enabled, nothing sent.'

def EmailStudent(student, go, password):
    """
    emails a student about any no passing scores.
    :param student: a student object
    :param go: if this email should send
    :param password: email password
    :return: email sent to student unless go is False
    """

    tests=''
    for test in student.tests:
        if not test.passing:        #failing tests only
            tests+='<li>{test_name}'.format(test_name=test.name)

    mail_cnxn = EmailConnect(password)

    #add the end of the email
    with open('student_email.txt', 'r') as text:
        email_body=text.read().format(stu_first=student.pref_first, tests=tests)

    html_text=MIMEText(email_body, 'html')      #prepares the email body to go into the message
    msg=MIMEMultipart('alternative')
    msg.attach(html_text)
    sender = formataddr((str(Header(u'Jacob Stoebel', 'utf-8')), "stoebelj@berea.edu"))     #use this to insert what ever name you like into the from field.

    msg['Subject'] = 'Praxis results'
    msg['From'] = sender
    msg['To'] = student.stu_email
    msg['bcc']='stoebelj@berea.edu'

    print 'Sending results to {0} re: {1}...'.format(student.stu_email, 'retest'),
    if go:

        while True:
            if not TestCnxnOpen(mail_cnxn):     #create a new connection if the current one is closed.
                mail_cnxn = EmailConnect(password)

            try:
                mail_cnxn.sendmail('stoebelj@berea.edu', [student.stu_email, 'stoebelj@berea.edu'], msg.as_string())
                print '-> Sent!'
                break
                #wait 60 seconds if disconnected
            except smtplib.SMTPServerDisconnected:
                print '-> Disconnected. Waiting 60 seconds.'
                time.sleep(60)

    else:
        print 'Send not enabled, nothing sent.'

def EmailChair(students, go, password):
    """
    emails dept chair about any no passing scores.
    :param student: a student object
    :param go: if this email should send
    :param password: email password
    :return: email sent to dept chair unless go is False
    """

    tests=''
    for student in students:
        tests+='''<li>{stu_name}
        <ul>'''.format(stu_name=student.name)
        for test in student.tests:
            tests+=test.text
            tests+='<p>'
        tests+='</ul>'

    mail_cnxn = EmailConnect(password)

    with open('yoli_email.txt', 'r') as data:
        email_body=data.read().format(tests=tests)

    html_text=MIMEText(email_body, 'html')      #prepares the email body to go into the message
    msg=MIMEMultipart('alternative')
    msg.attach(html_text)
    sender = formataddr((str(Header(u'Jacob Stoebel', 'utf-8')), "stoebelj@berea.edu"))     #use this to insert what ever name you like into the from field.
    msg['Subject'] = 'Praxis Results'
    msg['From'] = sender
    msg['To'] = 'Cartery@berea.edu'
    msg['bcc']='stoebelj@berea.edu'

    print 'Sending results to Yoli...',

    if go:

        while True:
            if not TestCnxnOpen(mail_cnxn):     #create a new connection if the current one is closed.
                mail_cnxn = EmailConnect(password)

            try:
                mail_cnxn.sendmail('stoebelj@berea.edu', ['Cartery@berea.edu', 'stoebelj@berea.edu'], msg.as_string())
                print '-> Sent!'
                break
                #wait 60 seconds if disconnected
            except smtplib.SMTPServerDisconnected:
                print '-> Disconnected. Waiting 60 seconds.'
                time.sleep(60)

    else:
        print 'Send not enabled, nothing sent.'

def EmailHandler(all_students, cursor, password, go=False):
    '''Accepts a student object:

        advisors are informed when a student does not pass.
        Yoli recieves all non passing information in one email
        Student is emailed when any test comes back as not passing.

        go must be set to True or emails won't send!
        '''

    for_yoli=[]
    for student in all_students:

        student.compute_pass_status()

        if not student.all_pass:

            student.fetch_email_info(cursor)        #finds relevant info on student and advisor to send emails.


            for test in student.tests:              #build HTML on tests only if an email is needed.
                test.build_html()

            if student.advisor_email:
                EmailAdvisor(student, go, password)
            else:
                'NOTE: {0} does not have an EDS advisor. No advisor was emailed'.format(student.name)

            if student.enrollment_status!='Graduation':     #only current students get an email or are refered to Yoli
                EmailStudent(student, go, password)
                for_yoli.append(student)

    if for_yoli:
        EmailChair(for_yoli, go, password)