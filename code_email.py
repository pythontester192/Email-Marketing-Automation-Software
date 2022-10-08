import smtplib  # to deal with gmail mails


def Email_send_function(to, subject, message, uname, pasw):
    # print(to,subject,message,uname,pasw)
    s = smtplib.SMTP("smtp.gmail.com", 587)  # create session for gmail
    s.starttls()  # transport layer
    s.login(uname, pasw)
    msg = "Subject: {}\n\n{}".format(subject, message)
    s.sendmail(uname, to, msg)
    x = s.ehlo()
    if x[0] == 250:
        return "s"
    else:
        return "f"
    s.close()
