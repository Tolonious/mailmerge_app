import win32com.client as win32

def send_email(subject, body, to, cc=None, bcc=None):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = body
    mail.To = to
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Send()
