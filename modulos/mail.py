# https://docs.python.org/3.4/library/smtplib.html
# https://docs.python.org/3.4/library/email-examples.html
# https://www.anomaly.net.au/blog/constructing-multipart-mime-messages-for-sending-emails-in-python/
# https://docs.python.org/3/library/email.mime.html

from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from win32com.client import Dispatch
from os.path import basename
from modulos.formPw import FormPw
from abc import ABC, abstractmethod

class Mail(ABC):
    def __init__(self, subject, body, to, cc=None, bcc=None, attach=None, bodyIsHtml=False):
        self.subject = subject
        self.body = body
        self.to = to
        self.cc = cc
        self.bcc = bcc
        self.attach = attach
        self.bodyIsHtml = bodyIsHtml
        self.mail = None
    
    @abstractmethod
    def sendMail(self):
        pass

class MailByOutlook(Mail):
    def __init__(self, subject, body, to, cc=None, bcc=None, attach=None, display=False, bodyIsHtml=None):
        super().__init__(subject, body, to, cc=cc, bcc=bcc, attach=attach, bodyIsHtml=bodyIsHtml)
        self.display = display
        app = Dispatch('Outlook.Application')
        self.mail = app.CreateItem(0)

    def sendMail(self):
        self.mail.Subject = self.subject
        if self.bodyIsHtml:
            self.mail.HTMLBody = self.body
        else:
            self.mail.Body = self.body
        self.mail.To = self.to
        if self.cc: self.mail.CC = self.cc
        if self.bcc: self.mail.BCC = self.bcc
        if self.attach: self.mail.Attachments.Add(self.attach)
        if self.display:
            self.mail.Display()
        else:
            self.mail.Send()

class MailBySmtp(Mail):
    def __init__(self, subject, body, to, cc=None, bcc=None, attach=None, bodyIsHtml=False, sender=None, smtp='smtp.office365.com', port=587):
        super().__init__(subject, body, to, cc=cc, bcc=bcc, attach=attach, bodyIsHtml=bodyIsHtml)
        fPW = FormPw()
        if not (fPW.userId and fPW.userPw):
            self.canceled = True
        else:
            self.canceled = False
            self.mail = SMTP(smtp, port)
            self.mail.starttls()
            self.sender = sender if sender else fPW.userId
            try:
                self.mail.login(fPW.userId, fPW.userPw)
            except:
                raise Exception('Erro ao fazer login no servidor de e-mail. Verifique as credenciais')

    def sendMail(self):
        if not self.canceled:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = self.subject
            if self.bodyIsHtml:
                msg.attach(MIMEText(self.body, 'html'))
            else:
                msg.attach(MIMEText(self.body, 'plain'))
            msg['to'] = self.to
            if self.cc: msg['Cc'] = self.cc
            if self.bcc: msg['Bcc'] = self.bcc
            if self.attach:
                filename = basename(self.attach)
                with open(self.attach, 'rb') as f:
                    attachFile = MIMEApplication(f.read(), name=filename)
                    f.close()
                attachFile.add_header('Content-Disposition', 'attachment', filename=filename)
                msg.attach(attachFile)
            self.mail.sendmail(self.sender, self.to, msg.as_string())
            self.mail.quit()