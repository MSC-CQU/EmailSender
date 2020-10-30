from docx import Document
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import time
import smtplib


class SendEmail(object):
    def __init__(self):
        self.to_addrs = ["ym13883407586@outlook.com"]  # receivers_address
        self.sys_date = time.strftime('%Y-%m-%d', time.localtime())
        self.from_addr = 'cqu@microsoftstudent.club'  # sender_address
        self.password = 'xxxxxx'  # sender_password
        self.smtp_server = 'smtp.mxhichina.com'  # sender_smtpserver

    def _format_addr(self, s):
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))

    def send_email(self, title, path):
        document = Document(path)
        content = ''
        for paragraph in document.paragraphs:
            content = content + '\n' + paragraph.text
        msg = MIMEText(content, 'plain', 'utf-8')
        msg['From'] = self._format_addr(u'MSC_CQU<%s>' % self.from_addr)  # sender_name
        msg['Subject'] = Header(title, 'utf-8').encode()

        server = smtplib.SMTP(self.smtp_server, 25)
        server.login(self.from_addr, self.password)
        for receiver in self.to_addrs:
            msg['To'] = self._format_addr(u'MSC_Member<%s>' % receiver)  # receiver_name
            server.sendmail(self.from_addr, [receiver], msg.as_string())
            time.sleep(1)
        server.quit()


sender = SendEmail()
sender.send_email(u'测试邮件', r"D:\EmailSender\test.docx")  # word file position
print('Done.')
