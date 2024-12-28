import logging
import multiprocessing
import re
import email
import imaplib
import email
import time
from email.header import decode_header

from connect_client import send_data


class AutoGet:

    def __init__(self):
        self.__imap_host = 'imap.qq.com'  # 邮箱服务器地址
        self.__imap_user = 'your_email_address' # 邮箱账号
        self.__imap_pass = 'fpzduoqjczvvbhbc' # 邮箱密码
    # 邮箱服务器地址，参数有：邮箱服务器地址，邮箱账号，邮箱密码
    def connect_to_email_server(self):
        """
        连接到 IMAP 邮件服务器并登录。
        """
        mail = imaplib.IMAP4_SSL(self.__imap_host) # 连接到 IMAP 服务器
        mail.login(self.__imap_user, self.__imap_pass) # 登录邮箱
        return mail
        #返回已登录的mail对象，可以进行邮件操作


    # 解码邮件内容
    def decode_payload(self, payload, charset):
        """
        尝试使用不同的编码来解码邮件内容。
        """
        try:
            return payload.decode(charset)
        except UnicodeDecodeError:
            try:
                return payload.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    return payload.decode('gbk')
                except UnicodeDecodeError:
                    return payload.decode('latin1')

    # 获取指定数量的最新邮件，参数包含：邮箱连接，邮件数量
    def fetch_email(self, mail, email_id):
        """
        根据邮件ID获取邮件内容。
        """
        status, data = mail.fetch(email_id, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])

        # 解码邮件主题
        subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or 'utf-8', errors='replace')

        # 提取纯文本内容
        text_content = ""
        if msg.is_multipart():
            for part in msg.walk():  # 遍历邮件的各个部分
                if part.get_content_type() == 'text/plain':
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset()
                    text_content = self.decode_payload(payload, charset)
                    break
        else:
            if msg.get_content_type() == 'text/plain':
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset()
                text_content = self.decode_payload(payload, charset)

        return subject, text_content


    # 获取所有邮件的ID，参数为：邮箱连接
    def get_all_email_ids(self, mail):
        """
        获取邮箱中所有邮件的ID。
        """
        mail.select('inbox')
        status, messages = mail.search(None, 'ALL')
        return messages[0].split()


    # 监控邮箱，参数包含：邮箱服务器地址，邮箱账号，邮箱密码，监控间隔
    def monitor_inbox(self, interval=30):
        """
        监控邮箱，新邮件到达时打印邮件内容。
        """
        print("开始监控邮箱...")
        mail = self.connect_to_email_server()
        last_checked_ids = set(self.get_all_email_ids(mail))

        try:
            while True:
                current_ids = set(self.get_all_email_ids(mail))
                new_ids = current_ids - last_checked_ids
                if new_ids:
                    for email_id in new_ids:
                        subject, text_content = self.fetch_email(mail, email_id)
                        print("新邮件到达！")
                        print("Subject: ", subject)
                        print("Content:", text_content)
                        last_checked_ids = current_ids
                        yzm = self.extract_verification_code(text_content)
                        send_data(yzm)
                        time.sleep(interval)
        finally:
            mail.logout() # 退出邮箱


    # 以下为修改后的代码，增加了验证码提取功能

    def extract_verification_code(self, email_content):
        """
        从邮件内容中提取验证码，假设验证码为6位数字。
        """
        # 使用正则表达式匹配6位数字验证码
        match = re.search(r'\b\d{6}\b', email_content)
        if match:
            return match.group()
        return None



if __name__ == '__main__':
    port_list = [7171]
    process = [multiprocessing.Process(target=AutoGet().monitor_inbox, args=(i,)) for i in
               port_list]
    for p in process:
        p.start()
    for p in process:
        p.join()
