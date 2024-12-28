import re
import time

import requests
import imaplib
import email
import sys
from requests.compat import chardet  # 用于自动检测文件编码

'''
邮箱帐号、密码、刷新令牌
ulachnaiomy@hotmail.com----rYxUlFxIUGmZu----M.C504_BAY.0.U.-Cpq1sUGkrZkE3TN8nWZkStu0rFJcMNIsvp7Jt2N5Xe2TJSFaMXcop8HzG7JT5SNr6OxF0mPnqW2C50lF!U*M1I0KQjBhhGvC2UE2DtyS7U7L2711kyriWlKnMfpHxzACP66H3Rxoc9AB1MfILudRIaaFF9JFsn0yjKJV9k8NA3AV0nOsCbepvaVJamCqsfwAi79BJ!efswsBM6LdL!f7Pr!x7glR*IG*Q!7!LoQktd9U4nTDp*V81uoZ8zPTzreFiBfH9K61B0ycVQOqk2YmXFVsQ0K5NZlry4TgXjRb8ecQaP9nvvpZDLxjjCUiLTegMVz2j*zgaJdBPfFXvLHxBGwerbh9VgFudcwaIJWOE!6RE4LsYV5lYLlfj43ABs4HHG9IoEhdBDG2cjCJwDru56M$
De4Ur4Qd8Kf8Cn3@outlook.com----Ty3Rp2Cc5Qs1----M.C536_BAY.0.U.-Cmc4Kn*V2kw8Y5RRsXFCWNTbNf7!RZ5nCjqrJzClJV054gLEiKOimSFkmGqqAUJVYb6nPvsF3BmFky2M3Z6ZdQJkCsK6Qzf6k3pHym8ffpHOd!eaLZH6K2LmTLs6*BKFlx9yAY96biIPsmXklxWF5hc4h*PzykJPL*dCma3ybSfLZXHXUINDPb04FbCWUUFJBbVO4FfISli1Wz*U3SnkVQvs199bvLZsvWIe9IphszhIK7L3y9nArdD*RGCFhnCXWWxwYQvYKT1KRVewAmBswuyQlY5WTvkZLnPxV*mVwRpfBjnDu03JWLiGBcPaRzfMHDaiMu5LdhnrduE6FKw4U4jfv*sDh7pHSVfAz2NNRMZ9IPVqAbP8uMPdXBKpUiYFCuyk3M0loW9!MfVHbprUHBI$
'''


def fetch_without_proxy(url, data=None, headers=None):
    try:
        session = requests.Session()
        proxies = {
            'http': None,
            'https': None
        }
        response = session.get(url, proxies=proxies, data=data, headers=headers,timeout=10)
        return response
    except requests.exceptions.RequestException as e:
        print(f"获取网页内容失败: {e}")
        return None


# 1. 获取 access_token
def get_accesstoken(refresh_token):
    data = {
        'client_id': '9e5f94bc-e8a4-4e73-b8be-63364c29d753',
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}

    try:
        # 发送 POST 请求以获取 Access Token  request设置成无代理
        proxies = {
            'http': None,
            'https': None
        }
        ret = fetch_without_proxy('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',data=data, headers=headers)
        # ret = requests.get('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',data=data, headers=headers)
        # ret = requests.post('https://login.microsoftonline.com/consumers/oauth2/v2.0/token',data=data, headers=headers)

        ret.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"获取 Access Token 失败: {e}")
        return None

    response = ret.json()
    if 'access_token' not in response:
        print("未能成功获取 Access Token")
        return None
    return response['access_token']


# 2. 生成认证字符串
def generate_auth_string(user, token):
    auth_string = f"user={user}\x01auth=Bearer {token}\x01\x01"
    return auth_string


# 3. 使用 IMAP 访问 Outlook 邮箱
def connect_imap(user_email, access_token):
    try:
        mail = imaplib.IMAP4_SSL('outlook.office365.com')

        mail.authenticate('XOAUTH2', lambda x: generate_auth_string(user_email,access_token).encode('utf-8'))
        return mail
    except Exception as e:
        print(f"IMAP 连接失败: {e}")
        return None


# 4. 从文本文件或手动输入获取账号信息
def get_account_info(choice):
    # choice = input("输入 1 从文本文件获取邮箱信息，输入 2 手动输入账户信息: ").strip()
    if choice == '1':
        file_path = input(
            r"请输入包含邮箱帐号文本路径（例如：C:\Users\admin\Desktop\accounts.txt）: ").strip()
        accounts = []
        try:
            with open(file_path, 'rb') as file:
                raw_data = file.read()
                result = chardet.detect(raw_data)
                encoding = result['encoding']
                confidence = result['confidence']
                print(
                    f"检测到文件编码: {encoding}（置信度: {confidence * 100:.2f}%）")
                if not encoding:
                    print(
                        "无法检测文件编码，请手动指定编码或确保文件使用标准编码格式。")
                    return None
                text = raw_data.decode(encoding, errors='ignore')
                for line in text.splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    parts = line.split("----")
                    if len(parts) != 3:
                        print(
                            f"文本内容格式不正确，请使用 '----' 作为分隔符: {line}")
                        continue
                    accounts.append(
                        (parts[0].strip(), parts[1].strip(), parts[2].strip()))
        except FileNotFoundError:
            print(f"文件 {file_path} 不存在，请检查路径是否正确")
            return None
        except Exception as e:
            print(f"读取文件时出错: {e}")
            return None
        if not accounts:
            print("文本文件中没有有效的账户信息")
            return None
        return accounts
    elif choice == '2':
        account_input = input(
            "请输入账户信息，格式为 邮箱----密码----刷新令牌: ").strip()
        parts = account_input.split("----")
        if len(parts) != 3:
            print("输入格式不正确，请使用 '----' 作为分隔符")
            return None
        user_email, password, refresh_token = parts
        return [(user_email.strip(), password.strip(), refresh_token.strip())]
    else:
        print("无效选项，请输入 1 或 2")
        return None


# 5. 处理单个账户
def process_account(account):
    user_email, password, refresh_token = account
    print(f"\n处理账户: {user_email}")

    # 使用刷新令牌获取新的 Access Token
    access_token = get_accesstoken(refresh_token)
    if access_token:
        mail = connect_imap(user_email, access_token)
        if mail:
            try:
                mail.select('INBOX')
                # 搜索所有邮件，并获取最新一封
                result, data = mail.search(None, 'ALL')
                mail.select('Junk')
                result_2, data_2 = mail.search(None, 'ALL')
                print(data_2)
                while True:

                    if result == 'OK':
                        mail_ids = data[0].split()
                        if mail_ids:
                            latest_email_id = mail_ids[-1]
                            result, msg_data = mail.fetch(latest_email_id,
                                                          '(RFC822)')
                            if result == 'OK':
                                msg = email.message_from_bytes(msg_data[0][1])
                                subject = msg['Subject']
                                from_ = msg['From']
                                print(
                                    f"最新邮件 - 发件人: {from_}, 主题: {subject}")
                                match = re.search(r'(\d{6})', subject)
                                print(f"邮箱地址: {match.group()}")
                                # if msg.is_multipart():
                                #     for part in msg.walk():
                                #         content_type = part.get_content_type()
                                #         content_disposition = str(
                                #             part.get("Content-Disposition"))
                                #         if (content_type == "text/plain" and
                                #                 "attachment" not in
                                #                 content_disposition):
                                #             try:
                                #                 body = part.get_payload(
                                #                     decode=True).decode(
                                #                     part.get_content_charset() or
                                #                     'utf-8',
                                #                     errors='ignore')
                                #                 print("邮件正文（纯文本）：")
                                #                 match = re.search(r'\b\d{6}\b',body)
                                #                 print(body)
                                #                 print(match.group())
                                #                 break
                                #             except Exception as e:
                                #                 print(f"解码邮件正文时出错: {e}")
                                # else:
                                #     content_type = msg.get_content_type()
                                #     if content_type == "text/plain":
                                #         try:
                                #             body = msg.get_payload(
                                #                 decode=True).decode(
                                #                 msg.get_content_charset() or
                                #                 'utf-8',
                                #                 errors='ignore')
                                #             print("邮件正文（纯文本）：")
                                #             print(body)
                                #         except Exception as e:
                                #             print(f"解码邮件正文时出错: {e}")

                        else:
                            print("收件箱中没有邮件。")
                    else:
                        print("无法搜索邮件。")
                    print("等待 15 秒后继续获取最新邮件...")
                    time.sleep(15)
            except Exception as e:
                print(f"获取最新邮件时出错: {e}")
            finally:
                mail.logout()
        else:
            print(f"账户 {user_email} 无法连接 IMAP。")
    else:
        print(f"无法获取账户 {user_email} 的 Access Token。")


# 获取验证码


# 6. 等待用户按下空格键继续
def wait_for_space():
    while True:
        user_input = input("按空格键继续... (按其他键退出): ")
        if user_input == " ":
            break
        else:
            print("退出程序。")
            sys.exit(0)


# 7. 主函数
def main(choice='2'):
    accounts = get_account_info(choice)
    if not accounts:
        print("无法获取账户信息，请检查输入")
        return

    total_accounts = len(accounts)
    for index, account in enumerate(accounts, start=1):
        process_account(account)
        if index < total_accounts:
            wait_for_space()
        else:
            input("所有账户处理完毕。按 Enter 键退出。")


# 程序入口
if __name__ == "__main__":
    main()
