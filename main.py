import datetime
import tools_common

def main():
    sender = 'confluence@siyecapital.com'
    receivers = ['cherryplant1@163.com']
    cc_emails = ['535674963@qq.com', '328247921@qq.com']
    password = 'Siye123456'
    smtp_server = 'smtp.exmail.qq.com'
    tools_common.send_email(sender=sender, receivers=receivers, msg_title='标题C', msg_body='正文B', smtp_server=smtp_server, password=password,
        cc_emails=cc_emails, attachment_filepath='test_folder/测试.txt')


class Test():
    def fun(self):
        self.fun.x = 2
        print(self.fun.x)

if __name__ == "__main__":
    main()