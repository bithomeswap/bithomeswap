import os
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

class Mail:
    def __init__(self, sender_email, receiver_email, password, subject, body,
                 excel_file_path=None, jpg_file_path=None):
        self.sender_email = sender_email
        self.receiver_email = receiver_email
        self.password = password
        self.subject = subject
        self.body = body
        self.excel_file_path = excel_file_path
        self.jpg_file_path = jpg_file_path
    @staticmethod
    def _add_attachment(message, file_path, mime_subtype, filename):
        """通用的添加附件方法（静态方法）"""
        if file_path and os.path.exists(file_path):
            try:
                with open(file_path, 'rb') as f:
                    file_data = f.read()
                attachment = MIMEApplication(file_data, _subtype=mime_subtype)
                attachment.add_header('Content-Disposition', 'attachment', filename=filename)
                message.attach(attachment)
                print(f"已添加附件: {filename}")
            except FileNotFoundError:
                print(f"文件不存在: {filename}")
            except Exception as e:
                print(f"添加附件失败: {filename} - {str(e)}")
    def send_mail(self):
        message = MIMEMultipart()
        message['Subject'] = self.subject
        message['From'] = self.sender_email
        message['To'] = self.receiver_email
        # 添加邮件正文
        message.attach(MIMEText(self.body, "plain", "utf-8"))
        # 添加Excel附件（如果存在）
        self._add_attachment(
            message,
            self.excel_file_path,
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            str(self.excel_file_path)#文件名
        )
        # 添加图片附件（如果存在）
        self._add_attachment(
            message,
            self.jpg_file_path,
            "jpg",
            "发送显示图片名.jpg"
        )
        try:
            with smtplib.SMTP("smtp.163.com", 25) as server:
                server.starttls()
                server.login(self.sender_email, self.password)
                server.sendmail(
                    self.sender_email,
                    self.receiver_email,
                    message.as_string()
                )
                print(f'邮件发送成功 {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        except smtplib.SMTPException as error:
            print(f"SMTP错误: {str(error)}")
        except Exception as e:
            print(f'邮件发送失败: {str(e)} {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')

import pandas as pd
df=pd.read_csv("天眼查河北教育邮箱.csv")
for index,thisdf in df.iterrows():
  thismail=thisdf["邮箱列表"]
  con=thisdf["省份"]
  city=thisdf["城市"]
  name=thisdf["公司简称"]
  mailer1 = Mail(
      sender_email="19511189162@163.com",
      receiver_email=thismail,
      password="QDTSX5aUM26Z2Qnw",
      subject="主题-合作洽谈",
      body=("问候"
            f"\n{con}{city}{name}您好"
            "\n我叫王腾鹤，高中衡中实验班（高三班主任郭春雨现任衡中校长），本科河北工业大学，本职业工作在一家北京的留学机构，教授本科和研究生阶段学生数据分析和金融工程，目前个人在做高考志愿填报，以下资料请您斧正，希望跟你建立联系共同发展合作。"
            "\n微信：lianghuajiaoyi123456"),
      excel_file_path="./高考志愿参考资料.pdf",  # 存在的文件路径
  )
  mailer1.send_mail()
# # 示例1：发送带两个附件的邮件
# mailer1 = Mail(
#     sender_email="发件人@163.com",
#     receiver_email="收件人邮箱",
#     password="授权码-到163设置-POP3/SMTP/IMAP-开启服务：IMAP/SMTP服务-复制授权密码到这",
#     subject="主题-附件-",
#     body=("问候"
#           "\n你好"
#           "\nTel:XXX"),
#     excel_file_path="file/sample.xlsx",  # 存在的文件路径
#     jpg_file_path="file/a.jpg"  # 存在的文件路径 file/sample.xlsx 在.py中创建了file文件夹将文件放在了此文件夹
# )
# mailer1.send_mail()

# # 示例2：只发送Excel附件
# mailer2 = Mail(
#     sender_email="发件人@163.com",
#     receiver_email="收件人邮箱",
#     password="授权码-到163设置-POP3/SMTP/IMAP-开启服务：IMAP/SMTP服务-复制授权密码到这",
#     subject="主题-附件",
#     body=("问候"
#           "\n你好"
#           "\nTel:XXX"),
#     excel_file_path="file/sample.xlsx"  # 存在的文件路径 file/sample.xlsx 在.py中创建了file文件夹将文件放在了此文件夹
# )
# mailer2.send_mail()

# # 示例3：无附件邮件
# mailer3 = Mail(
#     sender_email="19511189162@163.com",
#     receiver_email="1348006516@qq.com",
#     password="QDTSX5aUM26Z2Qnw",
#     subject="主题",
#     body=("问候"
#           "\n你好"
#           "\nTel:XXX"),
# )
# mailer3.send_mail()