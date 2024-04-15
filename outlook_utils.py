# -*- coding: utf-8 -*-

import typer
import pywintypes
from win32com import client
from win32com.client.gencache import EnsureDispatch as Dispatch
from prettytable import PrettyTable

app = typer.Typer()
table = PrettyTable()

class OutlookUtilsBase:
    """
    Outlook Utils 基础类, 用于实现获取邮件内容
    """

    def __init__(self, email_addr):
        # 初始化类变量
        self.email_addr = email_addr

    def get_emails(self, email_addr=None, max_emails: int = 100, filter_by_folder: str = "收件箱, Inbox"):
        """
        获取邮件内容并返回邮件对象的列表, 默认返回获取的前100封邮件对象
        :return: list or False
        """
        email_items = []
        if email_addr is None:
            email_addr = self.email_addr
        if max_emails == -1 or max_emails >= 1:
            max_emails = int(max_emails)
        filter_by_folder = [i.strip() for i in filter_by_folder.split(',')]

        # 构造 Outlook 对象
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        folders = outlook.Folders
        for account_folder in folders:
            # 如果找到目标邮箱
            if account_folder.Name == email_addr:
                # 获取收件夹并排序
                for folder in account_folder.Folders:
                    # 如果收件夹在过滤列表中
                    if folder.Name in filter_by_folder:
                        try:
                            index = folder.Items
                            index.Sort("[ReceivedTime]", True)
                            for item in index:
                                email_items.append(item)
                        except pywintypes.com_error as e:
                            # 预期报错, 有的文件夹无法排序, 因此这部分错误直接忽略即可
                            pass

        # 判断获取的邮件
        if max_emails == -1 and len(email_items) > 0:
            return email_items
        elif 0 < max_emails < len(email_items):
            return email_items[:max_emails]
        elif max_emails > 0 and len(email_items) <= max_emails:
            return email_items
        else:
            return False

    def send_email(self, from_addr=None, to_addr=None, cc_addr=None, subject=None, content=None, attachment=None):
        """
        发送邮件
        """
        # 初始化需要的变量
        to_addr = self.email_addr if to_addr is None else to_addr
        if to_addr is not None:
            to_addr = [i.strip() for i in to_addr.split(';')]
        cc_addr = self.email_addr if cc_addr is None else cc_addr
        if cc_addr is not None:
            cc_addr = [i.strip() for i in cc_addr.split(';')]
        subject = "Test for OutlookUtils" if subject is None else subject
        content = "This is test for OutlookUtils" if content is None else content
        # 构建邮件对象
        outlook = client.Dispatch("Outlook.Application")
        mail_item = outlook.CreateItem(0)
        # 设置发件人, 需要设置权限, 一般无效
        # https://learn.microsoft.com/zh-cn/Exchange/recipients/mailbox-permissions?redirectedfrom=MSDN&view=exchserver-2019&preserve-view=true
        if from_addr is not None:
            mail_item.SentOnBehalfOfName = from_addr
        # 添加收件人
        if isinstance(to_addr, list):
            for addr in to_addr:
                mail_item.Recipients.Add(addr)
        # 添加抄送
        if isinstance(cc_addr, list):
            for addr in cc_addr:
                mail_item.Recipients.Add(addr).Type = 2
        # 添加附件
        if attachment is not None:
            mail_item.Attachments.Add(attachment)
        # 标题和内容, 2 代表为 HTML 格式的邮件
        mail_item.Subject = subject
        mail_item.BodyFormat = 2
        mail_item.HTMLBody = "<div>{}</div>".format(content)

        # 发送邮件
        mail_item.Send()


@app.command()
def send_email(
        to_addr: str = typer.Option(help="收件人邮箱地址, 支持多个写入多个邮箱, 使用 ';' 分割"),
        cc_addr: str = typer.Option(None, help="抄送的邮箱地址, 支持多个写入多个邮箱, 使用 ';' 分割"),
        from_addr: str = typer.Option(None, help="发件人邮箱地址"),
        subject: str = typer.Option("Test for OutlookUtils", help="邮箱标题"),
        content: str = typer.Option("This is test for OutlookUtils", help="邮箱正文"),
        attachment: str = typer.Option(None, help="附件路径"),
):
    """
    发送邮件 / Send Email from local Outlook Client
    """
    # 初始化类变量
    outlook_utils = OutlookUtilsBase(to_addr)
    # 发送邮件
    outlook_utils.send_email(from_addr, to_addr, cc_addr, subject, content, attachment)

@app.command()
def get_emails_subject(
        email_addr: str = typer.Option(help="邮箱地址"),
        max_emails: int = typer.Option(100, help="最大邮件数量, -1 代表没限制"),
        filter_by_folder: str = typer.Option("收件箱, Inbox", help="检索邮件的文件夹")
):
    """
    获取邮件的标题 / Subject of Email
    """
    table.field_names = ["ID", "Subject"]
    email_id = 1
    outlook_utils = OutlookUtilsBase(email_addr)
    for mail in outlook_utils.get_emails(email_addr, max_emails, filter_by_folder):
        table.add_row([email_id, mail.Subject])
        email_id += 1
    print(table)

@app.command()
def get_emails_summary(
        email_addr: str = typer.Option(help="邮箱地址"),
        max_emails: int = typer.Option(100, help="最大邮件数量, -1 代表没限制"),
        filter_by_folder: str = typer.Option("收件箱, Inbox", help="检索邮件的文件夹")
):
    """
    邮件的汇总信息 / Summary of Email Information
    """
    table.field_names = ["ID", "Subject", "SenderName", "Recipients", "ReceivedTime"]
    email_id = 1
    outlook_utils = OutlookUtilsBase(email_addr)
    for mail in outlook_utils.get_emails(email_addr, max_emails, filter_by_folder):
        Recipient_list = str([i.Name for i in mail.Recipients])[1:-1].replace("', '", "; ")
        table.add_row([email_id, mail.Subject, mail.SenderName, Recipient_list, mail.ReceivedTime])
        email_id += 1
    print(table)


if __name__ == "__main__":
    app()
