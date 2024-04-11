# -*- coding: utf-8 -*-

import typer
import pywintypes
from win32com.client.gencache import EnsureDispatch as Dispatch


class OutlookUtilsBase:
    """
    Outlook Utils 基础类, 用于实现获取邮件内容
    """
    def __init__(self, email_addr, max_emails: int = 100, filter_by_folder: str = "收件箱, inbox"):
        self.email_addr = email_addr
        # 最大获取的邮件数量
        if max_emails == -1 or max_emails >= 1:
            self.max_emails = int(max_emails)
        else:
            self.max_emails = 100
        # 需要过滤的收件夹
        self.filter_by_folder = [i.strip() for i in filter_by_folder.split(',')]

    def get_emails(self):
        """
        获取邮件内容并返回邮件对象的列表
        :return: list or False
        """
        email_items = []
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        folders = outlook.Folders
        for account_folder in folders:
            # 如果找到目标邮箱
            if account_folder.Name == self.email_addr:
                # 获取收件夹并排序
                for folder in account_folder.Folders:
                    # 如果收件夹在过滤列表中
                    if folder.Name in self.filter_by_folder:
                        try:
                            index = folder.Items
                            index.Sort("[ReceivedTime]", True)
                            for item in index:
                                email_items.append(item)
                        except pywintypes.com_error as e:
                            # 预期报错, 有的文件夹无法排序, 因此这部分错误直接忽略即可
                            pass

        # 判断获取的邮件
        if len(email_items) > self.max_emails:
            return email_items[:self.max_emails]
        elif len(email_items) <= self.max_emails:
            return email_items
        else:
            return False


if __name__ == "__main__":
    # typer.run(OutlookUtilsBase)
    OutlookUtilsBase(email_addr="demo@demo.com")