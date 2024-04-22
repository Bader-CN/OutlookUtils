# -*- coding: utf-8 -*-

import re
import os
import typer
import pywintypes
import pandas as pd
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
        to_addr: str = typer.Option(help="Recipient email address, supports multiple emails by writing them separated by ';'"),
        cc_addr: str = typer.Option(None, help="CC email addresses, multiple email addresses can be written by using ';' as a separator"),
        from_addr: str = typer.Option(None, help="Sender's email address."),
        subject: str = typer.Option("Test for OutlookUtils", help="Email Subject"),
        content: str = typer.Option("This is test for OutlookUtils", help="Email Content"),
        attachment: str = typer.Option(None, help="Attachment Path"),
):
    """
    Send Email from local Outlook Client
    """
    # 初始化类变量
    outlook_utils = OutlookUtilsBase(to_addr)
    # 发送邮件
    outlook_utils.send_email(from_addr, to_addr, cc_addr, subject, content, attachment)


@app.command()
def get_emails_subject(
        email_addr: str = typer.Option(help="Email Address"),
        max_emails: int = typer.Option(100, help="Maximum number of emails, -1 means no limit"),
        filter_by_folder: str = typer.Option("收件箱, Inbox", help="Filter by email folders")
):
    """
    Get the subject of the email
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
        email_addr: str = typer.Option(help="Email Address"),
        max_emails: int = typer.Option(100, help="Maximum number of emails, -1 means no limit"),
        filter_by_folder: str = typer.Option("收件箱, Inbox", help="Filter by email folders")
):
    """
    Summary of Email Information
    """
    table.field_names = ["ID", "Subject", "SenderName", "Recipients", "ReceivedTime"]
    email_id = 1
    outlook_utils = OutlookUtilsBase(email_addr)
    for mail in outlook_utils.get_emails(email_addr, max_emails, filter_by_folder):
        recipient_list = str([i.Name for i in mail.Recipients])[1:-1].replace("', '", "; ")
        table.add_row([email_id, mail.Subject, mail.SenderName, recipient_list, mail.ReceivedTime])
        email_id += 1
    print(table)


@app.command()
def generate_sf_monthly_report(
        email_addr: str = typer.Option(help="Email Address"),
        raw_cases_report: str = typer.Option(None, help="Attachment Name Prefix: Original Cases Report, Format: <report_name>-%Y-%m-%d-%H-%M-%S.csv"),
        raw_survey_report: str = typer.Option(None, help="Attachment Name Prefix: Original Survey Report, Format: <report_name>-%Y-%m-%d-%H-%M-%S.csv"),
        max_emails: int = typer.Option(100, help="Maximum number of emails, -1 means no limit"),
        filter_by_folder: str = typer.Option("收件箱, Inbox", help="Filter by email folders"),
        month_offset: int = typer.Option(0, help="Month offset, please enter a negative value. default is 0, which means to calculate the information for the current month"),
        output_file: str = typer.Option(None, help="Save the file with the current path and name, in csv format"),
):
    """
    Generate SalesForce Monthly Report
    """
    # 至少要保证指定了一个原始报告
    if raw_cases_report is None and raw_survey_report is None:
        print("At least one report must be specified!")
        exit(0)

    # 如果筛选不到指定的附件, 则退出
    csv_data = []
    case_list = []
    surv_list = []
    case_list_obj = []
    surv_list_obj = []
    outlook_utils = OutlookUtilsBase(email_addr)
    for mail in outlook_utils.get_emails(email_addr, max_emails, filter_by_folder):
        attachments = mail.Attachments
        if attachments.Count > 0:
            for attachment in attachments:
                if raw_cases_report is not None and re.match(raw_cases_report, attachment.FileName, re.IGNORECASE):
                    case_list.append(attachment.FileName)
                    case_list_obj.append(attachment)
                if raw_survey_report is not None and re.match(raw_survey_report, attachment.FileName, re.IGNORECASE):
                    surv_list.append(attachment.FileName)
                    surv_list_obj.append(attachment)

    # 计算指定的年月
    if pd.Timestamp.now().month + month_offset >= 1:
        y_offset = pd.Timestamp.now().year
        m_offset = pd.Timestamp.now().month + month_offset
    elif pd.Timestamp.now().month + month_offset == 0:
        y_offset = pd.Timestamp.now().year - 1
        m_offset = 12
    elif pd.Timestamp.now().month + month_offset > -12:
        y_offset = pd.Timestamp.now().year - 1
        m_offset = 12 - (abs(pd.Timestamp.now().month + month_offset) % 12)
    else:
        y_offset = pd.Timestamp.now().year - (abs(pd.Timestamp.now().month + month_offset) // 12) - 1
        m_offset = 12 - (abs(pd.Timestamp.now().month + month_offset) % 12)
    # 生成表头
    table.field_names = ["KPI", "{}-{}".format(str(y_offset), str(m_offset))]

    # 处理 Case Report
    if len(case_list) > 0:
        try:
            case_list.sort(reverse=True)
            raw_case_report_name = case_list[0]
            raw_case_report_path = os.path.abspath(os.path.join("./", raw_case_report_name))
            for attachment in case_list_obj:
                if attachment.FileName == raw_case_report_name:
                    attachment.SaveAsFile(raw_case_report_path)
                    # 读取文件并计算分析数据
                    rawcase = pd.read_csv(raw_case_report_path)
                    # 数据预处理
                    rawcase["Date/Time Opened"] = pd.to_datetime(rawcase["Date/Time Opened"], format="%Y-%m-%d %p%I:%M")
                    rawcase["Closed Date"] = pd.to_datetime(rawcase["Closed Date"], format="%Y-%m-%d")
                    # 根据年份和月份筛选数据
                    open_cases_y = rawcase[rawcase["Date/Time Opened"].dt.year == y_offset]
                    open_cases_m = open_cases_y[open_cases_y["Date/Time Opened"].dt.month == m_offset]
                    close_cases_y = rawcase[rawcase["Closed Date"].dt.year == y_offset]
                    close_cases_m = close_cases_y[close_cases_y["Closed Date"].dt.month == m_offset]
                    # 计算当前状态下状态为非 Closed 的 cases
                    backlog = rawcase[rawcase["Status"] != "Closed"]
                    backlog = backlog[backlog["Date/Time Opened"] <= pd.Timestamp(y_offset, m_offset, 1) + pd.offsets.MonthEnd()]
                    backlog_history = rawcase[rawcase["Closed Date"] > pd.to_datetime("{}-{}".format(y_offset, m_offset, 1), format="%Y-%m") + pd.offsets.MonthEnd()]
                    backlog_history = backlog_history[backlog_history["Date/Time Opened"] <= pd.Timestamp(y_offset, m_offset, 1) + pd.offsets.MonthEnd()]
                    # KCS 相关
                    kcs_all = close_cases_m[close_cases_m["Knowledge Base Article"].notna() | close_cases_m["Idol Knowledge Link"].notna()]
                    # 分析数据并得出结果
                    table.add_row(["Open Cases", len(open_cases_m)])
                    csv_data.append(["Open Cases", len(open_cases_m)])
                    table.add_row(["Close Cases", len(close_cases_m)])
                    csv_data.append(["Close Cases", len(close_cases_m)])
                    # Closure Rate
                    if len(open_cases_m) != 0:
                        table.add_row(["Closure Rate", (str(round(len(close_cases_m) / len(open_cases_m) * 100, 2)) + "%")])
                        csv_data.append(["Closure Rate", (str(round(len(close_cases_m) / len(open_cases_m) * 100, 2)) + "%")])
                    else:
                        table.add_row(["Closure Rate", "-"])
                        csv_data.append(["Closure Rate", "-"])
                    # R&D Assist Rate
                    if len(close_cases_m) != 0:
                        table.add_row(["R&D Assist Rate", str(round(len(close_cases_m[close_cases_m["R&D Incident"].notna()]) / len(close_cases_m) * 100, 2)) + "%"])
                        csv_data.append(["R&D Assist Rate", str(round(len(close_cases_m[close_cases_m["R&D Incident"].notna()]) / len(close_cases_m) * 100, 2)) + "%"])
                    else:
                        table.add_row(["R&D Assist Rate", "-"])
                        csv_data.append(["R&D Assist Rate", "-"])
                    # Backlog
                    table.add_row(["Backlog", len(backlog) + len(backlog_history)])
                    csv_data.append(["Backlog", len(backlog) + len(backlog_history)])
                    # Backlog > 30
                    if month_offset == 0:
                        try:
                            table.add_row(["Backlog > 30", len(backlog[backlog["Age (Days)"] >= 30.0])])
                            csv_data.append(["Backlog > 30", len(backlog[backlog["Age (Days)"] >= 30.0])])
                        except KeyError:
                            table.add_row(["Backlog > 30", len(backlog[backlog["Age"] >= 30.0])])
                            csv_data.append(["Backlog > 30", len(backlog[backlog["Age"] >= 30.0])])
                    else:
                        table.add_row(["Backlog > 30", "-"])
                        csv_data.append(["Backlog > 30", "-"])
                    # Backlog Index
                    if len(open_cases_m) != 0:
                        table.add_row(["Backlog Index", str(round(len(backlog) / len(open_cases_m) * 100, 2)) + "%"])
                        csv_data.append(["Backlog Index", str(round(len(backlog) / len(open_cases_m) * 100, 2)) + "%"])
                    else:
                        table.add_row(["Backlog Index", "-"])
                        csv_data.append(["Backlog Index", "-"])
                    # KCS Articles Created
                    table.add_row(["KCS Articles Created", len(close_cases_m[close_cases_m["Knowledge Base Article"].notna()])])
                    csv_data.append(["KCS Articles Created", len(close_cases_m[close_cases_m["Knowledge Base Article"].notna()])])
                    # KCS Created / Closed Cases
                    if len(close_cases_m) != 0:
                        table.add_row(["KCS Created / Closed Cases", str(round(len(close_cases_m[close_cases_m["Knowledge Base Article"].notna()]) / len(close_cases_m) * 100, 2)) + "%"])
                        csv_data.append(["KCS Created / Closed Cases", str(round(len(close_cases_m[close_cases_m["Knowledge Base Article"].notna()]) / len(close_cases_m) * 100, 2)) + "%"])
                    else:
                        table.add_row(["KCS Created / Closed Cases", "-"])
                        csv_data.append(["KCS Created / Closed Cases", "-"])
                        # KCS Linkage
                        table.add_row(["KCS Linkage", str(round(len(kcs_all) / len(close_cases_m) * 100, 2)) + "%"])
                        csv_data.append(["KCS Linkage", str(round(len(kcs_all) / len(close_cases_m) * 100, 2)) + "%"])
                    # 删除文件
                    try:
                        os.remove(raw_case_report_path)
                    except Exception as e:
                        pass
        except KeyError as e:
            print("The specified column is missing in the report. Please ensure that the specified report is correct and contains the required columns.")
            print("Error message:{}".format(str(e)))
            exit(1)

    if len(surv_list) > 0:
        try:
            surv_list.sort(reverse=True)
            raw_surv_report_name = surv_list[0]
            raw_surv_report_path = os.path.abspath(os.path.join("./", raw_surv_report_name))
            for attachment in surv_list_obj:
                if attachment.FileName == raw_surv_report_name:
                    attachment.SaveAsFile(raw_surv_report_path)
                    # 读取文件并计算分析数据
                    rawsurv = pd.read_csv(raw_surv_report_path)
                    # 数据预处理
                    rawsurv["Customer Feed Back Survey: Last Modified Date"] = pd.to_datetime(rawsurv["Customer Feed Back Survey: Last Modified Date"], format="%Y-%m-%d")
                    rawsurv["Closed Data"] = pd.to_datetime(rawsurv["Closed Data"], format="%Y-%m-%d")
                    rawsurv = rawsurv.sort_values(by=["Customer Feed Back Survey: Last Modified Date", ], ascending=False)
                    rawsurv = rawsurv.drop_duplicates(subset="Case Number")
                    # 根据年份和月份筛选数据
                    survey_y = rawsurv[rawsurv["Closed Data"].dt.year == y_offset]
                    survey_m = survey_y[survey_y["Closed Data"].dt.month == m_offset]
                    survey_ces = survey_m[survey_m["OpenText made it easy to handle my case"] >= 8.0]
                    survey_cast = survey_m[survey_m["Satisfied with support experience"] >= 7.0]
                    # 分析数据并得出结果
                    if len(survey_m) > 0:
                        table.add_row(["Survey CES", str(round(len(survey_ces) / len(survey_m) * 100, 2)) + "%"])
                        csv_data.append(["Survey CES", str(round(len(survey_ces) / len(survey_m) * 100, 2)) + "%"])
                        table.add_row(["Survey CAST", str(round(len(survey_cast) / len(survey_m) * 100, 2)) + "%"])
                        csv_data.append(["Survey CAST", str(round(len(survey_cast) / len(survey_m) * 100, 2)) + "%"])
                    else:
                        table.add_row(["Survey CES", "-"])
                        csv_data.append(["Survey CES", "-"])
                        table.add_row(["Survey CAST", "-"])
                        csv_data.append(["Survey CES", "-"])
                    # 删除文件
                    try:
                        os.remove(raw_surv_report_path)
                    except Exception as e:
                        pass
        except KeyError as e:
            print("The specified column is missing in the Survey Report. Please ensure that the specified report is correct and contains the required columns.")
            print("Error message:{}".format(str(e)))
            exit(1)

    if len(case_list) == 0 and len(surv_list) == 0:
        print(" No specified report found!")
        exit(0)
    if output_file is None:
        # 打印结果
        print(table)
    else:
        if not re.findall(r"\.csv$", output_file):
            output_file = output_file + ".csv"
        df = pd.DataFrame(csv_data, columns=csv_data[0])
        df.to_csv(output_file, index=False, header=["KPI", "{}-{}".format(str(y_offset), str(m_offset))])


if __name__ == "__main__":
    app()
