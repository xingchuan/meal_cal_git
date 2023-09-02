import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText


# 此函数可通用，实现1、删除多余的列；2、重命名列名
def excel_exec(excel_name):
    # 读入excel表格
    df = pd.read_excel(excel_name, header=None)
    # 删除第一行
    df.drop(0, inplace=True)
    # 将第一行作为列名
    df.rename(columns=df.iloc[0], inplace=True)
    # 删除第二行，第一行作为列名后，会在第二行重复出现
    df.drop(1, inplace=True)
    # 删除不需要的列
    df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14]], axis=1)
    # 重命名列名
    df.rename(columns={
        '姓名\nName': '姓名',
        '部门\nDepartment': '部门',
        '人员编号\nPerson ID': '工号',
        '识别时间\nRecognition time': '识别时间'
    },
        inplace=True)
    return df


df1 = excel_exec('./uploads/食堂1.xlsx')
df2 = excel_exec('./uploads/食堂2.xlsx')

# 合并食堂1、食堂2的数据
df = pd.concat([df1, df2], axis=0)
# 将合并后的表格保存为新的Excel文件
df = df[df.iloc[:, 0] != '人员未注册']
df['部门'] = df['部门'].str.split('/').str[-1]
df.to_excel('merge.xlsx', index=False)


# 该函数提取相对应时间段内的打卡记录
def faceRecords():
# 读取原始表格数据
    df = pd.read_excel("merge.xlsx")

    # 将日期时间列解析为datetime类型
    df['识别时间'] = pd.to_datetime(df['识别时间'])

    # 选择需要提取的时间范围
    morning_range = pd.date_range(start="07:40:00", end="09:10:00", freq="T")
    afternoon_range = pd.date_range(start="17:00:00", end="18:10:00", freq="T")

    # 提取符合时间范围的打卡记录
    selected_records = df[df['识别时间'].dt.time.between(morning_range.time.min(), morning_range.time.max()) |
                        df['识别时间'].dt.time.between(afternoon_range.time.min(), afternoon_range.time.max())]

    # 去除同一天同一个时间段内同一个人的重复记录，只保留第一条记录
    selected_records = selected_records.drop_duplicates(subset=['姓名', '识别时间'], keep='first').copy()

    # 按照姓名排序
    selected_records.sort_values(by=['姓名', '识别时间'], inplace=True)

    # 保存到新的表格中
    selected_records.to_excel("face_records.xlsx", index=False)


faceRecords()


def toEmails():
    # 以下代码通过邮件发送附件
    # 电子邮件参数
    from_email = "1558351557@qq.com"    # 发件人邮箱
    password = "osnkujiavlmajdbh"       # 发件人邮箱密码
    to_emails = ["251696664@qq.com", "gu.xingchuan@rml138.com"]  # 收件人邮箱列表
    subject = "餐费计算"                # 邮件主题
    body = "请查收附件"                 # 邮件正文

    # 构建邮件对象
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ', '.join(to_emails)  # 将收件人列表转为逗号分隔的字符串
    msg['Subject'] = subject

    # 添加邮件正文
    msg.attach(MIMEText(body, 'plain'))

    # 添加Excel附件
    attachment_paths = [
        "face_records.xlsx",  # 第一个附件的路径
        # "每人次数.xlsx",  # 第二个附件的路径
    ]

    for attachment_path in attachment_paths:
        with open(attachment_path, "rb") as file:
            attachment = file.read()
        excel_attachment = MIMEApplication(attachment)
        excel_attachment.add_header('Content-Disposition', 'attachment', filename=f'{attachment_path}')
        msg.attach(excel_attachment)

    # 发送邮件
    try:
        with smtplib.SMTP_SSL("smtp.qq.com", 465) as server:
            server.login(from_email, password)
            server.sendmail(from_email, to_emails, msg.as_string())  # 将收件人列表传递给sendmail()方法
        print("邮件发送成功！")
    except Exception as e:
        print(f"邮件发送失败：{e}")
    
toEmails()