#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Author  : Miss Mango

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
import pandas as pd
import cx_Oracle
from email import encoders

global file_path
file_path = r'C:\Users\18810\Desktop'


def loadDataSet():
    """
    加载数据集
    :param fileName: 文件名称
    :return:list类型
    """
    col = ['EMAIL']
    conn = cx_Oracle.connect("BUS_DA/BUS_DA_20181205@172.16.12.121/biwork")
    # 用自己的实际数据库用户名、密码、主机ip地址 替换即可
    curs = conn.cursor()
    # sql='select distinct email from bus_da.TEST_DATA_END'
    sql = 'SELECT distinct 邮箱 FROM BUS_DA.SQUAD_LEADER'  # sql语句
    curs.execute(sql)
    result = curs.fetchall()
    dataMat = pd.DataFrame(result, columns=col)
    print(dataMat)

    return dataMat


def send_email(sender, password, receiver, smtpserver):

    # 邮件的主题，发件人，收件人
    subject = '业务绩效考核跟进s8-（0701-0902）'
    subject = Header(subject, 'utf-8').encode()  # 中文编码
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver

    # 数据附件
    data_path = file_path + '\\' + '业务绩效考核跟进s8-（0701-0902）.xlsx'
    print(data_path)
    send_data_file = MIMEText(open(data_path, "rb").read(), "base64", "utf-8")
    send_data_file["Content-Type"] = 'application/octet-stream'
    send_data_file.add_header("Content-Disposition","attachment", filename=('业务绩效考核跟进s8-（0701-0902）.xlsx'))
    encoders.encode_base64(send_data_file)
    msg.attach(send_data_file)

    # 邮件发送
    try:
        print ('****** start connect ******')
        smtp = smtplib.SMTP()
        smtp.connect(smtpserver)
        print('connect successfully')
        if sender !='':
            smtp.login(sender, password)
            print('login successfully')
        smtp.sendmail(sender, receiver.split(','), msg.as_string())
        print("Email send successfully")
        smtp.quit()
    except Exception:
        print('Email send failed')




if __name__ == "__main__":
    # 读取数据
    data_list = loadDataSet()

    # 设置smtplib所需的参数
    my_smtpserver = 'smtp.maitian.cn'
    my_sender = 'sjfxb@maitian.cn'
    my_password = 'mtsjfxb2016()'

    # 邮箱地址转成string
    my_receiver = data_list['EMAIL']
    my_receiver = ','.join(my_receiver)

    print(my_receiver)

    # 发邮件
    send_email(my_sender, my_password, my_receiver, my_smtpserver)



