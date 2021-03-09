# coding=utf-8
import random
import smtplib
from email.header import Header
from email.mime.text import MIMEText

class Supervisor:
    # 全局运行监视
    def usage_supervisor(self):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('.\\resources\\license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('Well integrity software is running', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(10000000, 100000000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' Well integrity software is running')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()
        print('网络成功连接')

    # 加载原始记录登记表模块运行监视
    def load_raw_table_usage_supervisor(self, well_Name):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('.\\resources\\license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('Raw table info is loading', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(500, 1000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' ' + well_Name + ' raw table info is loading')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()

    # 生成LEAD TXT模块运行监视
    def lead_txt_usage_supervisor(self):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('.\\resources\\license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('LEAD TXT is creating', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(500, 1000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' LEAD TXT is creating')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()

    # 生成水泥胶结报告模块运行监视
    def generate_report_usage_supervisor(self):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('.\\resources\\license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('CBL/VDL report is creating', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(500, 1000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' CBL/VDL report is creating')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()

    # 生成快速解释结论模块运行监视
    def generate_CHL_result_usage_supervisor(self):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('.\\resources\\license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('CHL fast report is creating', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(500, 1000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' CHL fast report is creating')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()

    # 生成签名模块运行监视
    def generate_signature_usage_supervisor_100(self):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('.\\resources\\license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('Signature on 100 pixel map is creating', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(500, 1000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' Signature on 100 pixel map is creating')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()

    # 生成签名模块运行监视
    def generate_signature_usage_supervisor_150(self):
        # 发信方的信息：发信邮箱，QQ 邮箱授权码
        from_addr = 'kiwifruitloves123@qq.com'
        with open('license_email.txt', "r") as f:
            license_str = f.read()
        password = license_str

        # 收信方邮箱
        to_addr = '978030836@qq.com'

        # 发信服务器
        smtp_server = 'smtp.qq.com'

        # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
        msg = MIMEText('Signature on 150 pixel map is creating', 'plain', 'utf-8')

        # 邮件头信息
        random_num = random.randint(500, 1000)
        random_num = str(random_num)
        msg['From'] = Header(from_addr)
        msg['To'] = Header(to_addr)
        msg['Subject'] = Header(random_num + ' Signature on 150 pixel map is creating')

        # 开启发信服务，这里使用的是加密传输
        server = smtplib.SMTP_SSL('smtp.qq.com')
        server.connect(smtp_server, port=465)
        # 登录发信邮箱
        server.login(from_addr, password)
        # 发送邮件
        server.sendmail(from_addr, to_addr, msg.as_string())
        # 关闭服务器
        server.quit()