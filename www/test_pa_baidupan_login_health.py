#!/usr/bin/env python3.6
# -*- coding:UTF-8 -*-
import datetime
import json
from venv import logger
import xlrd
import xlwt
from bs4 import BeautifulSoup
from pip._vendor.distlib.compat import raw_input
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import requests
from PIL import Image
from collections import OrderedDict
from collections import Counter
import pandas as pd
import itertools
import smtplib
import email
# 负责构造文本
from email.mime.text import MIMEText
# 负责构造图片
from email.mime.image import MIMEImage
# 负责将多个对象集合起来
from email.mime.multipart import MIMEMultipart
from email.header import Header


# 为selenium请求添加头以及作一个初始化，注意请求头格式设置为电脑版浏览器，否则请求的页面会不同，导致后面的元素定位会找不到而报错。
dcap = dict(DesiredCapabilities.PHANTOMJS)
# win10 谷歌浏览器请求头 的格式
dcap['phantomjs.page.settings.userAgent'] = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.87 Safari/537.36")
driver = webdriver.PhantomJS(executable_path='D:/phantomjs/bin/phantomjs.exe', desired_capabilities=dcap)

# 提取所有的xpath（几乎所有的pc浏览器都支持获取xpath）,这里大家不需要自定义修改，除非百度把登录页面重做了。
'''
xpath是啥这里也不做解释了，不懂的自行查阅资料。获取xpath的方式举个例子, 要想获取“登录”标签的xpath，可按如下步骤进行：
1>. 鼠标移动到到你要获取xpath的标签位置，然后鼠标右击选择 ‘审查’
2>. 如下图，元素审查结果将自动定位到该标签，然后继续右击蓝色区域，选择Copy -> Copy XPath， 这个时候XPath文本就已经在你粘贴板中，可直接Ctrl+c粘贴到你想要的位置。
'''
driver.set_page_load_timeout(30)

url = "https://www.wjx.cn/login.aspx"
# 验证码
yzm_xpath = '//*[@id="AntiSpam1_txtValInputCode"]'

# 用户名的输入文本框
user_xpath = '//*[@id="UserName"]'
# 密码的输入文本框
pwd_xpath = '//*[@id="Password"]'

# 登录的Button
login_xpath = '//*[@id="LoginButton"]'


# 这里是模拟点击 切换账号登录标签 前后状态的对比
driver.get(url) # 使用Selenium driver 模拟加载百度登录页面

time.sleep(3) # 等待3s网页加载完毕，否则后面的 截图 或者 元素定位无效，导致报错。

driver.get_screenshot_as_file('./scraping.png') # 对模拟网页实时状态截图

gotologin = driver.find_element_by_xpath(login_xpath) # 使用Selenium driver 定位到 登录按钮
gotologin.click() # 模拟点击 点击登录按钮
time.sleep(1) # 这里其实可以不用sleep函数，因为切换到账号登录的过程只是本地js程序执行，不需要和服务器交互。
driver.get_screenshot_as_file('./scraping_2.png') # 对模拟网页实时状态截图，可与click()之前的截图对比。

gotologin = driver.find_element_by_xpath(yzm_xpath) # 使用Selenium driver 定位到 验证码
gotologin.click() # 模拟点击 验证码
time.sleep(1) # 这里其实可以不用sleep函数，因为切换到账号登录的过程只是本地js程序执行，不需要和服务器交互。
driver.get_screenshot_as_file('./scraping_3.png') # 对模拟网页实时状态截图，可与click()之前的截图对比。

time.sleep(1)
im = Image.open('./scraping_3.png')
im.show()
input_ = raw_input('请输入验证码： ')

time.sleep(8)

health_user_textedit=driver.find_element_by_xpath(user_xpath)
health_pwd_textedit=driver.find_element_by_xpath(pwd_xpath)
yzm_textedit = driver.find_element_by_xpath(yzm_xpath)

health_user_textedit.send_keys('13405803107')
health_pwd_textedit.send_keys('password')
yzm_textedit.send_keys(input_)

driver.get_screenshot_as_file('./scraping_4.png')
login_textedit=driver.find_element_by_xpath(login_xpath)
login_textedit.click()
# # ActionChains是一个动作链，使用动作链与否，其优劣各位自己评判
# actions = ActionChains(driver).click(health_user_textedit).send_keys("13405803107").click(health_pwd_textedit).send_keys("tuenkers123").click(yzm_textedit).send_keys(input_).send_keys(Keys.RETURN)
# 设定动作链之后要调用perform()函数才生效
# actions.perform()
# 等待3s后，再截个图看看当前是什么状态
time.sleep(2)
driver.get_screenshot_as_file('./scraping_5.png')


# 获取cookies保存到本地并且返回此cookies
cookies = driver.get_cookies()
driver.close()

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.87 Safari/537.36'
}

sess = requests.Session()
sess.headers.clear()
# 将selenium的cookies放到session中
for cookie in cookies:
    sess.cookies.set(cookie['name'], cookie['value'])
    url = 'https://www.wjx.cn/wjx/activitystat/viewjoinlist.aspx?'
    data = {'activity': '55678796', 'qid': '1'}
    # post 方式
    # html = sess.post(url, data).text
    html = sess.get(url=url, params=data, headers=headers).text

# print(html)

soup = BeautifulSoup(html, "html.parser")
title = soup.find('div', id='divNotJoin').get_text()[2:]
count = soup.find('div', id='divNotJoinCount').get_text()
# print(type(title))

print(title)
print(count)
today_list = title
# today_list = ['洪立浩', '牟维哲', '朱华成', '路晓飞', '杨枫', '高飞', '方忆', '陈杨']
excel_home = xlrd.open_workbook('D:\\chromeDownload\\pyexcel\\2-21.xls')

home_sheet = excel_home.sheet_by_name('Sheet1')  # 正确的字典表格sheet

list_all_name = home_sheet.col_values(6)  # 字典表里面的姓名 符合姓名
list_all_name.pop(0)  # 去除标题

list_all_dept = home_sheet.col_values(7)  # 字典表中的部门  顺序
list_all_dept.pop(0)

today = datetime.date.today()

dict_ = dict(zip(list_all_name, list_all_dept))

mydic_r_ = {}
for name_, dept_ in dict_.items():
    if name_ in today_list:
        mydic_r_[name_] = dept_
print(mydic_r_)

deal_dic = mydic_r_

# pandas ---------------------------------------------------------------------

# df = pd.read_excel('D:\\chromeDownload\\pyexcel\\testpandas.xlsx')
frame = pd.DataFrame.from_dict(deal_dic, orient='index', columns=['values'])
# print(frame)
frame2 = frame.reset_index()
# print(frame2)
frame3 = frame2.rename(columns={'index': 'kys'})
# print(frame3)
# df.groupby('部门')['姓名'].agg(','.join).str.split(',', expand=True)
# print(df.groupby('部门')['姓名'].agg(','.join).str.split(',', expand=True))


df3 = frame3.groupby('values')['kys'].agg(','.join).str.split(',', expand=True)
# print(df3)
# print(frame3.groupby('values')['kys'].agg(','.join).str.split(',', expand=True))
# for indexs in df3.index:
#     print(df3.loc[indexs].values[0:-1])

re_frame = df3.reset_index()
print(re_frame)
list_str = []
print_str = ''
for indexs in re_frame.index:
    print(re_frame.loc[indexs].values)
    # 过滤nono
    # list_str.append(re_frame.loc[indexs].values)

    print_str = print_str + str(list(filter(None, re_frame.loc[indexs].values))).replace('[', '').replace(']', '').replace("'", '').replace(",", ":", 1) + '\n'
# print(list_str)
# print(print_str)

str_head = str(today) + '  今日无发烧'
str_mid = ''
if (len(today_list) == 0):
    str_mid = '\n所有同事都填写了健康登记\n'
else:
    str_mid = '\n截至'+str(time.strftime('%H:%M:%S',time.localtime(time.time())))+'未填写健康登记的有：\n'

print_str = str_head + str_mid + print_str

print(print_str)
# print(frame3.groupby('values')['kys'].agg(','.join).str.split(',', expand=True))
# str_ = str(frame3.groupby('values')['kys'].agg(','.join).str.split(',', expand=True))
# print(str_)


# pandas-------------------------------------------------------------------------

# 发送邮件
mail_str = 'Dear Xie:   \n' + print_str
# SMTP服务器,这里使用163邮箱
mail_host = "smtp.163.com"
# 发件人邮箱
mail_sender = "hlh246@163.com"
# 邮箱授权码,注意这里不是邮箱密码,如何获取邮箱授权码,请看本文最后教程
mail_license = "password"
# 收件人邮箱，可以为多个收件人
mail_receivers = ["lihao.hong@tuenkers.com.cn", "hlh246@outlook.com"]
# 构建MIMEMultipart对象代表邮件本身，可以往里面添加文本、图片、附件等
mm = MIMEMultipart('related')

# 邮件主题 设置邮件头部内容
subject_content = str(today) + """ 健康登记汇报"""
# 设置发送者,注意严格遵守格式,里面邮箱为发件人邮箱
mm["From"] = "sender_name<hlh246@163.com>"
# 设置接受者,注意严格遵守格式,里面邮箱为接受者邮箱
mm["To"] = "receiver_1_name<lihao.hong@tuenkers.com.cn>,receiver_2_name<hlh246@outlook.com>"
# 设置邮件主题
mm["Subject"] = Header(subject_content, 'utf-8')

# 邮件正文内容 添加正文文本
body_content = mail_str
# 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
message_text = MIMEText(body_content, "plain", "utf-8")
# 向MIMEMultipart对象中添加文本对象
mm.attach(message_text)

# 二进制读取图片 6、添加图片
# image_data = open('D:\\chromeDownload\\pyexcel\\1.jpg','rb')
# 设置读取获取的二进制数据
# message_image = MIMEImage(image_data.read())
# 关闭刚才打开的文件
# image_data.close()
# 添加图片文件到邮件信息当中去
# mm.attach(message_image)


# 构造附件 7、添加附件(excel表格)
# atta = MIMEText(open('D:\\chromeDownload\\pyexcel\\3-6.xlsx', 'rb').read(), 'base64', 'utf-8')
# 设置附件信息
# atta["Content-Disposition"] = 'attachment; filename="today.xlsx"'
# 添加附件到邮件信息当中去
# mm.attach(atta)


# 创建SMTP对象发送邮件
stp = smtplib.SMTP()
# 设置发件人邮箱的域名和端口，端口地址为25
stp.connect(mail_host, 25)
# set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
stp.set_debuglevel(1)
# 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
stp.login(mail_sender, mail_license)
# 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
stp.sendmail(mail_sender, mail_receivers, mm.as_string())
print("邮件发送成功")
# 关闭SMTP对象
stp.quit()

'''
# 保存cookie至本地
cookie_dict = {}
print(cookies)
for cookie in cookies:
    if 'name' in cookie.keys() and 'value' in cookie.keys():
        cookie_dict[cookie['name']] = cookie['value']
print(cookie_dict)
with open('cookies.txt', 'w') as f:
    # 保存cookies到本地
    f.write(json.dumps(cookies))
    print("保存cookis成功.")
'''

# headers = {
#     'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.87 Safari/537.36'
# }
#
# param_other = {
#     'activity':'55678796',
#     'qid':'1'
# }
# post_other_url = 'https://www.wjx.cn/wjx/activitystat/viewjoinlist.aspx?'
#
# response = requests.get(url=post_other_url, params=param_other, headers=headers)
# page_text = response.text
# print(page_text)
# soup = BeautifulSoup(page_text, "html.parser")
# title = soup.find('div', id='divNotJoin').get_text()
# count = soup.find('div', id='divNotJoinCount').get_text()
# print(type(title))
# print(title)
# print(count)
# strlist = title.split(',')
# for i in strlist:
#     print(i)

# # 检查模拟登录后页面的 用户名 标签，若存在此标签则说明登录成功。
# login_check_xpath = '//*[@id="ctl01_lblUserName"]'
# login_check = driver.find_element_by_xpath(login_check_xpath)
# if(driver.find_element_by_xpath(login_check_xpath)):
#         print("Successful login in.")
#         html=driver.page_source #获取网页的html数据
#         # soup=BeautifulSoup(html,'lxml')#对html进行解析
#         with open("login_aft.html","w") as f:
#                 f.write(html)
# else:
#         print("Failed to login .")
# 最后不要忘记关闭driver


