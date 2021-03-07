# -*- coding: utf-8 -*-
# @Time    : 2021/3/4 19:28
# @Author  : chenkang19736
# @File    : run.py
# @Software: PyCharm
import os
import requests

from bs4 import BeautifulSoup

from TS.public.login import Login
from TS.public.request_info import Request
from TS.public.excel_operation import WriteToExcel
from TS.public import logging
from TS.public.mail import Mail


# 获取配置文件信息
from TS import config_obj_of_authority
logging.info("Have getted configuration of authority.")

# 登录获取cookie等凭证
login = Login(url=config_obj_of_authority['ts']['url'], config_obj_of_authority=config_obj_of_authority)
login.phatomjs_login(executable_path="./config/chromedriver.exe")
valid_ts_cookie = login.get_valid_ts_cookie()
logging.info("Have getted login authority of HS plateform.")

# 获取ts数据信息列表
request = Request()
request.set_headers(headers={"Cookie": valid_ts_cookie})
support_portal = request.get_support_portal()
request.set_se_configurations(se_configurations=support_portal)
user_to_product_group_list = request.fetch_user_to_product_group_list()
logging.info("Have getted TS bug list from TS plateform.")
login.quit()

# 获取ts指定版本数据
data = {"param": '{"modifyStatus":"0,1,2,3,4,5,6,7,8,9,11,12","versionNo":"%FSP1.0V202101.02.000%"}',
        "start": 0, "limit": 200, "isUserDataValidity": "Y", "page": 1}
ts_data_list = request.fetch_ts_issues(data=data)

write_to_excel = WriteToExcel(filename="./static/ts导出表.xls", sheetname="ts_list")
write_to_excel.write_via_row(data_list=["修改单单号", "修改理由", "集成说明"], startcol=0, startrow=0)
write_to_excel.write_via_row(data_list=ts_data_list.get("resultBOList"),
                             keyname=["modifyNum", "modifyReason", "integationDesc"], startcol=0, startrow=1)
write_to_excel.close_file()
logging.info("Have written ts data to ./static/ts导出表.xls")

# 将缺陷数据写入excel文件
# write_to_excel = WriteToExcel(filename="./static/bug_list.xls", sheetname="缺陷列表")
# TS = TS(content={"isfile": False, "content": table_info})
# header, data = TS.get_table_data_from_html()
# write_to_excel.write_via_row(data_list=header.values(), startrow=0, startcol=0)
# write_to_excel.write_via_row(data_list=data.values(), startrow=1, startcol=0)
# write_to_excel.close_file()
# logging.info("Have written infomation to excel file.")

# 发送邮件
# sender = config_obj_of_authority['email']['username'].strip()
# password = config_obj_of_authority['email']['password'].strip()
# recipients = recipient = config_obj_of_authority['email']['recipient'].strip()
#
# mail = Mail(sender=sender, password=password)
# if "," in recipient:
#     recipients = recipient.split(",")
