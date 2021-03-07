# -*- coding: utf-8 -*-
# @Time    : 2021/3/3 10:18
# @Author  : chenkang19736
# @File    : deal_data.py
# @Software: PyCharm
import re

from bs4 import BeautifulSoup

from Jira.public import logging
from Jira.public.excel_operation import WriteToExcel

class DealData:
    """"""
    def __init__(self, content={"isfile": False, "content": ""}):
        if not content.get("isfile"):
            self.html = content.get("content")
            self.soup = BeautifulSoup(self.html, features="html.parser")
        else:
            self.soup = BeautifulSoup(open(file=content.get("content"), mode="r", encoding="utf-8"),
                                      features="html.parser")
        logging.debug(self.soup.prettify())

    def set_html(self, html):
        self.html = html

    def get_html(self):
        return self.html

    def data_to_html(self):
        with open(file="jira_table.html", mode="r", encoding="utf-8") as f:
            line = f.readline()
            # print(line)
            soup = BeautifulSoup(line, features="html.parser")

            with open("bs.html", mode="w", encoding="utf-8") as fw:
                fw.write(soup.prettify())

        self.set_html(soup.prettify())

    def get_table_data_from_html(self):
        """
        通过jira的接口返回的字段，使用BeautifulSoup进行封装，获取缺陷列表的表格头名称以及表格的数据
        :return: 返回表格头名称字典项_head_dic 与 表格数据字典项 _column_dic
        """
        table_data_dic = {}
        _head_dic = {}
        for i, th in enumerate(self.soup.thead.find_all("th")):
            try:
                column = th.span.string
                _head_dic.update({i: column})
                logging.info("The {1} Table Head Is [{0}]".format(column, i))
            except AttributeError as e:
                logging.warning("Get The Table Head occurs AttributeError: {}".format(str(e)))

        # print(self.soup.tbody.prettify())
        for i, tr in enumerate(self.soup.tbody.find_all("tr")):
            data_list = self.__get_value_from_tr(tr, column_relation=_head_dic)
            table_data_dic.update({i: data_list})
            # if not table_dic.get(i):
            #     table_dic.setdefault(i, [])
            # try:
            #
            #     table_dic[i].append(tr.td.string)
            #     logging.info("The Table Cell Is [{0}]".format(tr.td.a.string))
            # except AttributeError as e:
            #     logging.warning("Get The Table Head occurs AttributeError: {}".format(str(e)))

        return _head_dic, table_data_dic

    def __get_value_from_tr(self, tr, column_relation):
        """
        解析tbody.tr数据， 获取行数据的内容，以列表的形式返回
        :param tr:  html 表格中 tr的标签信息， tr是以 beautifulsoup 中的 Tag标签类型传入
        :param column_relation:  是一个类似{0: 缺陷， 1:优先级}的字典，key为表格的第几列信息，value为表格对应列表头名称
        :return: 返回一个列表，列表信息为该行数据各个字段的内容
        """

        tr_value = []
        for i, td in enumerate(tr.find_all("td")):
            if column_relation.get(i) == "T":
                content = "缺陷"
            elif column_relation.get(i) == "优先级":
                content = "优先级"
            elif column_relation.get(i) == "修改单号":
                content = "无"
            elif column_relation.get(i) == "打回次数":
                content = "0"
            elif column_relation.get(i) is None:
                continue
            else:
                # content = re.search("\s*(\S*)\s*", td.get_text(strip=True)).group(1)
                content = td.get_text(strip=True)
            try:
                tr_value.append(content)
                logging.info("Target Getted the {0} td -- [{1}]".format(i, content))
            except AttributeError as e:
                logging.warning("Its the {0} td: [{1}]".format(i, td))
        return tr_value


if __name__ == "__main__":
    write_to_excel = WriteToExcel(filename="bug_list.xls", sheetname="缺陷列表")

    jira = DealData(content={"isfile": True, "content": "jira_table.html"})
    header, data = jira.get_table_data_from_html()
    write_to_excel.write_via_row(data_list=header.values(), startrow=0, startcol=0)
    write_to_excel.write_via_row(data_list=data.values(), startrow=1, startcol=0)
    write_to_excel.close_file()
