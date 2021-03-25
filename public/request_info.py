import requests
import time

from bs4 import BeautifulSoup

from Jira.public import logging

class Request:
    def __init__(self):
        self.headers = {}

    def set_headers(self, headers):
        self.headers.update(headers)

    def get_headers(self):
        return self.headers

    def request(self, url, data, is_get_method=True, verify=False, headers={}, response_type="json"):
        logging.info("Will Reuqeuest Interface: {}".format(url))
        logging.info("Augues as follows: \n{}".format(data))
        if not headers:
            headers = self.headers
        logging.debug("Headers as follows:\n{}".format(headers))
        if is_get_method:
            r = requests.get(url=url, headers=headers, verify=verify, params=data)
        else:
            r = requests.post(url=url, headers=headers, verify=verify, data=data)
        if response_type == "text":
            result = r.text
        elif response_type == "json":
            result = r.json()
        else:
            result = r.raw
        logging.debug("Response: \n{}".format(result))
        return result

    def get_support_portal(self):
        """请求接口，从返回的html文件中解析出需要的数据信息"""
        url = "https://ts.hundsun.com/se/portal/SupportPortal.htm"
        r_text = self.request(url=url, data=None, is_get_method=True, response_type="text")
        logging.debug(r_text)
        return self._extract_se_configurations(text=r_text)

    def _extract_se_configurations(self, text):
        """获取ts请求接口时，一些必要的参数信息，数据来源于接口 SupportPortal.htm"""
        se_configurations = {}
        import re
        user_id = re.search('se.SEConfig.UserId\t*=\s*"(\d+)";', text).group(1)
        token = re.search('se.SEConfig.token = "(\d+)";', text).group(1)
        pennanter = re.search('se.SEConfig.pennanter\s*=\s*"(\d+)";', text).group(1)
        # token = re.search('se.SEConfig.pennanter\s*=\s*"(\d+)";', text).group(1)
        arhinoceros = re.search('se.SEConfig.arhinoceros\s*=\s*"(\d+)";', text).group(1)
        current_group = re.search('se.SEConfig.currentGroup	\s*=\s*"(\d+)";', text).group(1)
        se_configurations.setdefault("user_id", user_id)
        se_configurations.setdefault("token", token)
        se_configurations.setdefault("pennanter", pennanter)
        se_configurations.setdefault("arhinoceros", arhinoceros)
        se_configurations.setdefault("current_group", current_group)
        for k, v in se_configurations.items():
            logging.info("Getted Argue {0}, the value is {1}".format(k, v))
        return se_configurations

    def set_se_configurations(self, se_configurations):
        self.se_configurations = se_configurations

    def get_se_configurations(self):
        return self.se_configurations

    def fetch_user_to_product_group_list(self, have_set_se_configurations=True, data={}):
        """通过ts接口，请求获取该用户的所有产品组权限"""
        import time

        url = "https://ts.hundsun.com/se/services/sys/fetchUser2productGroupList.htm?_dc={}"\
            .format(str(int(time.time() * 1000)))
        data = self.get_se_configurations()
        # se_configurations = self.get_se_configurations()
        # if se_configurations:
        #     data={
        #         "param.userId": se_configurations.get("user_id"),
        #         "currentGroup": se_configurations.get("current_group"),
        #         "token": se_configurations.get("token"),
        #         # "token": "1615089590850",
        #         "arhinoceros": se_configurations.get("arhinoceros"),
        #         "pennanter": se_configurations.get("pennanter")
        #           }
        # self.headers.update({"Cookie": "JSESSIONID=4DFCA95C07AFF5A4C493AFAAC766D8CB"})
        self.headers.update({"Referer": "https://ts.hundsun.com/se/portal/SupportPortal.htm"})
        self.headers.update({"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        time.sleep(3)
        r_json = self.request(url=url, is_get_method=False, data=data, headers=self.headers)
        logging.debug(r_json)
        return r_json

    def get_data_from_se_configurations(self):
        se_configurations = self.get_se_configurations()
        if se_configurations:
            data = {
                "param.userId": se_configurations.get("user_id"),
                "currentGroup": se_configurations.get("current_group"),
                "token": se_configurations.get("token"),
                "arhinoceros": se_configurations.get("arhinoceros"),
                "pennanter": se_configurations.get("pennanter")
            }
        return data

    def fetch_ts_issues(self, data={}):
        url = "https://ts.hundsun.com/se/services/modify/fetchModifyByUserdefinedModluePaginatedLucene.htm?\
        _dc={}".format(int(time.time()*1000))
        full_data = self.get_se_configurations()
        full_data.update(data)
        return self.request(url=url, is_get_method=False, headers=self.headers, data=full_data)

    def filter_html_label(self, data_list):
        """去除列表数据元素中的html标签
        <div>【任务类型】太平洋</div><div>【任务方案】</div 
         -->
        【任务类型】太平洋
        【任务方案】
        """
        data_list_dealed = []
        for data in data_list:
            data_tmp = ""
            for each in BeautifulSoup(data).string:
                data_tmp += each + "\n"
            data_list_dealed.append(data_tmp)
        return data_list_dealed

    def get_jira_column_relation(self):
        """通过调用此接口，获取jira的字段关系信息(jira新增缺陷界面看到的字段与实际传入字段的名称对应关系)"""
        return self.request(url="https://se.hundsun.com/secure/QuickCreateIssue!default.jspa?decorator=none",
                      is_get_method=False)

    def get_issue_table(self, jql):
        url = "https://se.hundsun.com/rest/issueNav/1/issueTable"
        data = {"startIndex": 0,
                "jql": jql,
                "layoutKey": "list-view"}
        try:
            r = self.request(url=url, is_get_method=False, data=data)
            table_info = r.get("issueTable").get("table")
        except AttributeError as e:
            logging.error("Get Issue Table Info Error: {0}\n".format(r))
        except Exception as e:
            logging.error("Get Issue Table Info Beyond AttributeError: {0}".format(str(e)))
            raise Exception("Something wrong.")
        return table_info


if __name__ == "__main__":
    from Jira.public.login import Login
    from Jira.public.loading_config import LoadingConfig
    loading_config = LoadingConfig()
    config_obj_of_authority = loading_config.loading(filename="../config/authority.conf")

    login = Login(config_obj_of_authority=config_obj_of_authority)
    login.phatomjs_login()
    valid_cookie = login.get_valid_cookie()

    request = Request()
    jql = "project = ISEE AND issuetype = Bug and creator = currentUser()  ORDER BY updated DESC"
    headers = {"Cookie": valid_cookie,
               "X-Atlassian-Token": "no-check"}
    request.set_headers(headers)
    table_info = request.get_issue_table(jql="project = ISEE AND issuetype = Bug and creator = currentUser()  ORDER BY updated DESC")

