import time

from Jira.public import logging

class Login:
    def __init__(self, url, config_obj_of_authority):
        self.login_url = url
        self.authority = config_obj_of_authority
        self.jira_cookie = ""
        self.ts_cookie = ""
        self.driver = None

    def phatomjs_login(self, executable_path="../config/chromedriver.exe"):
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        # driver = webdriver.Chrome()
        # 模拟无界面谷歌浏览器操作
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(executable_path=executable_path, options=chrome_options)
        driver.get(url=self.login_url)
        time.sleep(1)
        # 登录
        driver.find_element_by_xpath(xpath='//*[@id="username"]').send_keys(self.authority['jira']['username'])
        driver.find_element_by_xpath(xpath='//*[@id="password"]').send_keys(self.authority['jira']['password'])
        driver.find_element_by_xpath(xpath='//*[@id="submit_psd"]/input').click()
        time.sleep(0.5)
        # 获取登录后token信息
        for cookie in driver.get_cookies():
            logging.info("Cookie From Login Is: {}".format(cookie))
        token = driver.get_cookie(name="atlassian.xsrf.token")
        logging.info("atlassian.xsrf.token info is [{}]".format(token))
        jessionid = driver.get_cookie(name="JSESSIONID")
        logging.info("JSESSIONID info is [{}]".format(jessionid))
        if "se.hundsun.com" in self.login_url:
            self.jira_cookie += token.get("name") + "=" + token.get("value") + "; jira.editor.user.mode=wysiwyg;"
            self.jira_cookie += jessionid.get("name") + "=" + jessionid.get("value")
            logging.info("The valid cookie is [{}]".format(self.jira_cookie))
        elif "ts.hundsun.com" in self.login_url:
            self.ts_cookie += jessionid.get("name") + "=" + jessionid.get("value")
            logging.info("The valid cookie is [{}]".format(self.ts_cookie))
        else:
            logging.warning("The request url [{}] is not dealed correctly.".format(self.login_url))
        self.driver = driver
        return driver

    def quit(self):
        self.driver.quit()

    def get_valid_jira_cookie(self):
        return self.jira_cookie

    def get_valid_ts_cookie(self):
        return self.ts_cookie

if __name__ == "__main__":
    from Jira.public.loading_config import LoadingConfig

    loading_config = LoadingConfig()
    config_obj_of_authority = loading_config.loading(filename="../config/authority.conf")
    login = Login(url=config_obj_of_authority['ts']['url'], config_obj_of_authority=config_obj_of_authority)
    login.phatomjs_login()
    valid_jira_cookie = login.get_valid_jira_cookie()
    valid_ts_cookie = login.get_valid_ts_cookie()
    import requests
    # r = requests.get(url="https://ts.hundsun.com/se/portal/SupportPortal.htm",
    #               headers={"Cookie": "modifyPageSize=100; JSESSIONID=4DFCA95C07AFF5A4C493AFAAC766D8CB"})
    r = requests.get(url="https://ts.hundsun.com/se/portal/SupportPortal.htm",
                  headers={"Cookie": valid_ts_cookie})
    logging.info(r.text)