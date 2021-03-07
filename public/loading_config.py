import os

import configparser

from Jira.public import logging

class LoadingConfig:
    """读取配置文件"""
    def __init__(self):
        pass

    def loading(self, filename):
        try:
            cf = configparser.ConfigParser()
            cf.read(filenames=filename, encoding="utf8")
        except FileNotFoundError:
            logging.error("The File {0} Was Not Found".format(filename))
        return cf

