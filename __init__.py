# -*- coding: utf-8 -*-
# @Time    : 2021/3/7 16:08
# @Author  : chenkang19736
# @File    : __init__.py.py
# @Software: PyCharm

import os

from Jira.public.loading_config import LoadingConfig
ROOT_PATH = os.path.abspath(".")

loading_config = LoadingConfig()
config_obj_of_jira_relation = loading_config.loading(filename="config/jira.conf")
config_obj_of_authority = loading_config.loading(filename="config/authority.conf")
