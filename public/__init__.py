# -*- coding: utf-8 -*-
# @Time    : 2021/3/2 13:38
# @Author  : chenkang19736
# @File    : __init__.py.py
# @Software: PyCharm

#!encoding:utf-8
import json, time
import logging.config
import os
import logging

def setup_logging(
    default_path='./config/logging.json',
    default_level=logging.INFO,
    env_key='LOG_CFG'
):
    """Setup logging configuration

    """
    path = default_path
    value = os.getenv(env_key, None)
    if value:
        path = value
    if os.path.exists(path):
        with open(path, 'rt') as f:
            config = json.load(f)
        logging.config.dictConfig(config)
    else:
        logging.basicConfig(level=default_level)
    return logging


setup_logging()