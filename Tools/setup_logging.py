# -*- coding: utf-8 -*-
"""
Created on Mon Apr 19 17:52:45 2021

@author: luzziem

"""

import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("Logger")

handler = logging.FileHandler('logfile.log')
handler.setLevel(logging.INFO)

# create a logging format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# add the file handler to the logger
logger.addHandler(handler)

