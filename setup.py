# -*- coding: gb18030 -*-
'''
    Created on 2016-01-11

    @author: Gavin.Bai
    @note: Setup program to convert the *.py to exe to run
    @version: v1.0
    @Modify:
    @License: (C)GPL
'''
from distutils.core import setup
import py2exe

options = {
           "py2exe" :
           {
            "compressed"   : 1,
            "bundle_files" : 1
            }
          }

setup(options = options,
      zipfile = None,
      console = ["main_gui.py"]
      )