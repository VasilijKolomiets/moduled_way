# -*- coding: utf-8 -*-
"""
Created on Mon Mar 29 18:52:53 2021

@author: Vasil
"""

from pathlib import Path

to_del = Path("del_me.txt")
to_del.unlink()
