# -*- coding: utf-8 -*-
"""
Created on Thu Feb 10 12:38:42 2022

@author: Vasil

!!!
================
shiny  - R tdatatable --> WEB
monet  - аналитическая база данных (столбцовая БД) + monet_lite
clickhose  - аналитическая база данных (столбцовая БД)


PostGressSQL  -  тоже хранит элементы в столбцовых БД - есть зачатки
================
"""

import csv


def find_delimiter(filename):
    sniffer = csv.Sniffer()
    with open(filename) as fp:
        delimiter = sniffer.sniff(fp.read(50000), delimiters=(',', '\t')).delimiter
    return delimiter


file01 = r"D:\Downloads\0658_Order discrepansy\0658_Event Detail_763305019030.csv"
file02 = r"D:\Downloads\0658_Order discrepansy\0658_Event Detail_763305019030___.csv"
file03 = r"D:\Downloads\0658_Order discrepansy\0658_customer returns_763569019031.csv"

file04 = r"D:\_\BI\Esko\ESKO_Date_Range_2021Feb7-2022Feb6CustomUnifiedTransaction.csv"

delimiter = find_delimiter(file03)

print(F"Sniffer detect delimiter simbol: '{delimiter}' with ord={ord(delimiter)}")
