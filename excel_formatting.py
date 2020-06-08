# -*- coding: utf-8 -*-
"""
Created on Thu Apr 30 11:47:58 2020

@author: Vasil
"""
import os
import sys

current_os_is_win = sys.platform in ['win32', 'cygwin'] and os.name == "nt"
if current_os_is_win:
    import win32com.client

# In[]:

def search_and_colorise(work_sheet, searched_texts_list, color_num=4):
    if type(searched_texts_list) is str:
        raise Exception('list of str expected!')
    for seached in searched_texts_list:
        work_sheet.Cells.Find(seached).Interior.ColorIndex = color_num


def excel_file_formatting(file_to_format):

    if not current_os_is_win:
        o_excel_file_formatting(file_to_format)
        return
    
    Excel = win32com.client.DispatchEx("Excel.Application")
    wb = Excel.Workbooks.Open(file_to_format)
    ws_rec = wb.Worksheets("Sheet01")

    search_and_colorise(ws_rec, ("01_", '10_', '20_'))

    ws_rec.Rows('1:1').WrapText = True
    ws_rec.Columns('B:AZ').ColumnWidth = 9.71
    ws_rec.Rows("1:1").EntireRow.AutoFit()

    ws_rec.Range("B2").Select()
    Excel.ActiveWindow.FreezePanes = True

    ws_pivot_adj = wb.Worksheets("Pivot_ADJ")
    ws_pivot_adj.Activate()

    search_and_colorise(ws_pivot_adj, ["Z__SELLABLE"])

    ws_pivot_adj.Columns("D:E").NumberFormat = "# ##0"

    ws_pivot_adj.Columns('A:Z').AutoFit()
    ws_pivot_adj.Range('A1:Z1').WrapText = True
    ws_pivot_adj.Columns('G:Z').ColumnWidth = 13.14
    ws_pivot_adj.Columns('C:C').ColumnWidth = 11.57
    ws_pivot_adj.Rows("1:1").EntireRow.AutoFit()
    ws_pivot_adj.Range("D2").Select()
    Excel.ActiveWindow.FreezePanes = True

    wb.Save()
    wb.Close()
    Excel.Application.Quit()
    Excel.Quit()
    

# %%  24
from pathlib import Path
import pandas as pd
import numpy as np
import re

import openpyxl as opx
from openpyxl.styles import Color, PatternFill, Font, Border


def o_search_and_colorise(work_sheet, searched_texts_list, color='EE1111'):
    if type(searched_texts_list) is str:
        raise Exception('list of str expected!')
    # openpx style
    fill_color = PatternFill(start_color=color,
                             end_color=color,
                             fill_type='solid')
    rows_, cols_ = work_sheet.max_row, work_sheet.max_column
    for searched in searched_texts_list:
        for j in range(1, cols_ + 1):
            print(work_sheet.cell(row=1, column=j).value)
            if work_sheet.cell(row=1, column=j).value:
                if work_sheet.cell(row=1, column=j).value.find(searched)>=0:
                    work_sheet.cell(row=1, column=j).fill = fill_color


def o_load(file_xl):
        import datetime
        print(datetime.datetime.now())
        wb = opx.load_workbook(file_xl)
        print(datetime.datetime.now())
        return wb


def o_save(wb, file_xl):
    import datetime
    print(datetime.datetime.now())
    wb.save(file_xl)
    print(datetime.datetime.now())
    #wb.close()
    return None


def o_excel_file_formatting(file_to_format):
    
    fill_color = "CCFFFF"
    
    wb = o_load(file_to_format)
    ws_rec = wb["Sheet01"]
    # ws_rec.Range("B2").Select()
    # Excel.ActiveWindow.FreezePanes = True
    ws_rec.freeze_panes = ws_rec["B2"]
    rows_, cols_ = ws_rec.max_row, ws_rec.max_column
    o_search_and_colorise(ws_rec, ("01_", '10_', '20_'), color=fill_color)
    '''
    ws_rec.Rows('1:1').WrapText = True
    ws_rec.Columns('B:AZ').ColumnWidth = 9.71
    ws_rec.Rows("1:1").EntireRow.AutoFit()
    '''
    def wrap_and_size(ws, cols_, size):
        for j in range(1, cols_ + 1):
            ws.cell(row=1, column=j).alignment = \
                    opx.styles.Alignment(
                        horizontal='general',
                        vertical='bottom',
                        text_rotation=0,
                        wrap_text=True,
                        shrink_to_fit=False,
                        indent=0)
            ws.column_dimensions[opx.utils.cell.get_column_letter(j)].width = size
    
    wrap_and_size(ws_rec, cols_, 9.71)
    ws_rec.column_dimensions["A"].width = 13.71
    ws_rec.row_dimensions[1].height = 30

    ws_pivot_adj = wb["Pivot_ADJ"]
    rows_, cols_ = ws_pivot_adj.max_row, ws_pivot_adj.max_column
    ws_pivot_adj.freeze_panes = ws_pivot_adj["D2"]
    o_search_and_colorise(ws_pivot_adj, ["Z__SELLABLE"])
    
    #  ws_pivot_adj.Columns("D:E").NumberFormat = "# ##0"
    D = opx.utils.cell.column_index_from_string("D")
    E = opx.utils.cell.column_index_from_string("E")
    for i in range(2,rows_+1):
        ws_pivot_adj.cell(row=i, column=D).number_format = '# ### ##0'
        ws_pivot_adj.cell(row=i, column=E).number_format = '# ### ##0'
    # ws_pivot_adj.Columns('G:Z').ColumnWidth = 13.14
    # ws_pivot_adj.Range('A1:Z1').WrapText = True
    wrap_and_size(ws_pivot_adj, cols_, 13.14)
    '''
    ws_pivot_adj.Columns('A:Z').AutoFit()        
    ws_pivot_adj.Columns('C:C').ColumnWidth = 11.57
    ws_pivot_adj.Rows("1:1").EntireRow.AutoFit()
    '''
    ws_pivot_adj.column_dimensions["A"].width = 13.71
    ws_pivot_adj.column_dimensions["B"].width = 23.86
    ws_pivot_adj.column_dimensions["C"].width = 11.57
    ws_pivot_adj.column_dimensions["D"].width = 18.71
    ws_pivot_adj.column_dimensions["E"].width = 18.71
    ws_pivot_adj.column_dimensions["F"].width = 6.37    
    
    ws_pivot_adj.row_dimensions[1].height = 30    

    o_save(wb, file_to_format)
