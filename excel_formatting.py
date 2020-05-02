# -*- coding: utf-8 -*-
"""
Created on Thu Apr 30 11:47:58 2020

@author: Vasil
"""
import win32com.client

# In[]:


def search_and_colorise(work_sheet, searched_texts_list, color_num=4):
    if type(searched_texts_list) is str:
        raise Exception('list of str expected!')
    for seached in searched_texts_list:
        work_sheet.Cells.Find(seached).Interior.ColorIndex = color_num


def excel_file_formatting(file_to_format):
    
    Excel = win32com.client.Dispatch("Excel.Application")
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

