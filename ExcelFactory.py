# -*- coding: utf-8 -*-
"""
Created on Sat Mar 10 17:45:32 2018

@author: User
"""

import openpyxl

class SheetDAOFactory:
    
    def __init__(self, file_name):
        self.file_name = file_name
        try:
            self.workbook = openpyxl.load_workbook(file_name)
            self.sheet_list = self.workbook.get_sheet_names()
        except FileNotFoundError:
            self.workbbook = openpyxl.Workbook()
        
    def save(self):
        self.workbook.save(self.file_name)
    
    def create_sheet_DAO(self, sheet_name, *headers):
        pass
    
    def get_sheet_DAO(self, sheet_name):
        pass
    
class SheetDAOImpl:
    
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.column_headers = []
        for column in self.worksheet.itercols():
            self.column_headers.append(column[0].value)
    
    def create_row(self, row):
        pass
    
    def find_row(self, row):
        pass
    
    def update_row(self, row):
        pass
    
    def delete_row(self, row):
        pass