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
            self.wb = openpyxl.load_workbook(file_name)
        except FileNotFoundError:
            self.wb = openpyxl.Workbook()
        
    def save(self):
        self.wb.save(self.file_name)
    
    def create_sheet_DAO(self, sheet_name, *headers):
        pass
    
    def get_sheet_DAO(self, sheet_name):
        pass