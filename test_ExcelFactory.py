# -*- coding: utf-8 -*-
"""
Created on Sun Mar 25 03:33:12 2018

@author: User
"""

import unittest
import openpyxl
import ExcelFactory
import os

class TestExcelFactory(unittest.TestCase):
    
    def setUp(self):
        self.file_name = "test.xlsx"
        workbook = openpyxl.Workbook()
        title_array = ["alpha", "beta", "charlie"]
        for title in title_array:
            workbook.create_sheet(title=title)
        self.beta_headers_array = ["header1", "header2", "header3"]
        for x, header in enumerate(self.beta_headers_array):
            workbook["beta"].cell(column = x+1, row = 1).value = header
        workbook.save(self.file_name)
        self.ExcelDAO = ExcelFactory.SheetDAOFactory(self.file_name)
    
    def tearDown(self):
        os.remove(self.file_name)
    
    def test_ExcelFactory_init(self):
        self.assertIsInstance(ExcelFactory.SheetDAOFactory(self.file_name), ExcelFactory.SheetDAOFactory)
        self.assertIsInstance(ExcelFactory.SheetDAOFactory("xyz.xlsx"), ExcelFactory.SheetDAOFactory)
        
#        with self.assertRaises(FileNotFoundError):
#            ExcelFactory.SheetDAOFactory("xyz.xlsx")

    def test_get_sheet_DAO(self):
         self.assertIsInstance(self.ExcelDAO.get_sheet_DAO("beta"), ExcelFactory.SheetDAOImpl)
         self.assertEqual(self.ExcelDAO.get_sheet_DAO("beta").column_headers, self.beta_headers_array)
    
    def test_create_sheet_DAO(self):
        delta_headers_array = ["header1", "header2", "header3"]
        self.assertIsInstance(self.ExcelDAO.create_sheet_DAO("delta", delta_headers_array), ExcelFactory.SheetDAOImpl)
        self.assertEqual(self.ExcelDAO.create_sheet_DAO("delta", delta_headers_array).column_headers, delta_headers_array)
    
if __name__ == "__main__":
    unittest.main()