# -*- coding: utf-8 -*-
"""
Created on Mon Mar 26 03:37:28 2018

@author: User
"""

import ExcelFactory

class test:

    def __init__(self):
        self.prop_a = 1
        self.prop_b = 2
        self.prop_c = True
#        self.prop_d = 4

    def test_method():
        print("test")

ExcelDAO = ExcelFactory.SheetDAOFactory("test.xlsx")
#ExcelDAO.workbook["alpha"]["A1"] = 10
#sheetDAO = ExcelDAO.create_sheet_DAO("delta", ["prop_a", "prop_b", "prop_c"])
sheetDAO = ExcelDAO.get_sheet_DAO("delta")
print(sheetDAO.column_headers)
#print(sheetDAO.column_dict)
x = test()
row = x.__dict__
sheetDAO.create_row(row)
sheetDAO.create_row(row)
ExcelDAO.save()
#ExcelDAO_2 = ExcelFactory.SheetDAOFactory("test.xlsx")
#print(ExcelDAO.workbook["alpha"]["A1"].value)