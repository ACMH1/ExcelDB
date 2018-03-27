# -*- coding: utf-8 -*-
"""
Created on Mon Mar 26 03:37:28 2018

@author: User
"""

import ExcelFactory

ExcelDAO = ExcelFactory.SheetDAOFactory("test.xlsx")
ExcelDAO.workbook["alpha"]["A1"] = 10
sheetDAO = ExcelDAO.create_sheet_DAO("delta", ["a", "b", "c"])
ExcelDAO.save()
ExcelDAO_2 = ExcelFactory.SheetDAOFactory("test.xlsx")
print(ExcelDAO.workbook["alpha"]["A1"].value)