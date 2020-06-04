# coding:utf-8
import test
import os

# version = '梯内屏售卖V3.0.6'
version = os.environ.get("version")
# sheets = '3'
sheets = os.environ.get("sheets")
# excel_file = u"测试用例_梯内屏售卖V3.0.4.xlsx"
excel_file = "excel_file"
ConvertTestCaseFromExcelToXml = test.ConvertTestCaseFromExcelToXml(excel_file=excel_file,
                                                                  version=version, sheets=sheets)
ConvertTestCaseFromExcelToXml.main()