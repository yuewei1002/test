# coding:utf-8

import xlrd
import sys, os
import tarfile

# if os.environ.get("WORKSPACE"):
#     sys.path.append(os.path.join(os.environ.get("WORKSPACE"), "AutoTest"))
# from ..Others import get_python_version

# if get_python_version() == 3:
from functools import reduce
# from .. import logger

# logging = logger
from files import HandleXml


class ConvertTestCaseFromExcelToXml(object):
    def __init__(self, excel_file,version,sheets):
        self.excel_file = excel_file
        self.version = version
        self.dic_testlink = {}
        self.sheets = map(lambda x: int(x.strip()), sheets.split(","))
        self.xml_tool = HandleXml()

    def get_sheet(self, sheet_index):
        sheet_obj = xlrd.open_workbook(self.excel_file).sheet_by_index(sheet_index)
        return sheet_obj

    def case_to_dic(self, case, sheet_name):
        testcase = {"testsuite": "", "name": "", "node_order": "100", "externalid": "", "version": "1", "summary": "",
                    "preconditions": "", "execution_type": "1", "importance": "3", "steps": [], "keywords": "P1"}
        # if get_python_version() == 2:
        #     testcase["testsuite"] = case[1].encode("utf-8")
        #     testcase["name"] = case[3].encode("utf-8")
        # else:
        testcase["testsuite"] = case[1]
        testcase["name"] = case[3]
        testcase["summary"] = ""
        try:
            testcase["importance"] = str(case[4])
        except:
            print(u"错误的用例优先级格式，请查看用例：%s, 对应的优先级为：%s" % (testcase["name"], case[4]))
            print(u"sheet名称为：%s" % sheet_name)
            sys.exit(1)
        # if get_python_version() == 2:
        #     testcase["preconditions"] = case[5].encode("utf-8")
        #     testcase["keywords"] = case[2].encode("utf-8")
        #     actions = case[6].strip().encode("utf-8").split("\n")
        #     expectedresults = case[7].strip().encode("utf-8").split("\n")
        # else:
        testcase["preconditions"] = case[5]
        testcase["keywords"] = case[2]
        actions = case[6].strip().split("\n")  # 去掉空格，通过换行符分隔成列表
        expectedresults = case[7].strip().split("\n")  # 去掉空格，通过换行符分隔成列表
        count = 0
        if len(actions) == len(expectedresults):
            while count < len(actions):
                step = {"execution_type": ""}
                step["step_number"] = str(count + 1)
                step["actions"] = actions[count]
                step["expectedresults"] = expectedresults[count]
                testcase["steps"].append(
                    step)  # {"execution_type": "","step_number":"1","actions":"步骤1","expectedresults":"期望1"}
                count += 1
        else:
            step = {"execution_type": ""}
            step["step_number"] = "1"
            step["actions"] = reduce(lambda x, y: x + y, actions)
            step["expectedresults"] = reduce(lambda x, y: x + y, expectedresults)
            testcase["steps"].append(step)  # {"execution_type": "","step_number":"","actions":"","expectedresults":""}
        return testcase

    def get_cases(self, case_sheet_obj):
        start_row = 4
        cases = []
        while True:
            try:
                case = case_sheet_obj.row_values(start_row)  # 返回由该行中所有单元格的数据组成的列表
                cases.append(case)
            except IndexError:
                break
            start_row += 1
        return cases

    def zip_file(self, zip_file, file_list):
        tar = tarfile.open(zip_file, 'w')
        for File in file_list:
            tar.add(File, arcname=File.split(os.sep)[-1])
        tar.close()

    def main(self):
        file_list = []
        for sheet in self.sheets:  # 单张表操作
            sheet_obj = self.get_sheet(sheet - 1)
            cases = self.get_cases(case_sheet_obj=sheet_obj)  # 获取单张表用例，列表嵌套列表
            split_number = 300  # testlink有导入大小限制，所以做拆分处理
            cases_splited = [cases[i:i + split_number] for i in range(0, len(cases), split_number)]  # 每300条用例拆分为一个子列表
            cases_splited_count = 0
            while cases_splited_count < len(cases_splited):
                for case in cases_splited[cases_splited_count]:
                    case_datails = self.case_to_dic(case, sheet_obj.name)  # 返回用例的字典格式
                    test_suite = case_datails["testsuite"].replace("\n", "")  # 对应二级模块名 兼容回车符问题
                    if not test_suite:
                        print("suite is empty for cases:%s" % case_datails)
                        sys.exit("suite is empty for cases:%s" % case_datails)
                    if sheet_obj.name not in self.dic_testlink.keys():
                        self.dic_testlink[sheet_obj.name] = {}  # 添加表名的空字典，字典嵌套字典，一张表对应一个子字典
                    if test_suite not in self.dic_testlink[sheet_obj.name].keys():
                        self.dic_testlink[sheet_obj.name][test_suite] = {"testcase": [case_datails]}
                    else:
                        self.dic_testlink[sheet_obj.name][test_suite]["testcase"].append(case_datails)
                        # {"表名":{"二级模块1":{"测试用例":[{用例1}，{用例2}]},"二级模块2":{"测试用例":[{用例1}，{用例2}]}...}}

                node_version = self.xml_tool.add_node("testsuite", {"name": self.version})  # 根节点，版本
                # node_order
                self.xml_tool.add_child_node(node_version, "node_order", content="4")
                # details
                self.xml_tool.add_child_node(node_version, "details")

                node_sheet = self.xml_tool.add_child_node(node_version, "testsuite", {"name": sheet_obj.name})  # 表名节点
                # node_order
                self.xml_tool.add_child_node(node_sheet, "node_order", content="4")
                # details
                self.xml_tool.add_child_node(node_sheet, "details")

                for sheet_name_from_dict in self.dic_testlink.keys():
                    if sheet_name_from_dict == sheet_obj.name:
                        for test_suite in self.dic_testlink[sheet_name_from_dict].keys():
                            node_modlue = self.xml_tool.add_child_node(node_sheet, "testsuite",
                                                                       {"name": test_suite})
                            testcase_list = self.dic_testlink[sheet_name_from_dict][test_suite]["testcase"]  # 列表格式
                            # print(testcase_list)
                            for testcase in testcase_list:
                                # testcase
                                node_case = self.xml_tool.add_child_node(node_modlue, "testcase",
                                                                         {"name": testcase["name"]})
                                # node_order
                                self.xml_tool.add_child_node(node_case, "node_order", content=testcase["node_order"])
                                # externalid
                                self.xml_tool.add_child_node(node_case, "externalid", content=testcase["externalid"])
                                # version
                                self.xml_tool.add_child_node(node_case, "version", content=testcase["version"])
                                # summary
                                self.xml_tool.add_child_node(node_case, "summary", content=testcase["summary"])
                                # preconditions
                                self.xml_tool.add_child_node(node_case, "preconditions",
                                                             content=testcase["preconditions"])
                                # execution_type
                                self.xml_tool.add_child_node(node_case, "execution_type",
                                                             content=testcase["execution_type"])
                                # importance
                                self.xml_tool.add_child_node(node_case, "importance", content=testcase["importance"])
                                # steps
                                node_steps = self.xml_tool.add_child_node(node_case, "steps")
                                for step in testcase["steps"]:
                                    # step
                                    node_step = self.xml_tool.add_child_node(node_steps, "step")
                                    # step_number
                                    self.xml_tool.add_child_node(node_step, "step_number", content=step["step_number"])
                                    # actions
                                    self.xml_tool.add_child_node(node_step, "actions", content=step["actions"])
                                    # expectedresults
                                    self.xml_tool.add_child_node(node_step, "expectedresults",
                                                                 content=step["expectedresults"])
                                    # execution_type
                                    self.xml_tool.add_child_node(node_step, "execution_type",
                                                                 content=step["execution_type"])
                                # keywords
                                node_keywords = self.xml_tool.add_child_node(node_case, "keywords")
                                # keyword
                                node_keyword = self.xml_tool.add_child_node(node_keywords, "keyword",
                                                                            {"name": testcase["keywords"]})
                                # notes
                                self.xml_tool.add_child_node(node_keyword, "notes", content="aaaa")
                        # if get_python_version() == 3:
                        xml_file = "%s_%s_%s.xml" % (self.version, sheet_obj.name, str(cases_splited_count + 1))
                        # else:
                        #     xml_file = "%s_%s_%s.xml" % (
                        #         self.version, sheet_obj.name.encode("utf-8"), str(cases_splited_count + 1))
                        self.xml_tool.write_xml(xml_file, root_node=node_version)
                        cases_splited_count += 1
                        file_list.append(xml_file)
        tar_file = self.excel_file.replace(".xlsx", "").replace(".xls", "") + ".tar"
        self.zip_file(tar_file, file_list)
        return tar_file

#
# if __name__ == '__main__':
#     version = '瑕疵1.1.0'
#     sheets = '3'
#     excel_file = "F:\\瑕疵\\01.瑕疵检测平台_V1.0.0_功能测试用例.xlsx"
#     ConvertXml = ConvertTestCaseFromExcelToXml(excel_file=excel_file,version=version,sheets=sheets)
#     ConvertXml.main()
