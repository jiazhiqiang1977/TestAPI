#!/usr/bin/python
# -*- coding: UTF-8 -*-
# author: 赫本z
# 业务包：通用函数


import core.mysql as mysql
import core.log as log
import core.request as request
import core.excel as excel
import xlwings as xw
import constants as cs
from prettytable import PrettyTable
from bs4 import BeautifulSoup
from datetime import datetime

logging = log.get_logger()


class ApiTest:
    """接口测试业务类"""
    filename = cs.FILE_NAME

    def __init__(self):
        pass

    def prepare_data(self, host, user, password, db, sql):
        """数据准备，添加测试数据"""
        mysql.connect(host, user, password, db)
        res = mysql.execute(sql)
        mysql.close()
        logging.info("Run sql: the row number affected is %s", res)
        return res

    def get_excel_sheet(self, path, module):
        """依据模块名获取sheet"""
        excel.open_excel(path)
        return excel.get_sheet(module)

    def get_prepare_sql(self, sheet):
        """获取预执行SQL"""
        return excel.get_content(sheet, cs.SQL_ROW, cs.SQL_COL)

    def run_test(self):
        """再执行测试用例"""
        # rows = excel.get_rows(sheet)
        app = xw.App(visible=False)
        wb = app.books.open(cs.FILE_NAME)
        sht = wb.sheets[0]
        rng = xw.Range('A2')
        rows= rng.end('down').last_cell.row
        fail = 0
        for i in range(3, rows+1):
            testNumber = int(sht.range(cs.CASE_NUMBER+str(i)).value)
            testName = sht.range(cs.CASE_NAME+str(i)).value
            testUrl = sht.range(cs.CASE_URL+str(i)).value
            testMethod = sht.range(cs.CASE_METHOD+str(i)).value
            testHeaders = eval(sht.range(cs.CASE_HEADERS+str(i)).value)
            testData = sht.range(cs.CASE_DATA+str(i)).value
            expectCode = sht.range(cs.CASE_CODE_EXP+str(i)).options(numbers=int).value
            actualResponse = request.api(testMethod, testUrl, testData, testHeaders)
            actualCode = actualResponse.status_code
            sht.range(cs.CASE_CODE_ACT+str(i)).value = actualCode
            if actualCode != expectCode:
                sht.range(cs.CASE_CODE_JUD+str(i)).value = 'NG'
                logging.info("FailCase %s", testName)
                print("FailureInfo") 
                fail += 1
            else:
                sht.range(cs.CASE_CODE_JUD+str(i)).value = 'OK'
                logging.info("Number %s", testNumber)
                logging.info("TrueCase %s", testName)

            soup = BeautifulSoup(actualResponse.content,"xml")
            actRescode = soup.find('e-ML').ResCode.string
            sht.range(cs.CASE_RESCODE_ACT+str(i)).value = actRescode
            expRescode = sht.range(cs.CASE_RESCODE_EXP+str(i)).options(numbers=int).value
            if actRescode != str(expRescode) :
                sht.range(cs.CASE_RESCODE_JUD+str(i)).value = 'NG'
                logging.info("FailCase %s", testName)
                fail += 1
            else:
                sht.range(cs.CASE_RESCODE_JUD+str(i)).value = 'OK'
                logging.info("Number %s", testNumber)
                logging.info("TrueCase %s", testName)
            
            # failResults = PrettyTable(["Number", "Name", "ActualCode", "ExpectCode","ActualResCode","ExpectResCode"])
            # failResults.align["Number"] = "l"
            # failResults.padding_width = 1
            # failResults.add_row([testNumber, testName,  actualCode, expectCode,actRescode,expRescode])

        wb.save( datetime.now().strftime('%Y%m%d%H%M%S')+'TestResult.xlsx')
        wb.close()
        app.quit()

        if fail > 0:
            return False
        return True
