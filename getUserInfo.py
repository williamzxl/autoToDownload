#!/usr/bin/env python3
# -*- coding:utf-8 -*-

import xlrd

class XlUserInfo(object):
    def __init__(self,path=''):
        self.path = path
        self.xl = xlrd.open_workbook(self.path)

    def get_sheet_info(self):
        all_info = []
        info0 = []
        info1 = []
        for row in range(0,self.sheet.nrows):
            info = self.sheet.row_values(row)
            info0.append(info[0])
            info1.append(info[1])
        temp = zip(info0,info1)
        all_info.append(dict(temp))
        return all_info.pop(0)

    def get_sheetinfo_by_name(self,name):
        self.name = name
        self.sheet = self.xl.sheet_by_name(self.name)
        return self.get_sheet_info()

if __name__ == '__main__':
    xl = XlUserInfo('userInfo.xlsx')
    userinfo = xl.get_sheetinfo_by_name('userInfo')
    webinfo = xl.get_sheetinfo_by_name('WebEle')
    print(userinfo)
    print(webinfo)