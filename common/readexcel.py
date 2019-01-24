import xlrd
import json


class ExcelUtil():
    def __init__(self, excel_path, sheet_name='Sheet'):
        self.data = xlrd.open_workbook(excel_path)
        self.table = self.data.sheet_by_name(sheet_name)
        # 获取第一行作为key值
        self.keys = self.table.row_values(0)
        # 获取第一列作为key值
        # self.keys = self.table.col_values(0)
        # 获取总行数
        self.rowNum = self.table.nrows
        # 获取总列数
        self.colNum = self.table.ncols

    def dict_data(self):
        if self.colNum <= 1:
            print("总列数小于1")
        else:
            r = []
            j = 1
            for i in range(self.rowNum-1):
                s = {}
                # 从第二列取对应的values值
                s['rowNum'] = i + 2
                values = self.table.row_values(j)
                for x in (range(self.colNum)):
                    s[self.keys[x]] = values[x]
                r.append(s)
                j += 1
            return r


if __name__ == "__main__":
    filepath = "testdata.xlsx"
    sheetName = "Sheet1"
    data = ExcelUtil(filepath, sheetName)
    data1 = data.dict_data()
    print(data1[0])
    print(data1)