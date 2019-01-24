import requests
import unittest
import os
import ddt
from MyTest.common import base_api, readexcel, writeexecl

# 获取debug.xlsx的路径
cur_path = os.path.dirname(os.path.realpath(__file__))
exc_path = os.path.join(cur_path, "test.xlsx")
# print(exc_path)
# 复制text.xlsx到report下
rep_path = os.path.join(os.path.dirname(cur_path), "report")
# print(rep_path)
res_path = os.path.join(rep_path, "result.xlsx")
# print(res_path)

testdata = readexcel.ExcelUtil(exc_path).dict_data()


@ddt.ddt
class TestApi(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.s = requests.session()
        writeexecl.copy_excel(exc_path, res_path)  # 复制test.xlsx

    @ddt.data(*testdata)
    def test_api(self, data):
        # 复制test.xlsx的数据到result.xlsx
        res = base_api.send_requests(self.s, data)
        test_num = data['id']
        if res["statuscode"] != "200":
            res["error"] = res["text"]
        else:
            res["error"] = ""
        if data["checkpoint"] in res["text"]:
            res["result"] = "pass"
            print("用例的执行结果：%s------%s" % (test_num, res["result"]))
        else:
            res["result"] = "fail"
        base_api.write_result(res, filename=res_path)
        check = data["checkpoint"]
        # 检查点 checkpoint
        # check = data["checkpoint"]
        # print("检查点->：%s" % check)
        # 返回结果
        res_text = res["text"]
        # print("返回实际结果->：%s" % res_text)
        # 断言
        self.assertTrue(check in res_text)

    if __name__ == "__main__":
        unittest.main()









