import unittest
import os
from MyTest.common.HTMLTestRunner_api import HTMLTestRunner

cur_dir = os.path.dirname(os.path.realpath(__file__))
case_dir = os.path.join(cur_dir, "testcase")
# print(test_dir)
report_dir = os.path.join(cur_dir, "report")
# print(report_dir)

discover = unittest.defaultTestLoader.discover(case_dir, pattern="test_*.py")

if __name__ == "__main__":
    filename = report_dir + "/" + "result.html"
    fp = open(filename, "wb")
    runner = HTMLTestRunner(stream=fp, verbosity=2, title="测试报告", description="测试用例执行情况： ")
    runner.run(discover)
    fp.close()