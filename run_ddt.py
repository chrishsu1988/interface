import unittest
import os
# from HTMLTestRunner import HTMLTestRunner

cur_dir = os.path.dirname(os.path.realpath(__file__))
case_dir = os.path.join(cur_dir, "testcase")
# print(test_dir)
report_dir = os.path.join(cur_dir, "report")
# print(report_dir)

discover = unittest.defaultTestLoader.discover(case_dir, pattern="test_*.py")

if __name__ == "__main__":
    runner = unittest.TextTestRunner()
    runner.run(discover)