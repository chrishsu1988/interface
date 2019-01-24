import requests
import json
from MyTest.common.readexcel import ExcelUtil
from MyTest.common.writeexecl import WriteExcel, copy_excel
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def send_requests(s, testdata):
    method = testdata["method"]
    url = testdata["url"]
    test_num = testdata['id']
    print("正在执行的用例：******------ %s ------******" % test_num)
    # url后面的params参数
    try:
        params = eval(testdata["params"])
    except:
        params = ""
    # 请求头headers
    try:
        headers = eval(testdata["headers"])
        print("请求头部:%s" % headers)
    except:
        headers = None

    # post请求body类型
    type = testdata["type"]
    print("请求方式：%s, 请求url：%s" % (method, url))
    # post请求body内容
    try:
        bodydata = eval(testdata["body"])
    except:
        bodydata = {}
    # 判断是data数据还是json
    if type == "data":
        body = bodydata
    if type == "json":
        body = json.dumps(bodydata)
    else:
        body = bodydata

    if method == "post":
        print("post请求的body类型为：%s, body的内容为：%s" % (type, body))
    verify = False
    res = {}  # 接收返回的数据
    try:
        r = s.request(method=method,
                      url=url,
                      params=params,
                      headers=headers,
                      data=body,
                      verify=verify
                      )
        print("接口返回的结果为：%s" % r.content.decode("utf-8"))
        res['id'] = testdata['id']
        res['rowNum'] = testdata['rowNum']
        res["statuscode"] = str(r.status_code)  # 状态码转成str
        res["text"] = r.content.decode("utf-8")
        res["times"] = str(r.elapsed.total_seconds())  # 接口请求时间转str
        # if res["statuscode"] != "200":
        #     res["error"] = res["text"]
        # else:
        #     res["error"] = ""
        # res["msg"] = ""
        # if testdata["checkpoint"] in res["text"]:
        #     res["result"] = "pass"
        #     print("用例的执行结果：%s------%s" % (test_num, res["result"]))
        # else:
        #     res["result"] = "fail"
        res["msg"] = ""
        return res

    except Exception as msg:
        res["msg"] = str(msg)
        return res


def write_result(result, filename="result.xlsx"):
    # 返回结果的行数row_num
    row_num = result["rowNum"]
    # 写入返回的结果
    wt = WriteExcel(filename)
    wt.write(row_num, 8, result['statuscode'])  # 写入返回状态码statuscode,第8列
    wt.write(row_num, 9, result['times'])    # 耗时
    wt.write(row_num, 10, result['error'])   # 状态码非200时的返回信息
    wt.write(row_num, 12, result['result'])  # 测试结果 pass 还是fail
    wt.write(row_num, 13, result['msg'])     # 抛异常


if __name__ == "__main__":
    data = ExcelUtil("debug.xlsx").dict_data()
    # print(data[0])
    s = requests.session()
    res = send_requests(s, data[0])
    copy_excel("debug.xlsx", "result.xlsx")
    write_result(res, filename="result.xlsx")






