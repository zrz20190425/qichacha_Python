import time
import requests
import xlrd
from xlutils.copy import copy

def get_html(url, args={}):
    i_headers = {
        "Accept" : "*/*",
        "Accept-Encoding" : "gzip, deflate, br",
        "Accept-Language" : "zh-CN,zh;q=0.9",
        "Connection" : "keep-alive",
        "Cookie" : "BIDUPSID=591ED2E1B51B7C076349BE2216664B31; PSTM=1586917409; BAIDUID=591ED2E1B51B7C0752270D8B1E062D71:FG=1; H_PS_PSSID=30968_1426_31169_21094_31254_31186_31271_31217_30823_31163_31196; BDORZ=B490B5EBF6F3CD402E515D22BCDA1598; delPer=0; PSINO=2; ZD_ENTRY=baidu",
        "Host" : "api.map.baidu.com",
        "Referer" : "https://maplocation.sjfkai.com/",
        "Sec-Fetch-Dest" : "script",
        "Sec-Fetch-Mode" : "no-cors",
        "Sec-Fetch-Site" : "cross-site",
        "User-Agent" : "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Mobile Safari/537.36",
    }
    req = requests.get(url, params=args, headers=i_headers)
    return req.json()


def get_address_text():
    print()


if __name__ == '__main__':
    
    read_xls = xlrd.open_workbook('1.xls')
    read_xls_sheet1 = read_xls.sheets()[0]
    read_xls_sheet1_nrows = read_xls_sheet1.nrows

    write_xls = copy(read_xls)
    write_xls_sheet1 = write_xls.get_sheet(0)

    for i in range(read_xls_sheet1_nrows-2):
        infoStatus = read_xls_sheet1.row_values(rowx=i+2)[17].strip()
        companyName = read_xls_sheet1.row_values(rowx=i+2)[1].strip()
        companyAddress = read_xls_sheet1.row_values(rowx=i+2)[15].strip()
        print("")
        print("序号:" + str(i+1) + "  " + companyName + " - 开始查询")
        if infoStatus == "1":
            print("序号:" + str(i+1) + "  " + companyName + " - 已有数据,跳过")
            continue
        elif read_xls_sheet1.row_values(rowx=i+2)[16] == "2":
            print("序号:" + str(i+1) + "  " + companyName + " - 未查询到该企业,跳过")
            write_xls_sheet1.write(i+2, 17, "2")
            write_xls.save('1.xls')
            continue
        #{'status': 0, 'result': {'location': {'lng': 112.94270003641395, 'lat': 28.358975003354967}, 'precise': 1, 'confidence': 75, 'comprehension': 18, 'level': '购物'}}
        parms = {
            "address" : companyAddress,
            "output" : "json",
            "ak" : "gQsCAgCrWsuN99ggSIjGn5nO"
        }
        location = get_html('https://api.map.baidu.com/geocoder/v2/',parms)
        if location["status"] == 0:
            if location['result']['location']['lng']:
                write_xls_sheet1.write( i+2, 13, location['result']['location']['lng'])
            else:
                continue

            if location['result']['location']['lat']:
                write_xls_sheet1.write( i+2, 14, location['result']['location']['lat'])
            else:
                continue
            write_xls_sheet1.write(i+2, 17, "1")
            write_xls.save('1.xls')
            print("保存信息成功")
        else:
            print("未查到相关地址")
            continue
        time.sleep(3)
    