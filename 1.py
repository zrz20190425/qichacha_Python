import time
from bs4 import BeautifulSoup
import xlrd
from xlutils.copy import copy

from selenium import webdriver
from selenium.webdriver import ChromeOptions



def get_html_selenium(url):

    chrome_options = ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option(
        'excludeSwitches', ['enable-automation'])
    chrome_options.add_argument('--disable-javascript')
    chrome_options.add_argument('--blink-settings=imagesEnabled=false')
    driver = webdriver.Chrome(executable_path="/Users/Mac/Development/Environment/Selenium/bin/chromedriver",
                              options=chrome_options, desired_capabilities={"pageLoadStrategy": "eager"})
    driver.get(url)
    content = driver.page_source.encode('utf-8')
    driver.close()
    return content


def get_search(searchKey):
    # http://m.qcc.com/search?key=%E4%B8%AD%E5%88%9B%E6%99%BA%E7%BB%B4%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8
    html_doc = get_html_selenium('http://m.qcc.com/search?key='+searchKey)
    soup = BeautifulSoup(html_doc, 'html.parser')

    if soup.select_one('div[class="nodata"]'):
        return ""

    companyName = soup.select_one(
        'div[class="list-item-name"]').get_text().strip()
    if str(searchKey).strip() == companyName:
        detailUrl = soup.select_one(
            'a[class="a-decoration"]').attrs['href'].strip()
        return "http://m.qcc.com"+detailUrl
    else:
        return ""


def get_detail(detailUrl):
    # https://m.qcc.com/firm_e74160e6ad4685a9a9ea61bdf1b96490.html
    html_doc = get_html_selenium(detailUrl)
    soup = BeautifulSoup(html_doc, 'html.parser')
    detailInfoDict = dict()
    # 法人代表
    detailInfoDict["companyBoss"] = soup.select_one(
        'a[class="text-primary oper"]').text.strip()
    # 电话号码
    try:
        detailInfoDict["companyPhone"] = soup.select_one('a[class="phone a-decoration"]').text.strip()
    except AttributeError:
        detailInfoDict["companyPhone"] = "-"    
    # 电子邮箱
    try:
        detailInfoDict["companyEmail"] = soup.select_one(
            'a[class="email a-decoration"]').text.strip()
    except AttributeError:
        detailInfoDict["companyEmail"] = "-"

    # 地理位置 companyAddress
    detailInfoDict["companyAddress"] = soup.select_one(
        'div[class="address"]').text.strip()
    # 公司状态 companyStatus
    detailInfoDict["companyStatus"] = soup.select_one(
        'span[class="ntag text-success"]').text.strip()

    for row in soup.select_one('table[class="info-table"]').find_all('td')[1:]:
        className = row.select_one('div[class="d"]').text.strip()
        classText = row.select_one('div[class="v"]').text.strip()
        if className == "成立日期":
            # 注册日期 companyRegistrationDate
            detailInfoDict["companyRegistrationDate"] = classText
        elif className == "注册资本":
            # 注册资本 companyRegisteredCapital
            detailInfoDict["companyRegisteredCapital"] = classText
        elif className == "统一社会信用代码":
            # 证件号码 TaxpayerID
            detailInfoDict["TaxpayerID"] = classText
        elif className == "组织机构代码":
            # 公司编码 organizationCode
            detailInfoDict["organizationCode"] = classText
        elif className == "经营范围":
            # 经营范围 businessScope
            detailInfoDict["businessScope"] = classText
        elif className == "所属行业":
            # 分类类别 companyClass
            detailInfoDict["companyClass"] = classText
        elif className == "营业期限":
            # 有效期至 businessTerm
            detailInfoDict["businessTerm"] = classText.split('至')[-1].strip()

    # 地理经度
    # 地理纬度
    return detailInfoDict


if __name__ == '__main__':

    read_xls = xlrd.open_workbook('1.xls')
    read_xls_sheet1 = read_xls.sheets()[0]
    read_xls_sheet1_nrows = read_xls_sheet1.nrows

    write_xls = copy(read_xls)
    write_xls_sheet1 = write_xls.get_sheet(0)

    for i in range(read_xls_sheet1_nrows-2):
        infoStatus = read_xls_sheet1.row_values(rowx=i+2)[16]
        companyName = read_xls_sheet1.row_values(rowx=i+2)[1].strip()
        print("")
        print("序号:" + str(i+1) + "  " + companyName + " - 开始查询")
        if infoStatus == "1":
            print("序号:" + str(i+1) + "  " + companyName + " - 已有数据,跳过")
            continue
        elif infoStatus == "2":
            print("序号:" + str(i+1) + "  " + companyName + " - 未查询到该企业,跳过")
            continue
        detailUrl = get_search(companyName)
        if detailUrl:
            print("企业查询信息成功，详情链接：" + detailUrl)
            # time.sleep(2)
            detailInfoDict = get_detail(detailUrl)
            # 公司简介
            detailInfoDict['companyProfile'] = companyName + "成立于" + detailInfoDict['companyRegistrationDate'] + "，注册地位于" + \
                detailInfoDict['companyAddress'] + "，法人代表人为" + detailInfoDict['companyBoss'] + \
                "，经营范围包括" + detailInfoDict['businessScope'] + "。"
            write_xls_sheet1.write( i+2, 2, detailInfoDict['companyProfile'])

            if(detailInfoDict['businessScope']):  # 经营范围
                write_xls_sheet1.write(i+2, 3, detailInfoDict['businessScope'])

            if(detailInfoDict['organizationCode']):  # 公司编码
                write_xls_sheet1.write(
                    i+2, 4, detailInfoDict['organizationCode'])

            if(detailInfoDict['companyClass']):  # 分类类别
                write_xls_sheet1.write(i+2, 5, detailInfoDict['companyClass'])

            if(detailInfoDict['TaxpayerID']):  # 证件号码
                write_xls_sheet1.write(i+2, 6, detailInfoDict['TaxpayerID'])

            if(detailInfoDict['companyBoss']):  # 法人代表
                write_xls_sheet1.write(i+2, 7, detailInfoDict['companyBoss'])

            if(detailInfoDict['companyRegisteredCapital']):  # 注册资本
                write_xls_sheet1.write(
                    i+2, 8, detailInfoDict['companyRegisteredCapital'])

            if(detailInfoDict['companyPhone']):  # 电话号码
                write_xls_sheet1.write(i+2, 9, detailInfoDict['companyPhone'])

            if(detailInfoDict['companyEmail']):  # 电子邮箱
                write_xls_sheet1.write(i+2, 10, detailInfoDict['companyEmail'])

            if(detailInfoDict['companyRegistrationDate']):  # 注册日期
                write_xls_sheet1.write(
                    i+2, 11, detailInfoDict['companyRegistrationDate'])

            if(detailInfoDict['businessTerm']):  # 有效期至
                write_xls_sheet1.write(i+2, 12, detailInfoDict['businessTerm'])

            if(detailInfoDict['companyAddress']):  # 地理位置
                write_xls_sheet1.write(
                    i+2, 15, detailInfoDict['companyAddress'])
            
            write_xls_sheet1.write(i+2, 16, "1")
            write_xls.save('1.xls')

            print("保存信息成功")
            # time.sleep(2)
        else:
            write_xls_sheet1.write(i+2, 16, "2")
            write_xls.save('1.xls')
            print("未查询到该企业")
