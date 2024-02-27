import requests
from bs4 import BeautifulSoup
import re
import openpyxl


wb = openpyxl.Workbook()
month = ["01", "02", "03", "04", "05", "06", "07", "08","09", "10", "11"]
page = 0
for k in month:
    # print(f"------------------------{k}月------------------------------")
    url = f"https://stock.wearn.com/cdata.asp?Year=111&month={k}&kind=00878"

    res = requests.get(url) # 請求網址
    soup = BeautifulSoup(res.text, "html.parser") # 讀取html內容
    dates = soup.select('td[class = "table-first-child"]')
    infors = soup.select('td[class = "table-first-child"] ~ td[ align = "right"]')

    # 儲存日期
    date = []
    for i in dates:
        date.append(i.text)
    # print(date)

    # 儲存資料
    count = 0
    temp = []
    infor_org = []
    for j in infors:
        count += 1
        temp.append(j.text)
        if count == 5:
            infor_org.append(temp)
            count = 0
            temp = []
    # print(infor_org)

    # 整理資料
    infor_new = []
    pattern = r'^\d+.\d+'
    temp = []
    for i in range(len(infor_org)):
        temp.append(date[i])
        for j in range(5):
            data = re.search(pattern, infor_org[i][j])
            if j < 4:
                data = data.group()
            elif j == 4:
                data = data.group().replace(",", "")
            temp.append(data)
            # print(data, end="\t")
        infor_new.append(temp)
        temp = []
        # print("")
    # print(len(date))
    # print(len(infor_org))
    infor_new.reverse()
    # for i in infor_new:
    #     print(i)
    
    sheet = wb.create_sheet(f"{k}月", page)     # 建立空白的工作表
    for i in infor_new:
        sheet.append(i) 
    page += 1

wb.save('data.xlsx')

