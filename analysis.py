import openpyxl
import numpy as np
import matplotlib.pyplot as plt

month = ["01月", "02月", "03月", "04月", "05月", \
    "06月", "07月", "08月", "09月", "10月", "11月"]
wb = openpyxl.load_workbook("data.xlsx")

# 初始化陣列
op_price = []
h_price = []
l_price = []
cl_price = []
trading_v = []
date_amount = [1]    # 計算天數
day = 0
for k in month:
    # 工作表
    sheet = wb[k]
    # 需要轉換資料型態，之前是字串
    # 開盤價
    for row in sheet['B']:
        op_price.append(float(row.value))

    # 最高價
    for row in sheet['C']:
        h_price.append(float(row.value))

    # 最低價
    for row in sheet['D']:
        l_price.append(float(row.value))

    # 收盤價
    for row in sheet['E']:
        cl_price.append(float(row.value))

    # 成交量
    for row in sheet['F']:
        trading_v.append(float(row.value))
    
    # 計算天數
    for row in sheet['B']:
        day += 1
    if len(date_amount) < len(month):
        date_amount.append(day)

# 計算漲跌幅
gap = []
for i in range(len(cl_price)):
    if i + 1 < len(cl_price):
        d = (cl_price[i+1]-cl_price[i])/cl_price[i]*100
        gap.append(d)
t = np.array(range(1, len(gap)+1))
g = np.array(gap)

# 處理標籤
month_label = ['Jan', 'Feb', 'Mar', 'April', 'May', 'June', \
    'July', 'Aug', 'Sep', 'Oct', 'Nov']


#繪圖
plt.figure(figsize=(12,6))
# plt.figure(1)

# 折線圖
plt.subplot(221)

x = np.array(range(1, len(op_price) + 1))
op = np.array(op_price)
h = np.array(h_price)
l = np.array(l_price)
cl = np.array(cl_price)
plt.plot(x, op, color = "tab:red",label="opening price")
# plt.plot(x, h, color = "tab:blue",label="highest price")
# plt.plot(x, l, color = "tab:orange",label="lowest price")
plt.plot(x, cl, color = "tab:green",label="closing price")
plt.xticks(date_amount, month_label, rotation = 30, fontsize=8)
plt.title("Market Trend",fontweight="bold")
plt.ylabel("Price")
plt.legend(loc="best")
plt.grid('true')

#直方圖
plt.subplot(222)
td = np.array(trading_v)
x = np.array(range(1, len(trading_v) + 1))

plt.bar(x,td,0.5, color = "tab:green")
plt.xticks(date_amount, month_label, rotation = 30,fontsize=8)
plt.yticks(np.arange(0,150000,15000))
plt.title("Trading Volume",fontweight="bold")
plt.ylabel("Amount")
plt.grid('true')

#點圖
plt.subplot(223)
plt.scatter(t, g, s=10, color = "tab:green")
plt.title("Rate of Price Spread",fontweight="bold")
plt.xticks(date_amount, month_label, rotation = 30,fontsize=8)
plt.ylabel('Rate')
plt.grid('true')

#漲跌幅直方圖
plt.subplot(224)
bin=np.arange(-2.5,2.5,0.5)
plt.hist(g,bin,orientation='horizontal', color = "tab:green")
plt.yticks(np.arange(-2.5, 2.5, 0.5))
plt.title("Rate of Price Spread",fontweight="bold")
plt.xlabel('Count')
plt.ylabel('Rate')
plt.grid('true')
plt.tight_layout()
plt.show()