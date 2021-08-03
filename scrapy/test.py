import requests

url = "http://app1.nmpa.gov.cn/data_nmpa/face3/base.jsp?" \
      "tableId=23&tableName=TABLE23&title=GMP%C8%CF%D6%A4&bcId=152911797956855568058860206743"

# url = "http://wjw.beijing.gov.cn/xwzx_20031/wnxw/202007/t20200706_1939095.html"

Headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.164 Safari/537.36'}

res = requests.get(url, headers=Headers)
print("状态码： ", res.status_code)
while res.status_code == 202:
    res = requests.get(url, headers=Headers)
    print("状态码： ", res.status_code)


Content = res.text
print("Content: \n", Content)


