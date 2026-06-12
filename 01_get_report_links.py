# 作者：Winter
# 功能：用巨潮资讯接口查询中证A50公司近10年年报链接和PDF链接

import os
import re
import time

import pandas as pd
import requests


# ==================== 配置区：主要改这里 ====================

START_YEAR = 2016
END_YEAR = 2025
COMPANY_NUM = 50

OUTPUT_EXCEL = r"data\links\中证A50近10年年报链接.xlsx"


# 中证A50公司：股票代码、公司简称
COMPANIES = [
    ["002594", "比亚迪"],
    ["600900", "长江电力"],
    ["300750", "宁德时代"],
    ["601012", "隆基绿能"],
    ["300760", "迈瑞医疗"],
    ["600030", "中信证券"],
    ["601919", "中远海控"],
    ["600276", "恒瑞医药"],
    ["601318", "中国平安"],
    ["600519", "贵州茅台"],
    ["002415", "海康威视"],
    ["601166", "兴业银行"],
    ["601899", "紫金矿业"],
    ["603501", "韦尔股份"],
    ["600690", "海尔智家"],
    ["000568", "泸州老窖"],
    ["600309", "万华化学"],
    ["002475", "立讯精密"],
    ["002352", "顺丰控股"],
    ["600887", "伊利股份"],
    ["600050", "中国联通"],
    ["300124", "汇川技术"],
    ["600028", "中国石化"],
    ["600089", "特变电工"],
    ["600150", "中国船舶"],
    ["688981", "中芯国际"],
    ["000858", "五粮液"],
    ["600031", "三一重工"],
    ["000333", "美的集团"],
    ["601888", "中国中免"],
    ["601668", "中国建筑"],
    ["601398", "工商银行"],
    ["601601", "中国太保"],
    ["600436", "片仔癀"],
    ["601288", "农业银行"],
    ["002714", "牧原股份"],
    ["600438", "通威股份"],
    ["600406", "国电南瑞"],
    ["688111", "金山办公"],
    ["601658", "邮储银行"],
    ["600941", "中国移动"],
    ["600426", "华鲁恒升"],
    ["600989", "宝丰能源"],
    ["600132", "重庆啤酒"],
    ["300059", "东方财富"],
    ["603288", "海天味业"],
    ["601225", "陕西煤业"],
    ["002371", "北方华创"],
    ["600588", "用友网络"],
    ["601138", "工业富联"],
]


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36",
    "Referer": "https://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search",
}


# ==================== 简单函数区 ====================

def get_board(stock_code):
    """根据股票代码判断板块。"""
    if stock_code.startswith(("0", "2", "3")):
        return "深市"
    if stock_code.startswith(("6", "9")):
        return "沪市"
    return ""


def clean_title(title):
    """去掉接口标题中的 <em> 标签。"""
    return re.sub(r"<.*?>", "", title)


def get_date(ms_time):
    """把毫秒时间转成日期。"""
    return time.strftime("%Y-%m-%d", time.localtime(ms_time / 1000))


def is_report_title(title, report_year):
    """判断是不是正式年报。"""
    bad_words = ["摘要", "英文", "半年度", "一季度", "三季度", "问询函", "回复"]

    for word in bad_words:
        if word in title:
            return False

    if str(report_year) in title and "年度报告" in title:
        return True

    return False


def get_org_id_dict():
    """获取 股票代码 -> orgId。"""
    url = "https://www.cninfo.com.cn/new/data/szse_stock.json"
    res = requests.get(url, headers=HEADERS, timeout=20)
    stock_data = res.json()["stockList"]

    org_id_dict = {}

    for item in stock_data:
        org_id_dict[item["code"]] = item["orgId"]

    return org_id_dict


def get_report_data(stock_code, company_name, report_year, org_id):
    """查询某家公司某一年的年报信息。"""
    query_url = "https://www.cninfo.com.cn/new/hisAnnouncement/query"
    begin_date = f"{report_year + 1}-01-01"
    end_date = f"{report_year + 1}-06-30"

    # 第一次用精确关键词；第二次放宽，不填标题关键词
    keyword_list = [f"{report_year}年年度报告", ""]

    for keyword in keyword_list:
        data = {
            "pageNum": "1",
            "pageSize": "30",
            "column": "szse",
            "tabName": "fulltext",
            "plate": "",
            "stock": stock_code + "," + org_id,
            "searchkey": keyword,
            "secid": "",
            "category": "category_ndbg_szsh",
            "trade": "",
            "seDate": begin_date + "~" + end_date,
            "sortName": "",
            "sortType": "",
            "isHLtitle": "true",
        }

        try:
            res = requests.post(query_url, headers=HEADERS, data=data, timeout=20)
            result = res.json()
        except:
            print("查询失败：", stock_code, company_name, report_year)
            continue

        announcements = result.get("announcements")

        if announcements is None:
            continue

        for item in announcements:
            title = clean_title(item["announcementTitle"])

            if not is_report_title(title, report_year):
                continue

            announcement_date = get_date(item["announcementTime"])
            detail_url = (
                "https://www.cninfo.com.cn/new/disclosure/detail"
                + f"?stockCode={item['secCode']}"
                + f"&announcementId={item['announcementId']}"
                + f"&orgId={item['orgId']}"
                + f"&announcementTime={announcement_date}"
            )

            pdf_url = "https://static.cninfo.com.cn/" + item["adjunctUrl"]

            return {
                "公司代码": item["secCode"],
                "公司简称": item["secName"],
                "板块": get_board(stock_code),
                "年份": report_year,
                "标题": title,
                "详情链接": detail_url,
                "PDF链接": pdf_url,
                "公告日期": announcement_date,
            }

    return None


def save_excel(rows):
    """保存结果。"""
    df = pd.DataFrame(rows)
    os.makedirs(os.path.dirname(OUTPUT_EXCEL), exist_ok=True)
    df.to_excel(OUTPUT_EXCEL, index=False)


# ==================== 主程序 ====================

rows = []
org_id_dict = get_org_id_dict()
year_list = list(range(START_YEAR, END_YEAR + 1))

for company in COMPANIES[:COMPANY_NUM]:
    stock_code = company[0]
    company_name = company[1]
    org_id = org_id_dict.get(stock_code, "")

    if org_id == "":
        print("没有找到 orgId：", stock_code, company_name)
        continue

    for report_year in year_list:
        print("\n" + "=" * 60)
        print("正在查询：", stock_code, company_name, report_year, "年报")

        report_data = get_report_data(stock_code, company_name, report_year, org_id)

        if report_data is None:
            print("没有找到正式年报：", stock_code, company_name, report_year)
        else:
            rows.append(report_data)
            print("找到年报：", report_data["标题"], report_data["公告日期"])
            save_excel(rows)

            if len(rows) % 50 == 0:
                print(f"已经获取 {len(rows)} 份年报链接")

        time.sleep(0.5)

save_excel(rows)

print("\n全部完成")
print("一共获取：", len(rows), "份年报链接")
print("年报链接数量：", len(rows))
print("保存位置：", OUTPUT_EXCEL)
