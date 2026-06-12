# 作者：Winter
# 功能：读取年报链接 Excel，直接下载 PDF，并快速转成 TXT

import os
import re
from urllib.parse import parse_qs, urlparse

import pandas as pd
import pdfplumber
import requests

try:
    import fitz
except:
    fitz = None


# ==================== 配置区：主要改这里 ====================

LINK_EXCEL = r"data\links\中证A50近10年年报链接.xlsx"
REPORT_DIR = "年报文件"
DELETE_PDF = True

# fitz 速度快；如果电脑没有 fitz，会自动改用 pdfplumber
USE_FITZ = True


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36",
    "Referer": "https://www.cninfo.com.cn/",
}


# ==================== 简单函数区 ====================

def clean_filename(text):
    """去掉 Windows 文件名不能用的符号。"""
    return re.sub(r'[\\/:*?"<>|]', "", text)


def get_pdf_url(row):
    """优先读取 PDF链接；旧 Excel 没有 PDF链接时，根据详情链接拼出来。"""
    if "PDF链接" in row and str(row["PDF链接"]) != "nan":
        return str(row["PDF链接"])

    if "详情链接" in row:
        detail_url = str(row["详情链接"])
    elif "年报链接" in row:
        detail_url = str(row["年报链接"])
    else:
        return ""

    query = parse_qs(urlparse(detail_url).query)
    announcement_id = query.get("announcementId", [""])[0]
    announcement_time = query.get("announcementTime", [""])[0]

    if announcement_id == "" or announcement_time == "":
        return ""

    return f"https://static.cninfo.com.cn/finalpage/{announcement_time}/{announcement_id}.PDF"


def download_pdf(pdf_url, pdf_path):
    """下载 PDF。"""
    url_list = [pdf_url]

    if pdf_url.endswith(".PDF"):
        url_list.append(pdf_url[:-4] + ".pdf")
    elif pdf_url.endswith(".pdf"):
        url_list.append(pdf_url[:-4] + ".PDF")

    for url in url_list:
        for try_num in range(3):
            try:
                res = requests.get(url, headers=HEADERS, timeout=60)

                if res.status_code != 200:
                    continue

                if not res.content.startswith(b"%PDF"):
                    continue

                with open(pdf_path, "wb") as f:
                    f.write(res.content)

                return True

            except:
                continue

    return False


def pdf_to_txt_by_fitz(pdf_path, txt_path):
    """用 fitz 把 PDF 转成 TXT，速度比较快。"""
    pdf = fitz.open(pdf_path)

    with open(txt_path, "w", encoding="utf-8") as f:
        for page in pdf:
            text = page.get_text()
            if text:
                f.write(text)
                f.write("\n")

    pdf.close()


def pdf_to_txt_by_pdfplumber(pdf_path, txt_path):
    """用 pdfplumber 把 PDF 转成 TXT。"""
    with pdfplumber.open(pdf_path) as pdf:
        with open(txt_path, "w", encoding="utf-8") as f:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    f.write(text)
                    f.write("\n")


def pdf_to_txt(pdf_path, txt_path):
    """优先用 fitz，失败或没有 fitz 时再用 pdfplumber。"""
    if USE_FITZ and fitz is not None:
        try:
            pdf_to_txt_by_fitz(pdf_path, txt_path)
            return
        except Exception as e:
            print("fitz 转换失败，改用 pdfplumber：", e)

    pdf_to_txt_by_pdfplumber(pdf_path, txt_path)


# ==================== 主程序 ====================

df = pd.read_excel(LINK_EXCEL, dtype={"公司代码": str, "年份": str})
finished_count = 0

for i in range(len(df)):
    row = df.loc[i]

    stock_code = str(row["公司代码"]).zfill(6)
    company_name = str(row["公司简称"])
    report_year = str(row["年份"])
    report_title = str(row["标题"])

    print("\n" + "=" * 60)
    print(f"正在处理第 {i + 1}/{len(df)} 份")
    print(stock_code, company_name, report_year, report_title)

    pdf_dir = os.path.join(REPORT_DIR, report_year, "pdf年报")
    txt_dir = os.path.join(REPORT_DIR, report_year, "txt年报")
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(txt_dir, exist_ok=True)

    file_name = clean_filename(f"{stock_code}_{company_name}_{report_year}")
    pdf_path = os.path.join(pdf_dir, file_name + ".pdf")
    txt_path = os.path.join(txt_dir, file_name + ".txt")

    if os.path.exists(txt_path):
        print("TXT 已存在，跳过")
        finished_count = finished_count + 1
        if finished_count % 50 == 0:
            print(f"已经完成 {finished_count} 份")
        continue

    pdf_url = get_pdf_url(row)

    if pdf_url == "":
        print("没有找到 PDF 链接")
        continue

    print("PDF 链接：", pdf_url)

    if os.path.exists(pdf_path):
        print("PDF 已存在，直接转 TXT：", pdf_path)
    else:
        if not download_pdf(pdf_url, pdf_path):
            print("PDF 下载失败")
            continue

        print("PDF 已保存：", pdf_path)

    try:
        pdf_to_txt(pdf_path, txt_path)
        print("TXT 已保存：", txt_path)

        if DELETE_PDF:
            os.remove(pdf_path)
            print("PDF 已删除，只保留 TXT")

        finished_count = finished_count + 1
        if finished_count % 50 == 0:
            print(f"已经完成 {finished_count} 份")

    except Exception as e:
        print("PDF 转 TXT 失败：", e)

print("\n全部完成")
print("成功完成：", finished_count, "份")
