# 作者：Winter
# 功能：把年报词频结果和外部行业城市数据合并，方便后面做行业、城市层面的可视化

import os

import pandas as pd


# ==================== 配置区：主要改这里 ====================

# 我们自己分析出来的年报词频结果
ANALYSIS_EXCEL = r"年报文件\中证A50近10年数字化词频分析结果.xlsx"
ANALYSIS_SHEET = "原始统计"

# 你提供的上市公司数据，这里只使用行业和城市信息
EXTERNAL_EXCEL = r"D:\桌面文件\数据\2024上市公司智能化转型\智能化转型程度（未缩尾未剔除）.xlsx"
EXTERNAL_SHEET = "Sheet1"

# 合并后的结果
OUTPUT_EXCEL = r"年报文件\中证A50近10年数字化词频_合并行业城市.xlsx"


# ==================== 简单函数区 ====================

def format_stock_code(code):
    """把股票代码统一成 6 位，例如 1 -> 000001。"""
    code = str(code).strip()

    if "." in code:
        code = code.split(".")[0]

    return code.zfill(6)


def read_analysis_data():
    """读取我们自己算出来的年报关键词结果。"""
    df = pd.read_excel(ANALYSIS_EXCEL, sheet_name=ANALYSIS_SHEET, dtype={"股票代码": str})
    df["股票代码"] = df["股票代码"].apply(format_stock_code)
    df["年份"] = df["年份"].astype(int)
    return df


def read_external_data():
    """读取外部数据，只保留行业、省份、城市信息，并统一字段名。"""
    df = pd.read_excel(EXTERNAL_EXCEL, sheet_name=EXTERNAL_SHEET)

    df["股票代码"] = df["证券代码"].apply(format_stock_code)
    df["年份"] = df["year"].astype(int)

    keep_columns = [
        "股票代码",
        "年份",
        "证券简称",
        "行业代码",
        "行业名称",
        "所属省份",
        "所属省份代码",
        "所属城市",
        "所属城市代码",
    ]

    df = df[keep_columns]
    return df


def find_previous_info(stock_code, year, external_df):
    """查找同一家公司当年或以前最近一年的行业、城市信息。"""
    sub_df = external_df[
        (external_df["股票代码"] == stock_code)
        & (external_df["年份"] <= year)
    ].copy()

    if len(sub_df) == 0:
        return None

    sub_df = sub_df.sort_values("年份", ascending=False)
    return sub_df.iloc[0]


def merge_with_previous_year(analysis_df, external_df):
    """按 股票代码+年份 合并；没有当年数据时，用之前年份补齐。"""
    rows = []

    for i in range(len(analysis_df)):
        row = analysis_df.loc[i].copy()
        stock_code = row["股票代码"]
        year = int(row["年份"])

        info = find_previous_info(stock_code, year, external_df)

        if info is None:
            row["证券简称"] = ""
            row["行业代码"] = ""
            row["行业名称"] = ""
            row["所属省份"] = ""
            row["所属省份代码"] = ""
            row["所属城市"] = ""
            row["所属城市代码"] = ""
            row["属性来源年份"] = ""
            row["匹配方式"] = "未匹配"
        else:
            row["证券简称"] = info["证券简称"]
            row["行业代码"] = info["行业代码"]
            row["行业名称"] = info["行业名称"]
            row["所属省份"] = info["所属省份"]
            row["所属省份代码"] = info["所属省份代码"]
            row["所属城市"] = info["所属城市"]
            row["所属城市代码"] = info["所属城市代码"]
            row["属性来源年份"] = int(info["年份"])

            if int(info["年份"]) == year:
                row["匹配方式"] = "当年匹配"
            else:
                row["匹配方式"] = "之前年份补齐"

        rows.append(row)

    return pd.DataFrame(rows)


def make_summary(merged_df, group_columns):
    """按照行业、城市、省份等维度汇总。"""
    summary_df = merged_df.groupby(group_columns, as_index=False).agg(
        样本数量=("股票代码", "count"),
        公司数量=("股票代码", "nunique"),
        平均总词数=("总词数", "mean"),
        数字化关键词总次数=("数字化关键词总次数", "sum"),
        平均标准化词频_每万词=("标准化词频_每万词", "mean"),
    )

    summary_df["平均总词数"] = summary_df["平均总词数"].round(2)
    summary_df["平均标准化词频_每万词"] = summary_df["平均标准化词频_每万词"].round(4)

    return summary_df


# ==================== 主程序 ====================

analysis_df = read_analysis_data()
external_df = read_external_data()

merged_df = merge_with_previous_year(analysis_df, external_df)

matched_count = (merged_df["匹配方式"] != "未匹配").sum()
same_year_count = (merged_df["匹配方式"] == "当年匹配").sum()
previous_year_count = (merged_df["匹配方式"] == "之前年份补齐").sum()
unmatched_count = (merged_df["匹配方式"] == "未匹配").sum()

industry_year_df = make_summary(merged_df, ["行业名称", "年份"])
city_year_df = make_summary(merged_df, ["所属城市", "年份"])
province_year_df = make_summary(merged_df, ["所属省份", "年份"])

industry_total_df = make_summary(merged_df, ["行业名称"])
city_total_df = make_summary(merged_df, ["所属城市"])
province_total_df = make_summary(merged_df, ["所属省份"])

check_df = pd.DataFrame([
    ["年报词频数据行数", len(analysis_df)],
    ["外部智能化转型数据行数", len(external_df)],
    ["成功匹配行数", matched_count],
    ["当年匹配行数", same_year_count],
    ["之前年份补齐行数", previous_year_count],
    ["未匹配行数", unmatched_count],
], columns=["项目", "数量"])

unmatched_df = merged_df[merged_df["匹配方式"] == "未匹配"][["股票代码", "公司简称", "年份"]]

os.makedirs(os.path.dirname(OUTPUT_EXCEL), exist_ok=True)

with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
    merged_df.to_excel(writer, sheet_name="合并明细", index=False)
    industry_year_df.to_excel(writer, sheet_name="行业年度汇总", index=False)
    city_year_df.to_excel(writer, sheet_name="城市年度汇总", index=False)
    province_year_df.to_excel(writer, sheet_name="省份年度汇总", index=False)
    industry_total_df.to_excel(writer, sheet_name="行业总体汇总", index=False)
    city_total_df.to_excel(writer, sheet_name="城市总体汇总", index=False)
    province_total_df.to_excel(writer, sheet_name="省份总体汇总", index=False)
    check_df.to_excel(writer, sheet_name="匹配检查", index=False)
    unmatched_df.to_excel(writer, sheet_name="未匹配记录", index=False)

print("合并完成")
print("年报词频数据行数：", len(analysis_df))
print("成功匹配行数：", matched_count)
print("当年匹配行数：", same_year_count)
print("之前年份补齐行数：", previous_year_count)
print("未匹配行数：", unmatched_count)
print("结果已保存到：", OUTPUT_EXCEL)
