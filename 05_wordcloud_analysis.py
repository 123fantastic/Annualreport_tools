# 作者：Winter
# 功能：生成年份阶段词云、行业前5关键词柱状图、城市分阶段排名图

import os
import re

import matplotlib.pyplot as plt
import pandas as pd
from digital_keywords import KEYWORDS
from pyecharts import options as opts
from pyecharts.charts import Geo, Map
from pyecharts.globals import ChartType
from visual_utils import BLUE, GREEN, MAP_COLORS, save_html_as_png
from wordcloud import WordCloud


# ==================== 配置区：主要改这里 ====================

INPUT_EXCEL = r"年报文件\中证A50近10年数字化词频_合并行业城市.xlsx"
INPUT_SHEET = "合并明细"

OUTPUT_DIR = r"年报文件\可视化结果"
YEAR_WORDCLOUD_DIR = os.path.join(OUTPUT_DIR, "年份阶段词云")
INDUSTRY_BAR_DIR = os.path.join(OUTPUT_DIR, "行业前5关键词柱状图")
CITY_BAR_DIR = os.path.join(OUTPUT_DIR, "城市阶段排名图")
PROVINCE_MAP_DIR = os.path.join(OUTPUT_DIR, "省份地图")
CITY_MAP_DIR = os.path.join(OUTPUT_DIR, "城市热力地图")
OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, "可视化数据汇总.xlsx")

# 前五年、后五年划分
FIRST_START_YEAR = 2016
FIRST_END_YEAR = 2020
SECOND_START_YEAR = 2021
SECOND_END_YEAR = 2025

# Windows 中文字体
FONT_PATH = r"C:\Windows\Fonts\simhei.ttf"

# 展示数量
TOP_INDUSTRY_NUM = 10
TOP_CITY_NUM = 15
TOP_KEYWORD_NUM = 5


# ==================== 简单函数区 ====================

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


def safe_filename(name):
    """去掉 Windows 文件名不能用的符号。"""
    name = str(name)
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def clean_province_name(name):
    """把省份名称处理成 pyecharts 中国地图能识别的格式。"""
    name = str(name).strip()
    name = name.replace("省", "")
    name = name.replace("市", "")
    name = name.replace("壮族自治区", "")
    name = name.replace("回族自治区", "")
    name = name.replace("维吾尔自治区", "")
    name = name.replace("自治区", "")
    name = name.replace("特别行政区", "")
    return name


def clean_city_name(name):
    """把城市名称处理成 pyecharts Geo 常用格式。"""
    name = str(name).strip()
    name = name.replace("市", "")
    name = name.replace("特别行政区", "")
    return name


def make_word_count(df):
    """把关键词列汇总成 词 -> 次数。"""
    word_count = {}

    for word in KEYWORDS:
        if word in df.columns:
            count = df[word].fillna(0).sum()
            if count > 0:
                word_count[word] = int(count)

    return word_count


def word_count_to_rows(group_type, group_name, word_count, top_n=None):
    """把词频字典转成表格行。"""
    rows = []
    sorted_words = sorted(word_count.items(), key=lambda x: x[1], reverse=True)

    if top_n is not None:
        sorted_words = sorted_words[:top_n]

    for rank, item in enumerate(sorted_words, start=1):
        rows.append({
            "类型": group_type,
            "名称": group_name,
            "排名": rank,
            "关键词": item[0],
            "词频": item[1],
        })

    return rows


def save_wordcloud(word_count, output_png):
    """根据词频字典保存词云图。"""
    if len(word_count) == 0:
        print("没有词频数据，跳过：", output_png)
        return

    wc = WordCloud(
        font_path=FONT_PATH,
        width=1400,
        height=900,
        background_color="white",
        max_words=100,
        collocations=False,
    )

    image = wc.generate_from_frequencies(word_count)
    image.to_file(output_png)
    print("已保存词云：", output_png)


def save_keyword_bar(word_count, title, output_png):
    """保存前5关键词横向柱状图。"""
    sorted_words = sorted(word_count.items(), key=lambda x: x[1], reverse=True)[:TOP_KEYWORD_NUM]

    if len(sorted_words) == 0:
        print("没有关键词数据，跳过：", output_png)
        return

    words = [item[0] for item in sorted_words][::-1]
    counts = [item[1] for item in sorted_words][::-1]

    plt.figure(figsize=(7, 4))
    plt.barh(words, counts, color=BLUE)
    plt.title(title)
    plt.xlabel("词频")
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存柱状图：", output_png)


def make_stage_df(df, start_year, end_year):
    """筛选一个年份阶段。"""
    return df[(df["年份"] >= start_year) & (df["年份"] <= end_year)]


def make_city_summary(df, stage_name):
    """城市阶段汇总，用于排名图或后续地图。"""
    city_df = df.dropna(subset=["所属城市"]).copy()

    summary_df = city_df.groupby(["所属省份", "所属城市", "所属城市代码"], as_index=False).agg(
        样本数量=("股票代码", "count"),
        公司数量=("股票代码", "nunique"),
        数字化关键词总次数=("数字化关键词总次数", "sum"),
        平均标准化词频_每万词=("标准化词频_每万词", "mean"),
    )

    summary_df["阶段"] = stage_name
    summary_df["平均标准化词频_每万词"] = summary_df["平均标准化词频_每万词"].round(4)
    summary_df = summary_df.sort_values("平均标准化词频_每万词", ascending=False)
    return summary_df


def make_province_summary(df, stage_name):
    """省份阶段汇总，后续如果做中国地图可以直接用。"""
    province_df = df.dropna(subset=["所属省份"]).copy()

    summary_df = province_df.groupby(["所属省份", "所属省份代码"], as_index=False).agg(
        样本数量=("股票代码", "count"),
        公司数量=("股票代码", "nunique"),
        数字化关键词总次数=("数字化关键词总次数", "sum"),
        平均标准化词频_每万词=("标准化词频_每万词", "mean"),
    )

    summary_df["阶段"] = stage_name
    summary_df["平均标准化词频_每万词"] = summary_df["平均标准化词频_每万词"].round(4)
    summary_df = summary_df.sort_values("平均标准化词频_每万词", ascending=False)
    return summary_df


def save_city_bar(city_summary_df, stage_name, output_png):
    """保存城市排名图。"""
    top_df = city_summary_df.head(TOP_CITY_NUM).copy()

    if len(top_df) == 0:
        print("没有城市数据，跳过：", output_png)
        return

    top_df = top_df.sort_values("平均标准化词频_每万词", ascending=True)

    plt.figure(figsize=(8, 6))
    plt.barh(top_df["所属城市"], top_df["平均标准化词频_每万词"], color=GREEN)
    plt.title(stage_name + " 城市数字化词频排名")
    plt.xlabel("平均标准化词频（每万词）")
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存城市排名图：", output_png)


def save_province_map(province_summary_df, stage_name, output_html):
    """保存省份中国地图。"""
    if len(province_summary_df) == 0:
        print("没有省份数据，跳过：", output_html)
        return

    map_data = []
    for _, row in province_summary_df.iterrows():
        province = clean_province_name(row["所属省份"])
        value = float(row["平均标准化词频_每万词"])
        map_data.append((province, value))

    max_value = province_summary_df["平均标准化词频_每万词"].max()

    chart = (
        Map()
        .add(stage_name, map_data, "china")
        .set_global_opts(
            title_opts=opts.TitleOpts(title=stage_name + " 省份数字化词频地图"),
            visualmap_opts=opts.VisualMapOpts(max_=float(max_value), range_color=MAP_COLORS),
        )
    )

    chart.render(output_html)
    save_html_as_png(output_html, output_html.replace(".html", ".png"))
    print("已保存省份地图：", output_html)


def save_city_heatmap(city_summary_df, stage_name, output_html):
    """保存城市热力地图。"""
    if len(city_summary_df) == 0:
        print("没有城市数据，跳过：", output_html)
        return

    chart = Geo()
    chart.add_schema(maptype="china")

    geo_data = []
    skipped_cities = []

    for _, row in city_summary_df.iterrows():
        city = clean_city_name(row["所属城市"])
        value = float(row["平均标准化词频_每万词"])

        if city == "" or chart.get_coordinate(city) is None:
            skipped_cities.append(city)
            continue

        geo_data.append((city, value))

    if len(geo_data) == 0:
        print("没有可识别坐标的城市，跳过：", output_html)
        return

    max_value = max([item[1] for item in geo_data])

    chart.add(stage_name, geo_data, type_=ChartType.HEATMAP)
    chart.add(
        stage_name + "城市名称",
        geo_data,
        type_=ChartType.SCATTER,
        symbol_size=3,
        label_opts=opts.LabelOpts(
            is_show=True,
            formatter="{b}",
            font_size=10,
            color="#222222",
            position="right",
        ),
    )
    chart.set_global_opts(
        title_opts=opts.TitleOpts(title=stage_name + " 城市数字化词频热力地图"),
        visualmap_opts=opts.VisualMapOpts(max_=float(max_value), range_color=MAP_COLORS),
        legend_opts=opts.LegendOpts(is_show=False),
    )

    chart.render(output_html)
    save_html_as_png(output_html, output_html.replace(".html", ".png"))
    print("已保存城市热力地图：", output_html)

    if len(skipped_cities) > 0:
        print("城市地图跳过无坐标地点：", sorted(set(skipped_cities)))


# ==================== 主程序 ====================

if not os.path.exists(INPUT_EXCEL):
    print("没有找到合并数据，请先运行：04_merge_industry_city_data.py")
else:
    os.makedirs(YEAR_WORDCLOUD_DIR, exist_ok=True)
    os.makedirs(INDUSTRY_BAR_DIR, exist_ok=True)
    os.makedirs(CITY_BAR_DIR, exist_ok=True)
    os.makedirs(PROVINCE_MAP_DIR, exist_ok=True)
    os.makedirs(CITY_MAP_DIR, exist_ok=True)

    df = pd.read_excel(INPUT_EXCEL, sheet_name=INPUT_SHEET, dtype={"股票代码": str})
    df["年份"] = df["年份"].astype(int)

    word_rows = []

    # 1. 年份阶段词云：总体、前五年、后五年
    stages = [
        ("总体", "2016-2025", df),
        ("前五年", f"{FIRST_START_YEAR}-{FIRST_END_YEAR}", make_stage_df(df, FIRST_START_YEAR, FIRST_END_YEAR)),
        ("后五年", f"{SECOND_START_YEAR}-{SECOND_END_YEAR}", make_stage_df(df, SECOND_START_YEAR, SECOND_END_YEAR)),
    ]

    for stage_type, stage_name, stage_df in stages:
        word_count = make_word_count(stage_df)
        output_png = os.path.join(YEAR_WORDCLOUD_DIR, f"{stage_type}_{stage_name}_词云.png")
        save_wordcloud(word_count, output_png)
        word_rows.extend(word_count_to_rows("年份阶段词云", stage_name, word_count))

    # 2. 行业：每个行业前5大关键词柱状图
    industry_df = df.dropna(subset=["行业名称"])
    industry_count = industry_df.groupby("行业名称")["股票代码"].count()
    top_industries = industry_count.sort_values(ascending=False).head(TOP_INDUSTRY_NUM).index

    for industry in top_industries:
        sub_df = industry_df[industry_df["行业名称"] == industry]
        word_count = make_word_count(sub_df)
        output_png = os.path.join(INDUSTRY_BAR_DIR, f"行业_{safe_filename(industry)}_前5关键词.png")
        save_keyword_bar(word_count, industry + " 前5大关键词", output_png)
        word_rows.extend(word_count_to_rows("行业前5关键词", industry, word_count, top_n=TOP_KEYWORD_NUM))

    # 3. 城市：总体、前五年、后五年排名图，并保存地图数据
    city_summary_list = []
    province_summary_list = []

    for stage_type, stage_name, stage_df in stages:
        city_summary_df = make_city_summary(stage_df, stage_name)
        province_summary_df = make_province_summary(stage_df, stage_name)

        city_summary_list.append(city_summary_df)
        province_summary_list.append(province_summary_df)

        output_png = os.path.join(CITY_BAR_DIR, f"城市排名_{stage_name}.png")
        save_city_bar(city_summary_df, stage_name, output_png)

        province_map_html = os.path.join(PROVINCE_MAP_DIR, f"省份地图_{stage_name}.html")
        city_map_html = os.path.join(CITY_MAP_DIR, f"城市热力地图_{stage_name}.html")
        save_province_map(province_summary_df, stage_name, province_map_html)
        save_city_heatmap(city_summary_df, stage_name, city_map_html)

    city_all_df = pd.concat(city_summary_list, ignore_index=True)
    province_all_df = pd.concat(province_summary_list, ignore_index=True)
    word_df = pd.DataFrame(word_rows)

    # 4. 保存可视化用数据
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        word_df.to_excel(writer, sheet_name="词云和行业关键词", index=False)
        city_all_df.to_excel(writer, sheet_name="城市阶段数据", index=False)
        province_all_df.to_excel(writer, sheet_name="省份地图数据", index=False)

    print("\n可视化分析完成")
    print("年份阶段词云：", YEAR_WORDCLOUD_DIR)
    print("行业前5关键词柱状图：", INDUSTRY_BAR_DIR)
    print("城市阶段排名图：", CITY_BAR_DIR)
    print("省份地图：", PROVINCE_MAP_DIR)
    print("城市热力地图：", CITY_MAP_DIR)
    print("可视化数据汇总：", OUTPUT_EXCEL)
