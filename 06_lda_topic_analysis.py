# 作者：Winter
# 功能：对年报文本进行 LDA 主题分析，输出主题关键词、每份年报主主题、年度主题变化

import os
import re

import jieba
import matplotlib.pyplot as plt
import pandas as pd
import requests
from digital_keywords import DIGITAL_KEYWORD_GROUPS, KEYWORDS
from pyecharts import options as opts
from pyecharts.charts import Geo, Map
from pyecharts.globals import ChartType
from sklearn.decomposition import LatentDirichletAllocation
from sklearn.feature_extraction.text import CountVectorizer
from visual_utils import MAP_COLORS, TOPIC_COLORS, save_html_as_png


# ==================== 配置区：主要改这里 ====================

REPORT_DIR = "年报文件"
OUTPUT_DIR = r"年报文件\LDA主题分析"
OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, "中证A50近10年LDA主题分析结果.xlsx")
PROVINCE_TOPIC_MAP_DIR = os.path.join(OUTPUT_DIR, "省份主题地图")
CITY_TOPIC_MAP_DIR = os.path.join(OUTPUT_DIR, "城市主题热力地图")
MAIN_TOPIC_CHART_DIR = os.path.join(OUTPUT_DIR, "主主题分布图")

# 如果存在合并后的行业城市表，就把行业城市信息也带到结果里
MERGED_EXCEL = r"年报文件\中证A50近10年数字化词频_合并行业城市.xlsx"
MERGED_SHEET = "合并明细"

# 停用词来源：https://github.com/123fantastic/Stopwords
STOPWORDS_FILE = r"data\stopwords\stopwords_cn.txt"
STOPWORDS_URL = "https://raw.githubusercontent.com/123fantastic/Stopwords/main/stopwords_cn.txt"

# LDA 参数
TOPIC_NUM = 5
TOP_WORD_NUM = 20
MAX_FEATURES = 2000
MIN_DF = 3
MAX_DF = 0.85


for word in KEYWORDS:
    jieba.add_word(word)


# ==================== 简单函数区 ====================

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


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


def clean_text(text):
    """保留中文、英文和数字，去掉大部分特殊符号。"""
    text = re.sub(r"\s+", "", str(text))
    text = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9]", "", text)
    return text


def get_stopwords():
    """读取停用词；如果本地没有，就从 GitHub 下载。"""
    if not os.path.exists(STOPWORDS_FILE):
        try:
            os.makedirs(os.path.dirname(STOPWORDS_FILE), exist_ok=True)
            res = requests.get(STOPWORDS_URL, timeout=30)
            res.encoding = "utf-8"

            with open(STOPWORDS_FILE, "w", encoding="utf-8") as f:
                f.write(res.text)

            print("停用词已下载：", STOPWORDS_FILE)

        except Exception as e:
            print("停用词下载失败，暂时不使用停用词：", e)
            return set()

    with open(STOPWORDS_FILE, "r", encoding="utf-8", errors="ignore") as f:
        stopwords = set()

        for line in f:
            word = line.strip()
            if word:
                stopwords.add(word)

    # 年报中很常见，但对主题解释帮助不大的词
    stopwords.update([
        "公司", "本公司", "集团", "报告", "年度", "年报", "股份", "有限公司",
        "情况", "业务", "发展", "管理", "主要", "相关", "进行", "实现",
        "提升", "推进", "加强", "建设", "产品", "服务", "市场", "客户",
        "经营", "投资", "项目", "收入", "资产", "利润", "风险",
    ])

    print("停用词数量：", len(stopwords))
    return stopwords


def get_file_info(filename):
    """从文件名中拆出股票代码、公司简称、年份。"""
    name_without_ext = os.path.splitext(filename)[0]
    parts = name_without_ext.split("_")

    if len(parts) >= 3:
        return parts[0], parts[1], int(parts[2])

    return "", name_without_ext, 0


def cut_words(text, stopwords):
    """清洗并分词，给 LDA 使用。"""
    text = clean_text(text)
    words = jieba.cut(text, HMM=False)
    result = []

    for word in words:
        word = word.strip()

        if not word:
            continue

        if len(word) <= 1 and word not in ["AI", "5G"]:
            continue

        if word in stopwords and word not in KEYWORDS:
            continue

        if re.fullmatch(r"\d+", word):
            continue

        result.append(word)

    return result


def read_txt_reports(stopwords):
    """读取所有 TXT 年报，返回文档信息和分词文本。"""
    rows = []

    for root, dirs, files in os.walk(REPORT_DIR):
        dirs.sort()
        files.sort()

        if "txt年报" not in root:
            continue

        for filename in files:
            if not filename.endswith(".txt"):
                continue

            txt_path = os.path.join(root, filename)
            stock_code, company_name, year = get_file_info(filename)

            with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()

            words = cut_words(content, stopwords)

            if len(words) == 0:
                continue

            rows.append({
                "股票代码": stock_code,
                "公司简称": company_name,
                "年份": year,
                "文件名": filename,
                "分词文本": " ".join(words),
                "分词数量": len(words),
            })

            if len(rows) % 50 == 0:
                print(f"已经读取并分词 {len(rows)} 份")

    return pd.DataFrame(rows)


def add_industry_city(doc_df):
    """如果第四步合并表存在，就补充行业和城市信息。"""
    if not os.path.exists(MERGED_EXCEL):
        return doc_df

    try:
        attr_df = pd.read_excel(MERGED_EXCEL, sheet_name=MERGED_SHEET, dtype={"股票代码": str})
        attr_df["年份"] = attr_df["年份"].astype(int)

        keep_columns = [
            "股票代码",
            "年份",
            "行业名称",
            "所属省份",
            "所属城市",
            "属性来源年份",
            "匹配方式",
        ]

        attr_df = attr_df[keep_columns]
        doc_df = pd.merge(doc_df, attr_df, on=["股票代码", "年份"], how="left")

    except Exception as e:
        print("行业城市信息合并失败，不影响 LDA：", e)

    return doc_df


def build_topic_words(lda_model, feature_names):
    """整理每个主题的前 N 个关键词。"""
    rows = []

    for topic_id, topic in enumerate(lda_model.components_):
        top_index = topic.argsort()[::-1][:TOP_WORD_NUM]

        for rank, word_index in enumerate(top_index, start=1):
            rows.append({
                "主题编号": topic_id,
                "排名": rank,
                "关键词": feature_names[word_index],
                "权重": round(float(topic[word_index]), 6),
            })

    return pd.DataFrame(rows)


def make_topic_names(topic_words_df):
    """根据每个主题前20个关键词生成简短主题名称。

    课堂版不调用大模型，先按主题词中的行业特征命名；
    如果主题词正好落在数字化词典里，再按词典类别命名。
    主题名称统一控制为 6 个汉字，方便图表展示。
    """
    candidate_names = {
        "人工智能技术": "智能算法主题",
        "大数据技术": "数据挖掘主题",
        "云计算技术": "云网计算主题",
        "区块链技术": "链上金融主题",
        "数字技术运用": "数字场景主题",
    }
    industry_rules = [
        ("银行金融主题", ["本行", "分行", "贷款", "垫款", "工商银行", "农业银行", "金融服务"]),
        ("保险医疗主题", ["平安", "保险", "太保", "寿险", "保费", "医疗", "诊断"]),
        ("新能源车主题", ["比亚迪", "汽车", "新能源", "精密", "半导体", "电力", "科技股份"]),
        ("装备制造主题", ["船舶", "装备", "重工", "电工", "汇川", "造船", "船舶工业"]),
        ("消费制造主题", ["美的", "食品", "伊利", "海天", "农牧", "饲料", "养殖"]),
    ]
    fallback_names = ["综合转型主题", "产业升级主题", "智能应用主题", "数字运营主题", "技术融合主题"]

    used_names = set()
    topic_name_rows = []

    for topic_id in sorted(topic_words_df["主题编号"].unique()):
        sub_df = topic_words_df[topic_words_df["主题编号"] == topic_id].head(20)
        words = sub_df["关键词"].tolist()
        word_text = " ".join(words)

        rule_scores = {}
        for rule_name, rule_words in industry_rules:
            score = 0
            for rank, word in enumerate(words, start=1):
                if any(rule_word in word_text or rule_word in word for rule_word in rule_words):
                    score += 21 - rank
            rule_scores[rule_name] = score

        group_scores = {}
        for group_name, group_words in DIGITAL_KEYWORD_GROUPS.items():
            score = 0
            for rank, word in enumerate(words, start=1):
                if word in group_words:
                    score += 21 - rank
            group_scores[group_name] = score

        best_group = max(group_scores, key=group_scores.get)
        best_rule = max(rule_scores, key=rule_scores.get)

        if rule_scores[best_rule] > 0:
            topic_name = best_rule
            main_type = "行业主题"
        elif group_scores[best_group] > 0:
            topic_name = candidate_names[best_group]
            main_type = best_group
        else:
            topic_name = next((name for name in fallback_names if name not in used_names), "综合转型主题")
            main_type = "综合主题"

        if topic_name in used_names:
            topic_name = next((name for name in fallback_names if name not in used_names), topic_name)
            main_type = "综合主题"
        used_names.add(topic_name)

        topic_name_rows.append({
            "主题编号": topic_id,
            "主题名称": topic_name,
            "主要类别": main_type,
            "前20主题词": "、".join(words),
        })

    return pd.DataFrame(topic_name_rows)


def add_topic_names(topic_words_df, topic_name_df):
    """给主题关键词表补充主题名称。"""
    return pd.merge(topic_words_df, topic_name_df[["主题编号", "主题名称"]], on="主题编号", how="left")


def add_topic_name_to_doc(doc_topic_df, topic_name_df):
    """给每份年报的主主题补充主题名称。"""
    name_dict = dict(zip(topic_name_df["主题编号"], topic_name_df["主题名称"]))
    doc_topic_df["主主题名称"] = doc_topic_df["主主题"].map(name_dict)
    return doc_topic_df


def build_doc_topics(doc_df, doc_topic_matrix):
    """整理每份年报的主题概率。"""
    result_df = doc_df.drop(columns=["分词文本"]).copy()

    for topic_id in range(TOPIC_NUM):
        result_df[f"主题{topic_id}概率"] = doc_topic_matrix[:, topic_id].round(6)

    result_df["主主题"] = doc_topic_matrix.argmax(axis=1)
    result_df["主主题概率"] = doc_topic_matrix.max(axis=1).round(6)

    return result_df


def build_year_topic(doc_topic_df):
    """按年份计算平均主题强度。"""
    topic_columns = [f"主题{i}概率" for i in range(TOPIC_NUM)]
    year_topic_df = doc_topic_df.groupby("年份", as_index=False)[topic_columns].mean()

    for column in topic_columns:
        year_topic_df[column] = year_topic_df[column].round(6)

    return year_topic_df


def build_region_topic(doc_topic_df, group_columns):
    """按省份或城市计算每个主题的平均概率。"""
    topic_columns = [f"主题{i}概率" for i in range(TOPIC_NUM)]
    region_df = doc_topic_df.dropna(subset=group_columns).copy()
    region_topic_df = region_df.groupby(group_columns, as_index=False)[topic_columns].mean()

    for column in topic_columns:
        region_topic_df[column] = region_topic_df[column].round(6)

    return region_topic_df


def save_year_topic_line(year_topic_df, topic_name_df, output_png):
    """保存年度主题变化折线图。"""
    plt.figure(figsize=(10, 6))
    name_dict = dict(zip(topic_name_df["主题编号"], topic_name_df["主题名称"]))

    for topic_id in range(TOPIC_NUM):
        plt.plot(
            year_topic_df["年份"],
            year_topic_df[f"主题{topic_id}概率"],
            marker="o",
            color=TOPIC_COLORS[topic_id % len(TOPIC_COLORS)],
            label=name_dict.get(topic_id, f"主题{topic_id}"),
        )

    plt.title("LDA主题强度年度变化")
    plt.xlabel("年份")
    plt.ylabel("平均主题概率")
    plt.legend()
    plt.grid(alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存年度主题变化图：", output_png)


def save_year_topic_heatmap(year_topic_df, topic_name_df, output_png):
    """保存年度主题热力图。"""
    topic_columns = [f"主题{i}概率" for i in range(TOPIC_NUM)]
    data = year_topic_df[topic_columns].values
    name_dict = dict(zip(topic_name_df["主题编号"], topic_name_df["主题名称"]))

    plt.figure(figsize=(9, 6))
    plt.imshow(data, aspect="auto", cmap="YlOrRd")
    plt.colorbar(label="平均主题概率")
    plt.xticks(range(TOPIC_NUM), [name_dict.get(i, f"主题{i}") for i in range(TOPIC_NUM)])
    plt.yticks(range(len(year_topic_df)), year_topic_df["年份"].tolist())
    plt.title("年份-主题强度热力图")
    plt.xlabel("主题")
    plt.ylabel("年份")
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存年度主题热力图：", output_png)


def save_topic_province_maps(province_topic_df, topic_name_df):
    """每个主题生成一张省份地图。"""
    os.makedirs(PROVINCE_TOPIC_MAP_DIR, exist_ok=True)
    name_dict = dict(zip(topic_name_df["主题编号"], topic_name_df["主题名称"]))

    for topic_id in range(TOPIC_NUM):
        topic_col = f"主题{topic_id}概率"
        topic_name = name_dict.get(topic_id, f"主题{topic_id}")
        map_data = []

        for _, row in province_topic_df.iterrows():
            province = clean_province_name(row["所属省份"])
            value = float(row[topic_col])
            map_data.append((province, value))

        if len(map_data) == 0:
            continue

        max_value = max([item[1] for item in map_data])
        output_html = os.path.join(PROVINCE_TOPIC_MAP_DIR, f"{topic_name}_省份地图.html")

        chart = (
            Map()
            .add(topic_name, map_data, "china")
            .set_global_opts(
                title_opts=opts.TitleOpts(title=f"LDA主题：{topic_name} 省份分布地图"),
                visualmap_opts=opts.VisualMapOpts(max_=float(max_value), range_color=MAP_COLORS),
            )
        )

        chart.render(output_html)
        save_html_as_png(output_html, output_html.replace(".html", ".png"))
        print("已保存主题省份地图：", output_html)


def save_topic_city_maps(city_topic_df, topic_name_df):
    """每个主题生成一张城市热力地图。"""
    os.makedirs(CITY_TOPIC_MAP_DIR, exist_ok=True)
    name_dict = dict(zip(topic_name_df["主题编号"], topic_name_df["主题名称"]))

    for topic_id in range(TOPIC_NUM):
        topic_col = f"主题{topic_id}概率"
        topic_name = name_dict.get(topic_id, f"主题{topic_id}")
        chart = Geo()
        chart.add_schema(maptype="china")

        geo_data = []
        skipped_cities = []

        for _, row in city_topic_df.iterrows():
            city = clean_city_name(row["所属城市"])
            value = float(row[topic_col])

            if city == "" or chart.get_coordinate(city) is None:
                skipped_cities.append(city)
                continue

            geo_data.append((city, value))

        if len(geo_data) == 0:
            continue

        max_value = max([item[1] for item in geo_data])
        output_html = os.path.join(CITY_TOPIC_MAP_DIR, f"{topic_name}_城市热力地图.html")

        chart.add(topic_name, geo_data, type_=ChartType.HEATMAP)
        chart.add(
            topic_name + "城市名称",
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
            title_opts=opts.TitleOpts(title=f"LDA主题：{topic_name} 城市热力地图"),
            visualmap_opts=opts.VisualMapOpts(max_=float(max_value), range_color=MAP_COLORS),
            legend_opts=opts.LegendOpts(is_show=False),
        )

        chart.render(output_html)
        save_html_as_png(output_html, output_html.replace(".html", ".png"))
        print("已保存主题城市热力地图：", output_html)

        if len(skipped_cities) > 0:
            print("主题城市地图跳过无坐标地点：", sorted(set(skipped_cities)))


def build_main_topic_summary(doc_topic_df, group_col):
    """按年份、行业、城市统计主主题分布。"""
    count_df = doc_topic_df.dropna(subset=[group_col]).groupby(
        [group_col, "主主题", "主主题名称"],
        as_index=False,
    ).agg(
        年报数量=("股票代码", "count"),
        平均主主题概率=("主主题概率", "mean"),
    )

    count_df["平均主主题概率"] = count_df["平均主主题概率"].round(6)

    total_df = count_df.groupby(group_col, as_index=False)["年报数量"].sum()
    total_df = total_df.rename(columns={"年报数量": "分组年报总数"})

    count_df = pd.merge(count_df, total_df, on=group_col, how="left")
    count_df["主题占比"] = (count_df["年报数量"] / count_df["分组年报总数"]).round(6)

    return count_df


def save_main_topic_stacked_bar(summary_df, group_col, title, output_png, topic_name_df, top_n=None):
    """保存主主题分布堆叠柱状图。"""
    if len(summary_df) == 0:
        print("没有主主题分布数据，跳过：", output_png)
        return

    plot_df = summary_df.copy()

    if top_n is not None:
        group_total = (
            plot_df.groupby(group_col, as_index=False)["分组年报总数"]
            .max()
            .sort_values("分组年报总数", ascending=False)
            .head(top_n)
        )
        keep_groups = group_total[group_col].tolist()
        plot_df = plot_df[plot_df[group_col].isin(keep_groups)]

    pivot_df = plot_df.pivot_table(
        index=group_col,
        columns="主主题名称",
        values="主题占比",
        fill_value=0,
    )

    if group_col != "年份":
        pivot_df["__sum__"] = pivot_df.sum(axis=1)
        pivot_df = pivot_df.sort_values("__sum__", ascending=True).drop(columns=["__sum__"])
    else:
        pivot_df = pivot_df.sort_index()

    color_list = TOPIC_COLORS[:len(pivot_df.columns)]
    ax = pivot_df.plot(kind="bar", stacked=True, figsize=(10, 6), color=color_list)
    ax.set_title(title)
    ax.set_xlabel(group_col)
    ax.set_ylabel("主主题占比")
    ax.legend(title="主主题", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存主主题分布图：", output_png)


# ==================== 主程序 ====================

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(PROVINCE_TOPIC_MAP_DIR, exist_ok=True)
os.makedirs(CITY_TOPIC_MAP_DIR, exist_ok=True)
os.makedirs(MAIN_TOPIC_CHART_DIR, exist_ok=True)

stopwords = get_stopwords()
doc_df = read_txt_reports(stopwords)
doc_df = add_industry_city(doc_df)

if len(doc_df) == 0:
    print("没有读取到可用于 LDA 的年报文本")
else:
    print("开始训练 LDA，文档数量：", len(doc_df))

    vectorizer = CountVectorizer(
        tokenizer=str.split,
        token_pattern=None,
        max_features=MAX_FEATURES,
        min_df=MIN_DF,
        max_df=MAX_DF,
    )

    doc_word_matrix = vectorizer.fit_transform(doc_df["分词文本"])
    feature_names = vectorizer.get_feature_names_out()

    lda_model = LatentDirichletAllocation(
        n_components=TOPIC_NUM,
        max_iter=20,
        learning_method="batch",
        random_state=2025,
    )

    doc_topic_matrix = lda_model.fit_transform(doc_word_matrix)

    topic_words_df = build_topic_words(lda_model, feature_names)
    topic_name_df = make_topic_names(topic_words_df)
    topic_words_df = add_topic_names(topic_words_df, topic_name_df)
    doc_topic_df = build_doc_topics(doc_df, doc_topic_matrix)
    doc_topic_df = add_topic_name_to_doc(doc_topic_df, topic_name_df)
    year_topic_df = build_year_topic(doc_topic_df)
    province_topic_df = build_region_topic(doc_topic_df, ["所属省份"])
    city_topic_df = build_region_topic(doc_topic_df, ["所属城市"])
    year_main_topic_df = build_main_topic_summary(doc_topic_df, "年份")
    industry_main_topic_df = build_main_topic_summary(doc_topic_df, "行业名称")
    city_main_topic_df = build_main_topic_summary(doc_topic_df, "所属城市")

    ppt_explain_df = pd.DataFrame([
        ["方法", "先对年报文本进行清洗和 jieba 分词，再使用 CountVectorizer 构建词频矩阵，最后用 LDA 识别潜在主题。"],
        ["主题数量", f"本脚本暂时设置为 {TOPIC_NUM} 个主题，课堂展示时便于解释。"],
        ["主题关键词", f"每个主题输出前 {TOP_WORD_NUM} 个高权重词，用来概括主题含义。"],
        ["主题命名", "脚本根据前20个主题词所属的数字化转型关键词类别，自动生成长度一致、便于展示的主题名称。"],
        ["主主题", "每份年报取概率最高的主题作为主主题。"],
        ["注意", "LDA 是无监督模型，主题名称是辅助解释，正式论文中可结合前20个主题词人工复核。"],
    ], columns=["项目", "说明"])

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        topic_name_df.to_excel(writer, sheet_name="主题命名", index=False)
        topic_words_df.to_excel(writer, sheet_name="主题关键词", index=False)
        doc_topic_df.to_excel(writer, sheet_name="年报主题概率", index=False)
        year_topic_df.to_excel(writer, sheet_name="年度主题强度", index=False)
        province_topic_df.to_excel(writer, sheet_name="省份主题强度", index=False)
        city_topic_df.to_excel(writer, sheet_name="城市主题强度", index=False)
        year_main_topic_df.to_excel(writer, sheet_name="年份主主题分布", index=False)
        industry_main_topic_df.to_excel(writer, sheet_name="行业主主题分布", index=False)
        city_main_topic_df.to_excel(writer, sheet_name="城市主主题分布", index=False)
        ppt_explain_df.to_excel(writer, sheet_name="PPT说明", index=False)

    save_year_topic_line(year_topic_df, topic_name_df, os.path.join(OUTPUT_DIR, "年度主题强度折线图.png"))
    save_year_topic_heatmap(year_topic_df, topic_name_df, os.path.join(OUTPUT_DIR, "年度主题强度热力图.png"))
    save_topic_province_maps(province_topic_df, topic_name_df)
    save_topic_city_maps(city_topic_df, topic_name_df)
    save_main_topic_stacked_bar(
        year_main_topic_df,
        "年份",
        "年份维度：年报主主题分布",
        os.path.join(MAIN_TOPIC_CHART_DIR, "年份维度_主主题分布.png"),
        topic_name_df,
    )
    save_main_topic_stacked_bar(
        industry_main_topic_df,
        "行业名称",
        "行业维度：年报主主题分布",
        os.path.join(MAIN_TOPIC_CHART_DIR, "行业维度_主主题分布.png"),
        topic_name_df,
        top_n=15,
    )
    save_main_topic_stacked_bar(
        city_main_topic_df,
        "所属城市",
        "城市维度：年报主主题分布",
        os.path.join(MAIN_TOPIC_CHART_DIR, "城市维度_主主题分布.png"),
        topic_name_df,
        top_n=15,
    )

    print("\nLDA 主题分析完成")
    print("结果 Excel：", OUTPUT_EXCEL)
    print("结果文件夹：", OUTPUT_DIR)
    print("省份主题地图：", PROVINCE_TOPIC_MAP_DIR)
    print("城市主题热力地图：", CITY_TOPIC_MAP_DIR)
    print("主主题分布图：", MAIN_TOPIC_CHART_DIR)
