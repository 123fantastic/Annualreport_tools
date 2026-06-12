# 作者：Winter
# 功能：使用中文 FinBERT 做年报语义分析、情感分析和多维度对比
import os
import re

import matplotlib.pyplot as plt
import pandas as pd
import torch
import torch.nn.functional as F
from digital_keywords import KEYWORDS
from pyecharts import options as opts
from pyecharts.charts import Geo
from pyecharts.globals import ChartType
from transformers import AutoModelForSequenceClassification, AutoTokenizer
from visual_utils import BLUE, GRAY, GREEN, MAP_COLORS, ORANGE, RED, save_html_as_png


# ==================== 配置区：主要改这里 ====================

MERGED_EXCEL = r"年报文件\中证A50近10年数字化词频_合并行业城市.xlsx"
MERGED_SHEET = "合并明细"
REPORT_DIR = "年报文件"

OUTPUT_DIR = r"年报文件\FinBERT分析"
OUTPUT_EXCEL = os.path.join(OUTPUT_DIR, "FinBERT分析结果.xlsx")
CITY_MAP_DIR = os.path.join(OUTPUT_DIR, "城市地图")

# 当前项目使用的是中文 FinBERT 模型
MODEL_NAME = "yiyanghkust/finbert-tone-chinese"

MAX_SENTENCES = 30
MAX_LENGTH = 512
BATCH_SIZE = 16

# 如果只想课堂展示小样本，可以改成 50；None 表示全量
MAX_REPORTS = None


# ==================== 语义基准句子 ====================

DIGITAL_SENTENCE = "企业正在推进数字化转型，使用人工智能、大数据、云计算和区块链等数字技术。"
FINTECH_SENTENCE = "企业正在进行数字技术运用，包括移动互联网、工业互联网、电子商务、移动支付、互联网金融、数字金融和金融科技。"


# ==================== 简单函数区 ====================

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


def clean_text(text):
    """简单清洗文本。"""
    text = str(text)
    text = re.sub(r"\s+", "", text)
    return text


def clean_city_name(name):
    """把城市名称处理成 pyecharts Geo 常用格式。"""
    name = str(name).strip()
    name = name.replace("市", "")
    name = name.replace("特别行政区", "")
    return name


def min_max_100(series):
    """把一列数据统一转成 0-100 的标准化指数，方便不同量纲比较。"""
    series = pd.to_numeric(series, errors="coerce").fillna(0)
    min_value = series.min()
    max_value = series.max()

    if max_value == min_value:
        return series * 0

    return ((series - min_value) / (max_value - min_value) * 100).round(6)


def add_standardized_scores(df):
    """给词典法和 FinBERT 语义结果都加上 0-100 标准化指数。"""
    df["词典标准化指数"] = min_max_100(df["标准化词频_每万词"])
    df["FinBERT综合语义得分标准化指数"] = min_max_100(df["FinBERT综合语义得分"])
    df["数字化语义相似度标准化指数"] = min_max_100(df["数字化语义相似度"])
    df["金融科技语义相似度标准化指数"] = min_max_100(df["金融科技语义相似度"])
    return df


def get_txt_path(stock_code, company_name, year):
    """根据股票代码、公司简称、年份找到 TXT 年报路径。"""
    file_name = f"{stock_code}_{company_name}_{year}.txt"
    return os.path.join(REPORT_DIR, str(year), "txt年报", file_name)


def make_analysis_text(content):
    """优先抽取含关键词的句子；如果没有关键词句，就取开头句子。"""
    content = clean_text(content)
    sentences = re.split(r"[。！？；;]", content)
    sentences = [sentence.strip() for sentence in sentences if sentence.strip()]

    keyword_sentences = []
    for sentence in sentences:
        if any(keyword in sentence for keyword in KEYWORDS):
            keyword_sentences.append(sentence)

    if len(keyword_sentences) == 0:
        selected = sentences[:MAX_SENTENCES]
    else:
        selected = keyword_sentences[:MAX_SENTENCES]

    return "。".join(selected)


def read_reports():
    """读取合并表，并抽取每份年报的分析文本。"""
    df = pd.read_excel(MERGED_EXCEL, sheet_name=MERGED_SHEET, dtype={"股票代码": str})
    df["股票代码"] = df["股票代码"].astype(str).str.zfill(6)
    df["年份"] = df["年份"].astype(int)

    if MAX_REPORTS is not None:
        df = df.head(MAX_REPORTS).copy()

    rows = []

    for i in range(len(df)):
        row = df.loc[i]
        stock_code = row["股票代码"]
        company_name = row["公司简称"]
        year = row["年份"]
        txt_path = get_txt_path(stock_code, company_name, year)

        if not os.path.exists(txt_path):
            continue

        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        analysis_text = make_analysis_text(content)

        if analysis_text == "":
            continue

        rows.append({
            "股票代码": stock_code,
            "公司简称": company_name,
            "年份": year,
            "行业名称": row.get("行业名称", ""),
            "所属省份": row.get("所属省份", ""),
            "所属城市": row.get("所属城市", ""),
            "标准化词频_每万词": row.get("标准化词频_每万词", 0),
            "分析文本": analysis_text,
        })

        if len(rows) % 50 == 0:
            print(f"已经准备 {len(rows)} 份年报文本")

    return pd.DataFrame(rows)


def move_inputs_to_device(inputs, device):
    """把 tokenizer 输出移动到 CPU 或 GPU。"""
    new_inputs = {}

    for key, value in inputs.items():
        new_inputs[key] = value.to(device)

    return new_inputs


def run_finbert(df):
    """运行 FinBERT 情感分析和语义相似度计算。"""
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    print("当前使用设备：", device)

    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
    model = AutoModelForSequenceClassification.from_pretrained(MODEL_NAME)
    model.to(device)
    model.eval()

    label_map = {
        0: "中性",
        1: "积极",
        2: "消极",
    }

    with torch.no_grad():
        base_inputs = tokenizer(
            [DIGITAL_SENTENCE, FINTECH_SENTENCE],
            padding=True,
            truncation=True,
            max_length=MAX_LENGTH,
            return_tensors="pt",
        )
        base_inputs = move_inputs_to_device(base_inputs, device)
        base_outputs = model(**base_inputs, output_hidden_states=True)
        base_embeddings = base_outputs.hidden_states[-1][:, 0, :]
        digital_embedding = base_embeddings[0:1]
        fintech_embedding = base_embeddings[1:2]

    result_rows = []

    for start in range(0, len(df), BATCH_SIZE):
        batch_df = df.iloc[start:start + BATCH_SIZE].copy()
        texts = batch_df["分析文本"].astype(str).tolist()

        inputs = tokenizer(
            texts,
            padding=True,
            truncation=True,
            max_length=MAX_LENGTH,
            return_tensors="pt",
        )
        inputs = move_inputs_to_device(inputs, device)

        with torch.no_grad():
            outputs = model(**inputs, output_hidden_states=True)
            probs = F.softmax(outputs.logits, dim=1)
            embeddings = outputs.hidden_states[-1][:, 0, :]

            digital_sim = F.cosine_similarity(embeddings, digital_embedding, dim=1)
            fintech_sim = F.cosine_similarity(embeddings, fintech_embedding, dim=1)

        for j in range(len(batch_df)):
            row = batch_df.iloc[j]
            prob = probs[j]
            label_id = int(torch.argmax(prob).item())

            result_rows.append({
                "股票代码": row["股票代码"],
                "公司简称": row["公司简称"],
                "年份": row["年份"],
                "行业名称": row["行业名称"],
                "所属省份": row["所属省份"],
                "所属城市": row["所属城市"],
                "标准化词频_每万词": row["标准化词频_每万词"],
                "情感标签": label_map.get(label_id, str(label_id)),
                "中性概率": round(float(prob[0]), 6),
                "积极概率": round(float(prob[1]), 6),
                "消极概率": round(float(prob[2]), 6),
                "数字化语义相似度": round(float(digital_sim[j]), 6),
                "金融科技语义相似度": round(float(fintech_sim[j]), 6),
                "FinBERT综合语义得分": round(float((digital_sim[j] + fintech_sim[j]) / 2), 6),
            })

        print(f"已经完成 FinBERT 分析：{min(start + BATCH_SIZE, len(df))}/{len(df)}")

    return pd.DataFrame(result_rows)


def make_summary(df, group_col):
    """按公司、年份、行业、省份、城市汇总。"""
    summary_df = df.dropna(subset=[group_col]).groupby(group_col, as_index=False).agg(
        年报数量=("股票代码", "count"),
        公司数量=("股票代码", "nunique"),
        平均词典标准化词频=("标准化词频_每万词", "mean"),
        平均词典标准化指数=("词典标准化指数", "mean"),
        平均数字化语义相似度=("数字化语义相似度", "mean"),
        平均金融科技语义相似度=("金融科技语义相似度", "mean"),
        平均FinBERT综合语义得分=("FinBERT综合语义得分", "mean"),
        平均FinBERT标准化指数=("FinBERT综合语义得分标准化指数", "mean"),
        平均中性概率=("中性概率", "mean"),
        平均积极概率=("积极概率", "mean"),
        平均消极概率=("消极概率", "mean"),
    )

    summary_df["FinBERT情感净得分"] = summary_df["平均积极概率"] - summary_df["平均消极概率"]

    number_columns = [
        "平均词典标准化词频",
        "平均词典标准化指数",
        "平均数字化语义相似度",
        "平均金融科技语义相似度",
        "平均FinBERT综合语义得分",
        "平均FinBERT标准化指数",
        "平均中性概率",
        "平均积极概率",
        "平均消极概率",
        "FinBERT情感净得分",
    ]

    for column in number_columns:
        summary_df[column] = summary_df[column].round(6)

    summary_df["词典对比指数_本维度"] = min_max_100(summary_df["平均词典标准化词频"])
    summary_df["FinBERT对比指数_本维度"] = min_max_100(summary_df["平均FinBERT综合语义得分"])

    summary_df = summary_df.sort_values("FinBERT对比指数_本维度", ascending=False)
    return summary_df


def make_stage_summary(df):
    """生成总体、前五年、后五年的 FinBERT 情感汇总。"""
    years = sorted(df["年份"].dropna().astype(int).unique().tolist())
    middle = len(years) // 2
    first_years = years[:middle]
    last_years = years[middle:]

    stage_settings = [
        ("总体", years),
        ("前五年", first_years),
        ("后五年", last_years),
    ]

    rows = []
    for stage_name, stage_years in stage_settings:
        stage_df = df[df["年份"].isin(stage_years)]

        rows.append({
            "阶段": stage_name,
            "年份范围": f"{min(stage_years)}-{max(stage_years)}",
            "年报数量": len(stage_df),
            "公司数量": stage_df["股票代码"].nunique(),
            "平均词典标准化词频": round(stage_df["标准化词频_每万词"].mean(), 6),
            "平均FinBERT综合语义得分": round(stage_df["FinBERT综合语义得分"].mean(), 6),
            "平均积极概率": round(stage_df["积极概率"].mean(), 6),
            "平均中性概率": round(stage_df["中性概率"].mean(), 6),
            "平均消极概率": round(stage_df["消极概率"].mean(), 6),
            "FinBERT情感净得分": round(stage_df["积极概率"].mean() - stage_df["消极概率"].mean(), 6),
        })

    return pd.DataFrame(rows)


def save_semantic_year_line(year_df, output_png):
    """年份维度：FinBERT语义得分与词典法对比。"""
    year_df = year_df.sort_values("年份")

    plt.figure(figsize=(10, 5))
    plt.plot(year_df["年份"], year_df["FinBERT对比指数_本维度"], marker="o", color=BLUE, label="FinBERT语义指数")
    plt.plot(year_df["年份"], year_df["词典对比指数_本维度"], marker="s", color=ORANGE, label="词典法指数")
    plt.title("年份维度：FinBERT语义指数与词典法对比")
    plt.xlabel("年份")
    plt.ylabel("标准化指数（0-100）")
    plt.legend()
    plt.grid(alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存年份语义图：", output_png)


def save_bar(df, group_col, value_col, title, xlabel, output_png, top_n=15, color=BLUE):
    """横向条形图。"""
    if len(df) == 0:
        print("没有数据，跳过：", output_png)
        return

    plot_df = df.sort_values(value_col, ascending=False).head(top_n)
    plot_df = plot_df.sort_values(value_col, ascending=True)

    plt.figure(figsize=(8, 6))
    plt.barh(plot_df[group_col].astype(str), plot_df[value_col], color=color)
    plt.title(title)
    plt.xlabel(xlabel)
    plt.grid(axis="x", alpha=0.25)
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存条形图：", output_png)


def save_finbert_year_line(year_df, output_png):
    """年份维度 FinBERT 情感概率折线图。"""
    year_df = year_df.sort_values("年份").reset_index(drop=True)

    plt.figure(figsize=(10, 5))
    plt.plot(year_df["年份"], year_df["平均积极概率"], marker="o", label="积极概率", color=GREEN)
    plt.plot(year_df["年份"], year_df["平均中性概率"], marker="s", label="中性概率", color=GRAY)
    plt.plot(year_df["年份"], year_df["平均消极概率"], marker="^", label="消极概率", color=RED)
    plt.title("年份维度：FinBERT情感概率趋势")
    plt.xlabel("年份")
    plt.ylabel("平均概率")
    plt.legend()
    plt.grid(alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存 FinBERT 年份情感图：", output_png)


def save_finbert_stage_pies(stage_df, output_png):
    """总体、前五年、后五年 FinBERT 情感概率饼图。"""
    labels = ["积极", "中性", "消极"]
    colors = [GREEN, GRAY, RED]

    fig, axes = plt.subplots(1, len(stage_df), figsize=(12, 4))

    if len(stage_df) == 1:
        axes = [axes]

    for i in range(len(stage_df)):
        row = stage_df.iloc[i]
        values = [
            row["平均积极概率"],
            row["平均中性概率"],
            row["平均消极概率"],
        ]

        axes[i].pie(
            values,
            labels=labels,
            colors=colors,
            autopct="%1.1f%%",
            startangle=90,
            textprops={"fontsize": 9},
        )
        axes[i].set_title(f"{row['阶段']}（{row['年份范围']}）")

    fig.suptitle("阶段维度：FinBERT情感概率结构", fontsize=15)
    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()

    print("已保存 FinBERT 阶段饼图：", output_png)


def save_city_map(df, city_col, value_col, title, output_html, series_name):
    """城市维度热力地图，同时保存静态 PNG。"""
    if len(df) == 0:
        print("没有城市数据，跳过：", output_html)
        return

    chart = Geo()
    chart.add_schema(maptype="china")

    geo_data = []
    skipped_cities = []

    for _, row in df.iterrows():
        city = clean_city_name(row[city_col])
        value = float(row[value_col])

        if city == "" or chart.get_coordinate(city) is None:
            skipped_cities.append(city)
            continue

        geo_data.append((city, value))

    if len(geo_data) == 0:
        print("没有可识别坐标的城市，跳过：", output_html)
        return

    values = [item[1] for item in geo_data]
    chart.add(series_name, geo_data, type_=ChartType.HEATMAP)
    chart.add(
        series_name + "城市名称",
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
        title_opts=opts.TitleOpts(title=title),
        visualmap_opts=opts.VisualMapOpts(
            min_=float(min(values)),
            max_=float(max(values)),
            range_color=MAP_COLORS,
        ),
        legend_opts=opts.LegendOpts(is_show=False),
        tooltip_opts=opts.TooltipOpts(formatter="{b}: {c}"),
    )

    chart.render(output_html)
    save_html_as_png(output_html, output_html.replace(".html", ".png"))
    print("已保存城市地图：", output_html)

    if len(skipped_cities) > 0:
        print("城市地图跳过无坐标地点：", sorted(set(skipped_cities)))


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(CITY_MAP_DIR, exist_ok=True)

    report_df = read_reports()

    if len(report_df) == 0:
        print("没有读取到年报文本")
        return

    result_df = run_finbert(report_df)
    result_df = add_standardized_scores(result_df)

    company_df = make_summary(result_df, "公司简称")
    year_df = make_summary(result_df, "年份")
    industry_df = make_summary(result_df, "行业名称")
    province_df = make_summary(result_df, "所属省份")
    city_df = make_summary(result_df, "所属城市")
    stage_df = make_stage_summary(result_df)

    ppt_explain_df = pd.DataFrame([
        ["方法", "当前脚本使用中文 FinBERT，同时进行语义向量分析和情感分类分析。"],
        ["文本截取", "每份年报优先抽取含企业数字化转型关键词的句子，避免整篇年报过长导致模型无法处理。"],
        ["FinBERT综合语义得分", "数字化语义相似度和金融科技/数字技术运用语义相似度的平均值。"],
        ["FinBERT情感分析", "使用积极、中性、消极三个概率描述年报相关句子的情感倾向，并用积极概率减消极概率得到情感净得分。"],
        ["多维度对比", "结果按公司、年份、行业、省份、城市、阶段汇总，便于和词典法以及后续大语言模型结果对比。"],
        ["标准化处理", "由于FinBERT语义得分和词典词频不在同一量纲，脚本将两者转成0-100标准化指数后再进行趋势对比。"],
    ], columns=["项目", "说明"])

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="年报FinBERT明细", index=False)
        company_df.to_excel(writer, sheet_name="公司维度", index=False)
        year_df.to_excel(writer, sheet_name="年份维度", index=False)
        industry_df.to_excel(writer, sheet_name="行业维度", index=False)
        province_df.to_excel(writer, sheet_name="省份维度", index=False)
        city_df.to_excel(writer, sheet_name="城市维度", index=False)
        stage_df.to_excel(writer, sheet_name="阶段维度", index=False)
        ppt_explain_df.to_excel(writer, sheet_name="PPT说明", index=False)

    save_semantic_year_line(year_df, os.path.join(OUTPUT_DIR, "年份维度_FinBERT语义指数与词典法对比.png"))
    save_bar(company_df, "公司简称", "FinBERT对比指数_本维度", "公司维度：FinBERT语义指数", "FinBERT标准化指数（0-100）", os.path.join(OUTPUT_DIR, "公司维度_FinBERT语义指数.png"), top_n=20, color=BLUE)
    save_bar(industry_df, "行业名称", "FinBERT对比指数_本维度", "行业维度：FinBERT语义指数", "FinBERT标准化指数（0-100）", os.path.join(OUTPUT_DIR, "行业维度_FinBERT语义指数.png"), color=BLUE)
    save_bar(province_df, "所属省份", "FinBERT对比指数_本维度", "省份维度：FinBERT语义指数", "FinBERT标准化指数（0-100）", os.path.join(OUTPUT_DIR, "省份维度_FinBERT语义指数.png"), color=BLUE)

    save_finbert_year_line(year_df, os.path.join(OUTPUT_DIR, "年份维度_FinBERT情感概率趋势.png"))
    save_finbert_stage_pies(stage_df, os.path.join(OUTPUT_DIR, "阶段维度_FinBERT情感概率饼图.png"))
    save_bar(industry_df, "行业名称", "FinBERT情感净得分", "行业维度：FinBERT情感净得分", "积极概率 - 消极概率", os.path.join(OUTPUT_DIR, "行业维度_FinBERT情感净得分.png"), color=GREEN)
    save_bar(province_df, "所属省份", "FinBERT情感净得分", "省份维度：FinBERT情感净得分", "积极概率 - 消极概率", os.path.join(OUTPUT_DIR, "省份维度_FinBERT情感净得分.png"), color=GREEN)

    save_city_map(city_df, "所属城市", "FinBERT对比指数_本维度", "城市维度：FinBERT语义指数地图", os.path.join(CITY_MAP_DIR, "城市维度_FinBERT语义指数地图.html"), "FinBERT语义")
    save_city_map(city_df, "所属城市", "FinBERT情感净得分", "城市维度：FinBERT情感净得分地图", os.path.join(CITY_MAP_DIR, "城市维度_FinBERT情感净得分地图.html"), "FinBERT情感")

    print("\nFinBERT 分析完成")
    print("结果 Excel：", OUTPUT_EXCEL)
    print("结果文件夹：", OUTPUT_DIR)


if __name__ == "__main__":
    main()
