# 作者：Winter
# 功能：使用 DeepSeek R1 对年报做语义级别分析，并与词典法、FinBERT 结果对比

import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import requests
from digital_keywords import KEYWORDS
from pyecharts import options as opts
from pyecharts.charts import Geo
from pyecharts.globals import ChartType
from visual_utils import BLUE, GRAY, GREEN, MAP_COLORS, ORANGE, RED, save_html_as_png


# ==================== 配置区：主要改这里 ====================

REPORT_DIR = Path("年报文件")
OUTPUT_DIR = REPORT_DIR / "DeepSeek分析"
CHART_DIR = OUTPUT_DIR / "可视化图"
CITY_MAP_DIR = OUTPUT_DIR / "城市地图"

OUTPUT_EXCEL = OUTPUT_DIR / "DeepSeekR1分析结果.xlsx"
CHECKPOINT_CSV = OUTPUT_DIR / "DeepSeekR1分析_checkpoint.csv"

# 全量分析设为 None；课堂小样本可以改成 20、50 等
MAX_REPORTS = None

# DeepSeek R1 每份年报都要请求 API，适当并发可以加速；如果遇到限流，改小一点
MAX_WORKERS = 3
SAVE_EVERY = 10
PRINT_EVERY = 50

# 每份年报只发送与数字化关键词相关的句子，避免把整份年报都发给 API
MAX_SENTENCES = 25
MAX_CHARS = 6500


# ==================== 基础设置 ====================

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


def load_env(env_path=".env"):
    """读取 .env 文件，不需要额外安装 python-dotenv。"""
    env = {}
    path = Path(env_path)
    if not path.exists():
        raise FileNotFoundError("没有找到 .env 文件，请先配置 DEEPSEEK_API_KEY")

    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def get_deepseek_config():
    env = load_env()
    api_key = env.get("DEEPSEEK_API_KEY", "")
    if not api_key or api_key == "your_deepseek_api_key_here":
        raise ValueError("DEEPSEEK_API_KEY 还没有设置，请先修改 .env 文件")

    return {
        "api_key": api_key,
        "model": env.get("DEEPSEEK_MODEL", "deepseek-reasoner"),
        "base_url": env.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/"),
    }


def find_excel(keyword):
    """按文件名关键词自动寻找 Excel，避免中文路径手写出错。"""
    paths = []
    for path in Path.cwd().rglob("*.xlsx"):
        if keyword in path.name:
            paths.append(path)
    if len(paths) == 0:
        raise FileNotFoundError(f"没有找到包含 {keyword} 的 Excel 文件")
    return sorted(paths, key=lambda p: len(str(p)))[0]


def clean_text(text):
    text = str(text)
    text = re.sub(r"\s+", "", text)
    return text


def clean_city_name(name):
    name = str(name).strip()
    name = name.replace("市", "")
    name = name.replace("特别行政区", "")
    return name


def min_max_100(series):
    series = pd.to_numeric(series, errors="coerce").fillna(0)
    min_value = series.min()
    max_value = series.max()
    if max_value == min_value:
        return series * 0
    return ((series - min_value) / (max_value - min_value) * 100).round(6)


def make_level(score):
    if pd.isna(score):
        return "不明显"
    if score >= 70:
        return "实质性转型"
    if score >= 40:
        return "表层叙事"
    return "不明显"


def find_txt_path(stock_code, company_name, year):
    code = str(stock_code).zfill(6)
    folder = REPORT_DIR / str(int(year)) / "txt年报"
    exact_path = folder / f"{code}_{company_name}_{int(year)}.txt"
    if exact_path.exists():
        return exact_path

    matches = list(folder.glob(f"{code}_*_{int(year)}.txt"))
    if matches:
        return matches[0]
    return None


def make_analysis_text(content):
    """优先抽取包含数字化关键词的句子；没有关键词时取开头句子。"""
    content = clean_text(content)
    sentences = re.split(r"[。！？；;]", content)
    sentences = [sentence.strip() for sentence in sentences if sentence.strip()]

    keyword_sentences = []
    for sentence in sentences:
        if any(keyword in sentence for keyword in KEYWORDS):
            keyword_sentences.append(sentence)

    if keyword_sentences:
        selected = keyword_sentences[:MAX_SENTENCES]
    else:
        selected = sentences[:MAX_SENTENCES]

    text = "。".join(selected)
    return text[:MAX_CHARS]


def read_base_data():
    merged_excel = find_excel("合并行业城市")
    df = pd.read_excel(merged_excel, dtype={"股票代码": str})
    df["股票代码"] = df["股票代码"].astype(str).str.zfill(6)
    df["年份"] = df["年份"].astype(int)

    if MAX_REPORTS is not None:
        df = df.head(MAX_REPORTS).copy()

    rows = []
    for _, row in df.iterrows():
        txt_path = find_txt_path(row["股票代码"], row["公司简称"], row["年份"])
        if txt_path is None:
            continue
        content = txt_path.read_text(encoding="utf-8", errors="ignore")
        row_dict = row.to_dict()
        row_dict["分析文本"] = make_analysis_text(content)
        row_dict["文本路径"] = str(txt_path)
        row_dict["任务ID"] = f'{row_dict["股票代码"]}_{row_dict["年份"]}'
        rows.append(row_dict)

        if len(rows) % PRINT_EVERY == 0:
            print(f"已经准备 {len(rows)} 份年报文本")

    return pd.DataFrame(rows)


def make_prompt(row):
    return f"""
你是一名研究上市公司数字化转型的文本分析助手。请根据下面这份年报节选，判断公司是否进行了实质性的数字化转型或金融科技使用。

请只输出 JSON，不要输出解释性段落。JSON 字段必须包括：
{{
  "判断结果": "实质性转型/表层叙事/不明显",
  "数字化转型得分": 0到100的整数,
  "金融科技使用得分": 0到100的整数,
  "主要依据": "不超过120字",
  "涉及技术关键词": ["关键词1", "关键词2"],
  "是否偏金融科技": "是/否"
}}

判断标准：
1. 实质性转型：文本中出现明确的系统、平台、业务流程、生产经营、客户服务、风控、研发、供应链等数字化应用。
2. 表层叙事：主要是口号、规划、泛泛描述，缺少具体应用。
3. 不明显：几乎没有数字化转型或金融科技相关内容。
4. 金融科技使用得分只在金融服务、支付、风控、智能投顾、互联网金融、数字金融等场景明显时给高分。

公司：{row.get("公司简称")}
股票代码：{row.get("股票代码")}
年份：{row.get("年份")}
行业：{row.get("行业名称")}
词典法标准化词频：{row.get("标准化词频_每万词")}

年报节选：
{row.get("分析文本")}
"""


def parse_json_from_text(text):
    text = str(text).strip()
    text = re.sub(r"^```json", "", text)
    text = re.sub(r"^```", "", text)
    text = re.sub(r"```$", "", text).strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    match = re.search(r"\{.*\}", text, re.S)
    if not match:
        raise ValueError("没有解析到 JSON")
    return json.loads(match.group(0))


def call_deepseek(row, config):
    url = config["base_url"] + "/chat/completions"
    headers = {
        "Authorization": "Bearer " + config["api_key"],
        "Content-Type": "application/json",
    }
    payload = {
        "model": config["model"],
        "messages": [
            {
                "role": "system",
                "content": "你只输出严格 JSON。不要输出 Markdown，不要输出多余说明。",
            },
            {"role": "user", "content": make_prompt(row)},
        ],
        "temperature": 0,
        "max_tokens": 1200,
        "stream": False,
    }

    last_error = ""
    for retry in range(3):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=120)
            response.raise_for_status()
            data = response.json()
            message = data["choices"][0]["message"]
            content = message.get("content", "")
            reasoning_content = message.get("reasoning_content", "")
            parsed = parse_json_from_text(content)

            result = {
                "任务ID": row["任务ID"],
                "股票代码": row["股票代码"],
                "公司简称": row["公司简称"],
                "年份": row["年份"],
                "行业名称": row.get("行业名称", ""),
                "所属省份": row.get("所属省份", ""),
                "所属城市": row.get("所属城市", ""),
                "标准化词频_每万词": row.get("标准化词频_每万词", 0),
                "DeepSeek判断结果": parsed.get("判断结果", ""),
                "DeepSeek数字化转型得分": parsed.get("数字化转型得分", 0),
                "DeepSeek金融科技使用得分": parsed.get("金融科技使用得分", 0),
                "DeepSeek主要依据": parsed.get("主要依据", ""),
                "DeepSeek涉及技术关键词": "、".join(parsed.get("涉及技术关键词", [])),
                "DeepSeek是否偏金融科技": parsed.get("是否偏金融科技", ""),
                "DeepSeek推理摘要": str(reasoning_content)[:500],
                "API状态": "成功",
            }
            return result
        except Exception as e:
            last_error = str(e)
            time.sleep(3 + retry * 5)

    return {
        "任务ID": row["任务ID"],
        "股票代码": row["股票代码"],
        "公司简称": row["公司简称"],
        "年份": row["年份"],
        "行业名称": row.get("行业名称", ""),
        "所属省份": row.get("所属省份", ""),
        "所属城市": row.get("所属城市", ""),
        "标准化词频_每万词": row.get("标准化词频_每万词", 0),
        "DeepSeek判断结果": "调用失败",
        "DeepSeek数字化转型得分": None,
        "DeepSeek金融科技使用得分": None,
        "DeepSeek主要依据": last_error[:300],
        "DeepSeek涉及技术关键词": "",
        "DeepSeek是否偏金融科技": "",
        "DeepSeek推理摘要": "",
        "API状态": "失败",
    }


def load_checkpoint():
    if CHECKPOINT_CSV.exists():
        return pd.read_csv(CHECKPOINT_CSV, dtype={"股票代码": str})
    return pd.DataFrame()


def save_checkpoint(df):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(CHECKPOINT_CSV, index=False, encoding="utf-8-sig")


def run_deepseek(base_df, config):
    old_df = load_checkpoint()
    done_ids = set(old_df["任务ID"].astype(str)) if len(old_df) > 0 else set()
    todo_df = base_df[~base_df["任务ID"].astype(str).isin(done_ids)].copy()

    print("已有结果：", len(done_ids))
    print("本次需要调用 DeepSeek：", len(todo_df))

    results = old_df.to_dict("records") if len(old_df) > 0 else []
    new_results = []

    if len(todo_df) == 0:
        return pd.DataFrame(results)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_map = {
            executor.submit(call_deepseek, row.to_dict(), config): row["任务ID"]
            for _, row in todo_df.iterrows()
        }

        for future in as_completed(future_map):
            result = future.result()
            results.append(result)
            new_results.append(result)
            finished = len(results)

            if finished % SAVE_EVERY == 0:
                save_checkpoint(pd.DataFrame(results))
                print(f"已保存 checkpoint：{finished}/{len(base_df)}")

            if finished % PRINT_EVERY == 0:
                print(f"已经完成 DeepSeek 分析：{finished}/{len(base_df)}")

    result_df = pd.DataFrame(results)
    save_checkpoint(result_df)
    return result_df


def merge_finbert(result_df):
    try:
        finbert_excel = find_excel("FinBERT分析结果")
        finbert_df = pd.read_excel(finbert_excel, sheet_name=0, dtype={"股票代码": str})
        finbert_df["股票代码"] = finbert_df["股票代码"].astype(str).str.zfill(6)
        keep_cols = [
            "股票代码",
            "年份",
            "FinBERT综合语义得分",
            "词典标准化指数",
            "FinBERT综合语义得分标准化指数",
        ]
        finbert_df = finbert_df[keep_cols].copy()
        result_df = pd.merge(result_df, finbert_df, on=["股票代码", "年份"], how="left")
    except Exception as e:
        print("FinBERT 结果合并失败，先跳过：", e)
        result_df["词典标准化指数"] = min_max_100(result_df["标准化词频_每万词"])
        result_df["FinBERT综合语义得分标准化指数"] = None

    result_df["DeepSeek数字化转型指数"] = min_max_100(result_df["DeepSeek数字化转型得分"])
    result_df["词典法判断"] = result_df["词典标准化指数"].apply(make_level)
    result_df["FinBERT判断"] = result_df["FinBERT综合语义得分标准化指数"].apply(make_level)
    result_df["DeepSeek与词典法是否一致"] = result_df["DeepSeek判断结果"] == result_df["词典法判断"]
    result_df["DeepSeek与FinBERT是否一致"] = result_df["DeepSeek判断结果"] == result_df["FinBERT判断"]
    return result_df


def build_summary(df, group_col):
    value_cols = [
        "DeepSeek数字化转型得分",
        "DeepSeek金融科技使用得分",
        "词典标准化指数",
        "FinBERT综合语义得分标准化指数",
    ]
    summary = df.groupby(group_col)[value_cols].mean().reset_index()
    return summary.round(4)


def save_year_line(year_df):
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    plt.figure(figsize=(10, 5))
    plt.plot(year_df["年份"], year_df["DeepSeek数字化转型得分"], marker="o", label="DeepSeek数字化得分", color=BLUE)
    plt.plot(year_df["年份"], year_df["词典标准化指数"], marker="s", label="词典法标准化指数", color=ORANGE)
    plt.plot(year_df["年份"], year_df["FinBERT综合语义得分标准化指数"], marker="^", label="FinBERT语义指数", color=GREEN)
    plt.title("年份维度：DeepSeek、词典法与FinBERT对比")
    plt.xlabel("年份")
    plt.ylabel("平均指数")
    plt.grid(alpha=0.25)
    plt.legend()
    plt.tight_layout()
    output = CHART_DIR / "年份维度_DeepSeek词典FinBERT对比.png"
    plt.savefig(output, dpi=300)
    plt.close()
    print("已保存年份对比图：", output)


def save_industry_bar(industry_df):
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    plot_df = industry_df.sort_values("DeepSeek数字化转型得分", ascending=False).head(12)
    plt.figure(figsize=(10, 6))
    plt.barh(plot_df["行业名称"], plot_df["DeepSeek数字化转型得分"], color=BLUE, label="DeepSeek数字化得分")
    plt.barh(plot_df["行业名称"], plot_df["DeepSeek金融科技使用得分"], color=GREEN, alpha=0.65, label="金融科技得分")
    plt.gca().invert_yaxis()
    plt.title("行业维度：DeepSeek数字化与金融科技得分")
    plt.xlabel("平均得分")
    plt.legend()
    plt.tight_layout()
    output = CHART_DIR / "行业维度_DeepSeek得分条形图.png"
    plt.savefig(output, dpi=300)
    plt.close()
    print("已保存行业条形图：", output)


def save_stage_pie(df):
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    counts = df["DeepSeek判断结果"].value_counts()
    colors = [BLUE, ORANGE, GRAY, RED]
    plt.figure(figsize=(6, 6))
    plt.pie(counts.values, labels=counts.index, autopct="%1.1f%%", startangle=90, colors=colors[:len(counts)])
    plt.title("DeepSeek判断结果分布")
    plt.tight_layout()
    output = CHART_DIR / "总体维度_DeepSeek判断结果饼图.png"
    plt.savefig(output, dpi=300)
    plt.close()
    print("已保存判断结果饼图：", output)


def save_city_map(city_df):
    CITY_MAP_DIR.mkdir(parents=True, exist_ok=True)
    map_df = city_df.dropna(subset=["所属城市"]).copy()
    map_df["地图城市"] = map_df["所属城市"].apply(clean_city_name)

    geo = Geo(init_opts=opts.InitOpts(width="980px", height="760px", bg_color="white"))
    geo.add_schema(maptype="china", itemstyle_opts=opts.ItemStyleOpts(color="#F7FBFF", border_color="#9AA6B2"))

    valid_values = []
    skipped = []
    for _, row in map_df.iterrows():
        city = row["地图城市"]
        value = float(row["DeepSeek数字化转型得分"])
        if city == "" or geo.get_coordinate(city) is None:
            skipped.append(city)
            continue
        valid_values.append((city, value))

    if len(valid_values) == 0:
        print("没有可识别坐标的城市，跳过城市地图")
        return

    max_value = max([item[1] for item in valid_values])

    geo.add(
        "DeepSeek数字化得分",
        valid_values,
        type_=ChartType.HEATMAP,
        symbol_size=13,
    )
    geo.add(
        "城市名称",
        valid_values,
        type_=ChartType.SCATTER,
        symbol_size=5,
        label_opts=opts.LabelOpts(is_show=True, formatter="{b}", font_size=9, color="#333333"),
    )
    geo.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    geo.set_global_opts(
        title_opts=opts.TitleOpts(title="城市维度：DeepSeek数字化转型得分"),
        legend_opts=opts.LegendOpts(is_show=False),
        visualmap_opts=opts.VisualMapOpts(
            min_=0,
            max_=max_value,
            range_color=MAP_COLORS,
            is_piecewise=False,
            pos_left="left",
            pos_bottom="8%",
        ),
    )

    output_html = CITY_MAP_DIR / "城市维度_DeepSeek数字化得分地图.html"
    geo.render(str(output_html))
    save_html_as_png(str(output_html), str(output_html.with_suffix(".png")))
    print("已保存城市地图：", output_html)
    if skipped:
        print("城市地图跳过无坐标地点：", sorted(set(skipped)))


def save_excel(result_df, year_df, industry_df, city_df):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    note_df = pd.DataFrame([
        ["模型", "DeepSeek R1，API模型名 deepseek-reasoner"],
        ["方法", "每份年报抽取数字化关键词相关句子，发送给 DeepSeek R1 做语义判断。"],
        ["判断结果", "实质性转型、表层叙事、不明显"],
        ["对比方法", "将 DeepSeek 得分与词典法标准化指数、FinBERT语义指数进行年份、行业、城市维度对比。"],
        ["注意", "DeepSeek结果属于语义判断，适合课堂展示和辅助分析，不等同于人工审稿结论。"],
    ], columns=["项目", "说明"])

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="年报DeepSeek明细", index=False)
        year_df.to_excel(writer, sheet_name="年份维度", index=False)
        industry_df.to_excel(writer, sheet_name="行业维度", index=False)
        city_df.to_excel(writer, sheet_name="城市维度", index=False)
        note_df.to_excel(writer, sheet_name="PPT说明", index=False)


def main():
    config = get_deepseek_config()
    print("当前模型：", config["model"])
    print("开始读取年报文本")
    base_df = read_base_data()
    print("可分析年报数量：", len(base_df))

    result_df = run_deepseek(base_df, config)
    result_df = merge_finbert(result_df)
    result_df = result_df.sort_values(["年份", "股票代码"]).reset_index(drop=True)

    year_df = build_summary(result_df, "年份")
    industry_df = build_summary(result_df, "行业名称")
    city_df = build_summary(result_df, "所属城市")

    save_excel(result_df, year_df, industry_df, city_df)
    save_year_line(year_df)
    save_industry_bar(industry_df)
    save_stage_pie(result_df)
    save_city_map(city_df)

    print("\nDeepSeek R1 分析完成")
    print("成功数量：", (result_df["API状态"] == "成功").sum())
    print("失败数量：", (result_df["API状态"] == "失败").sum())
    print("结果 Excel：", OUTPUT_EXCEL)
    print("结果文件夹：", OUTPUT_DIR)


if __name__ == "__main__":
    main()
