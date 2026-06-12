# Annualreport_tools 学习版

作者：Winter

这是一个用于课堂展示的上市公司年报文本分析项目。项目以中证 A50 公司为样本，围绕近 10 年年报文本，完成年报链接获取、PDF 下载、TXT 转换、数字化转型关键词词频分析、行业和城市信息合并、词云与地图可视化、LDA 主题分析，以及 FinBERT 语义和情感分析。

本项目保留课堂学习需要的核心代码，不上传大量 PDF、TXT、Excel 和图片输出文件。

## 项目结构

```text
01_get_report_links.py          获取中证A50年报链接
02_download_reports_to_txt.py   下载PDF并转换TXT
03_text_analysis.py             关键词词频统计
04_merge_industry_city_data.py  合并行业、城市、省份信息
05_wordcloud_analysis.py        词云、行业图、城市和省份地图
06_lda_topic_analysis.py        sklearn LDA主题分析
07_finbert_analysis.py          FinBERT语义和情感分析
08_deepseek_r1_analysis.py      DeepSeek R1语义级别分析
digital_keywords.py             数字化转型关键词词典
visual_utils.py                 统一颜色和静态地图截图工具
requirements.txt                依赖库
```

## 运行顺序

如果已经完成年报链接获取和 TXT 文本下载，可以直接从第 3 步开始：

```powershell
python .\03_text_analysis.py
python .\04_merge_industry_city_data.py
python .\05_wordcloud_analysis.py
python .\06_lda_topic_analysis.py
python .\07_finbert_analysis.py
python .\08_deepseek_r1_analysis.py
```

完整流程如下：

```powershell
python .\01_get_report_links.py
python .\02_download_reports_to_txt.py
python .\03_text_analysis.py
python .\04_merge_industry_city_data.py
python .\05_wordcloud_analysis.py
python .\06_lda_topic_analysis.py
python .\07_finbert_analysis.py
python .\08_deepseek_r1_analysis.py
```

## 关键词词典

关键词保存在：

```text
digital_keywords.py
```

当前词典按照企业数字化转型的结构化特征分为五类：

- 人工智能技术
- 大数据技术
- 云计算技术
- 区块链技术
- 数字技术运用

后续如果要调整关键词，只需要修改 `digital_keywords.py`，不做自动词典拓展。

## 分析模块

`03_text_analysis.py` 使用 `jieba` 分词、停用词过滤和 `Counter` 统计关键词出现次数，并计算标准化词频：

```text
关键词总次数 / 总词数 * 10000
```

`04_merge_industry_city_data.py` 合并上市公司行业、城市、省份信息。外部数据中的原有智能化转型词频、同年同行业智能化转型总词频、智能化转型程度三列不参与分析。

`05_wordcloud_analysis.py` 输出：

- 总体、前五年、后五年词云图
- 行业维度前 5 大关键词柱状图
- 城市排名图
- 静态省份地图和城市热力地图

`06_lda_topic_analysis.py` 使用 sklearn 的 LDA，固定 5 个主题，不做 K 值复杂评估、不做 pyLDAvis、不做词典拓展。输出每个主题的前 20 个主题词、每份年报的主主题和主题得分，并从年份、行业、城市维度进行可视化。

`07_finbert_analysis.py` 使用中文 FinBERT 进行语义指数和情感概率分析，并从公司、年份、行业、省份、城市、阶段等维度进行对比。

`08_deepseek_r1_analysis.py` 使用 DeepSeek R1 对每份年报进行语义级别判断，输出“实质性转型、表层叙事、不明显”三类结果，并计算数字化转型得分和金融科技使用得分。同时与词典法标准化指数、FinBERT 语义指数进行年份、行业、城市维度对比。

运行第 8 步前，需要在本地创建 `.env` 文件：

```text
DEEPSEEK_API_KEY=你的DeepSeek API Key
DEEPSEEK_MODEL=deepseek-reasoner
DEEPSEEK_BASE_URL=https://api.deepseek.com
```

`.env` 文件只保存在本地，不上传到 GitHub。

## 输出说明

运行后会在本地生成：

```text
年报文件/
```

这个文件夹中包含 PDF、TXT、Excel、PNG、HTML 等数据和结果文件，体积较大，只保留在本地，不上传到 GitHub。

## 依赖说明

不用单独创建新的虚拟环境，可以在已有爬虫环境中安装或检查依赖：

```powershell
pip install -r requirements.txt
```

主要使用到的库包括：

- pandas
- requests
- pdfplumber
- selenium
- jieba
- matplotlib
- wordcloud
- pyecharts
- scikit-learn
- torch
- transformers
