# 作者：Winter
# 功能：项目可视化的通用颜色和地图截图工具
import os
import re
import shutil
import time
from pathlib import Path

from PIL import Image, ImageChops
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions


BLUE = "#2F6F9F"
ORANGE = "#F28E2B"
GREEN = "#2E7D32"
GRAY = "#78909C"
RED = "#C62828"
TOPIC_COLORS = ["#2F6F9F", "#F28E2B", "#59A14F", "#E15759", "#76B7B2"]
MAP_COLORS = ["#E8F1FA", "#9DC3E6", "#67C2A5", "#F6E05E", "#F28E2B", "#C62828"]


def safe_filename(name):
    """去掉 Windows 文件名不能用的符号。"""
    return re.sub(r'[\\/:*?"<>|]', "_", str(name))


def get_browser():
    """打开无界面浏览器，用于把 pyecharts 地图保存成静态 PNG。"""
    try:
        options = EdgeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1400,900")
        return webdriver.Edge(options=options)
    except Exception:
        options = ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1400,900")
        return webdriver.Chrome(options=options)


def crop_white_border(image_path, padding=18):
    """裁剪 PNG 周围多余白边，适合放论文。"""
    image = Image.open(image_path).convert("RGB")
    background = Image.new("RGB", image.size, (255, 255, 255))
    diff = ImageChops.difference(image, background)
    bbox = diff.getbbox()

    if bbox is None:
        image.save(image_path)
        return

    left, top, right, bottom = bbox
    left = max(left - padding, 0)
    top = max(top - padding, 0)
    right = min(right + padding, image.size[0])
    bottom = min(bottom + padding, image.size[1])

    image.crop((left, top, right, bottom)).save(image_path)


def save_html_as_png(html_path, output_png):
    """把 pyecharts HTML 地图截图成静态 PNG。"""
    tmp_dir = Path("data") / "map_snapshot_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_html = tmp_dir / (safe_filename(Path(html_path).stem) + ".html")
    shutil.copyfile(html_path, tmp_html)

    driver = None
    try:
        driver = get_browser()
        html_uri = tmp_html.resolve().as_uri()

        for try_count in range(3):
            driver.get(html_uri)
            chart_element = None

            for _ in range(15):
                elements = driver.find_elements("css selector", "div[_echarts_instance_]")
                canvas_count = driver.execute_script("return document.querySelectorAll('canvas').length")

                if len(elements) > 0 and canvas_count > 0:
                    chart_element = elements[0]
                    break

                time.sleep(1)

            if chart_element is None:
                print("地图图表元素暂未加载，重试：", output_png, try_count + 1)
                continue

            time.sleep(2)
            chart_element.screenshot(output_png)
            crop_white_border(output_png)

            if os.path.exists(output_png) and os.path.getsize(output_png) > 20000:
                print("已保存静态地图：", output_png)
                return

            print("地图截图未加载完整，重试：", output_png, try_count + 1)

        print("静态地图可能未加载完整，请检查：", output_png)

    except Exception as e:
        print("静态地图截图失败：", output_png, e)

    finally:
        if driver is not None:
            driver.quit()
