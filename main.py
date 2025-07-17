"""
Creation date: 2025/7/17
Creation Time: 14:03
DIR PATH: 
Project Name: txwd_record
FILE NAME: main.py
Editor: cuckoo
"""

import logging
import re
import time
from typing import Dict, List, Optional

import pandas as pd
import requests
from demjson3 import decode

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class TencentSheetParser:
    """用于解析腾讯文档（docs.qq.com）在线表格的解析器"""

    _RECORD_REGEX = re.compile(r'const record=(.*?),replayRecord', re.S)
    _Q_COMMAND_PATTERN = re.compile(r'q\[(\d+),([\d.]+),([\d.]+)]')

    def __init__(self, url: str, y_tolerance: int = 5):
        """
        初始化解析器。

        Args:
            url (str): 目标腾讯文档表格的URL。
            y_tolerance (int, optional): 用于判断文本是否在同一行的Y轴容差值。默认为 5。
        """
        if not url.startswith("https://docs.qq.com/sheet/"):
            raise ValueError("提供的URL似乎不是一个有效的腾讯在线表格地址。")

        self.df = None
        self.url = url
        self.y_tolerance = y_tolerance
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        })

    def _fetch_page_content(self) -> Optional[str]:
        """从URL获取页面HTML内容"""
        try:
            response = self.session.get(self.url, timeout=10)
            response.raise_for_status()
            logger.info(f"成功获取页面内容，状态码: {response.status_code}")
            return response.text
        except requests.exceptions.RequestException as e:
            logger.error(f"网络请求失败: {e}")
            return None

    def _extract_record_json_string(self, html_content: str) -> Optional[str]:
        """从HTML内容中提取 'record' JSON字符串"""
        match = self._RECORD_REGEX.search(html_content)
        if match:
            logger.info("成功从HTML中匹配到 'record' 数据。")
            return match.group(1)

        logger.warning("在HTML内容中未找到 'const record' 变量，页面结构可能已更改。")
        return None

    @staticmethod
    def _decode_to_dict(json_string: str) -> Optional[Dict]:
        """将JSON字符串解码为Python字典"""
        try:
            data = decode(json_string)
            logger.info("成功将JSON字符串解码为字典。")
            return data
        except Exception as e:
            logger.error(f"解码JSON数据时发生错误: {e}")
            return None

    def _transform_record_to_dataframe(self, record_json: Dict) -> pd.DataFrame:
        """将解析后的record字典转换为结构化的DataFrame"""
        try:
            texts_pool = record_json['flyweight']['texts']
            actions_string = record_json['actions']
            commands = actions_string.split(';')
        except KeyError as e:
            logger.error(f"处理JSON时出错：缺少关键字段 {e}。")
            return pd.DataFrame()

        current_style_group = None
        extracted_data = []

        for cmd in commands:
            if cmd.startswith('g'):
                current_style_group = cmd

            q_match = self._Q_COMMAND_PATTERN.match(cmd)
            if q_match:
                text_index = int(q_match.group(1))
                try:
                    x_coord = float(q_match.group(2))
                    y_coord = float(q_match.group(3))
                    extracted_data.append({
                        "text": texts_pool[text_index],
                        "x": x_coord,
                        "y": y_coord,
                        "style": current_style_group
                    })
                except IndexError:
                    logger.warning(f"文本池索引越界: index {text_index} 超出范围。")
                    continue

        if not extracted_data:
            logger.warning("未在 'actions' 中找到任何可解析的文本数据。")
            return pd.DataFrame()

        # 过滤并排序数据
        main_data = self._filter_out_header_labels(extracted_data)
        sorted_data = sorted(main_data, key=lambda item: (item['y'], item['x']))

        # 重建表格结构
        table_rows_structured = self._group_data_into_rows(sorted_data)
        text_rows = [[item['text'] for item in row] for row in table_rows_structured]

        if not text_rows or not text_rows[0]:
            logger.warning("过滤后没有剩余的表格数据可用于构建DataFrame。")
            return pd.DataFrame()

        # 构建DataFrame
        headers = text_rows[0]
        data = text_rows[1:]
        max_cols = len(headers)

        # 确保所有行长度一致，以表头为准
        padded_data = [row + [''] * (max_cols - len(row)) if len(row) < max_cols else row[:max_cols] for row in data]

        return pd.DataFrame(padded_data, columns=headers)

    @staticmethod
    def _filter_out_header_labels(extracted_data: List[Dict]) -> List[Dict]:
        """智能过滤数据"""
        if not extracted_data:
            return []

        top_most_element = min(extracted_data, key=lambda item: item['y'])
        style_to_exclude = top_most_element['style']

        logger.info(f"检测到画布最顶部的元素为 '{top_most_element['text']}'。")
        logger.info(f"其使用的样式为 '{style_to_exclude}'，将被视为行/列标样式并予以排除。")

        main_data = [item for item in extracted_data if item['style'] != style_to_exclude]

        if not main_data:
            logger.warning("过滤掉行/列标后，没有剩余的表格数据。")

        return main_data

    def _group_data_into_rows(self, sorted_data: List[Dict]) -> List[List[Dict]]:
        """根据Y坐标和容差值将数据点分组为行"""
        if not sorted_data:
            return []

        table_rows = []
        current_row = [sorted_data[0]]
        last_y = sorted_data[0]['y']

        for item in sorted_data[1:]:
            if abs(item['y'] - last_y) > self.y_tolerance:
                table_rows.append(sorted(current_row, key=lambda i: i['x']))  # 在行内按x排序
                current_row = [item]
            else:
                current_row.append(item)
            last_y = item['y']

        table_rows.append(sorted(current_row, key=lambda i: i['x']))
        return table_rows

    def get_dataframe(self) -> pd.DataFrame:
        """获取DataFrame"""
        start_time = time.time()
        logger.info(f"开始处理URL: {self.url}")

        html_content = self._fetch_page_content()
        if not html_content:
            return pd.DataFrame()

        json_string = self._extract_record_json_string(html_content)
        if not json_string:
            return pd.DataFrame()

        record_data = self._decode_to_dict(json_string)
        if not record_data:
            return pd.DataFrame()

        self.df = self._transform_record_to_dataframe(record_data)

        end_time = time.time()
        logger.info(f"处理完成，总耗时: {end_time - start_time:.2f} 秒。")
        return self.df

    def output_to(self, file_path: str, output_format: str = "csv") -> bool:
        """将DataFrame输出为指定格式的文件"""
        if self.df.empty:
            logger.warning("DataFrame为空，无法输出。")
            return False
        if output_format == "csv":
            self.df.to_csv(file_path, index=False)
            logger.info(f"已输出为CSV文件: {file_path}")
            return True
        elif output_format == "excel":
            self.df.to_excel(file_path, index=False)
            logger.info(f"已输出为Excel文件: {file_path}")
            return True
        else:
            logger.error(f"不支持的输出格式: {output_format}")
            return False


if __name__ == "__main__":
    stime = time.time()

    target_url = "https://docs.qq.com/sheet/XXXXXXXXXXXXXXXXXXXXXXX"

    parser = TencentSheetParser(url=target_url)

    my_dataframe = parser.get_dataframe()
    parser.output_to("output.csv", "csv")
    parser.output_to("output.xlsx", "excel")

    print(f"\n脚本总运行时间：{time.time() - stime:.2f}秒")
