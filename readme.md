# txwd_record

## 项目简介

`txwd_record` 是一个用于解析腾讯文档（docs.qq.com）在线表格的 Python 工具。它能够自动抓取指定腾讯文档表格页面的数据，并将其结构化为
pandas DataFrame，便于后续的数据分析和处理。

## 主要功能

- 自动抓取腾讯文档表格页面内容
- 解析页面中的表格数据，智能识别表头和数据行
- 支持自定义 Y 轴容差，适应不同表格布局
- 输出为 pandas DataFrame，方便进一步处理

## 安装依赖

请确保已安装 Python 3.10 及以上版本。

安装依赖包：

```bash
pip install -r requirements.txt
```

## 使用方法

示例代码：

```python
from main import TencentSheetParser

url = "https://docs.qq.com/sheet/xxxxxx"  # 替换为你的腾讯文档表格链接
parser = TencentSheetParser(url)
df = parser.get_dataframe()
print(df)
```

## 依赖列表

- pandas~=2.3.1
- requests~=2.32.4
- demjson3~=3.0.6
- openpyxl~=3.1.5

## 注意事项

- 仅支持腾讯文档在线表格（docs.qq.com/sheet/）
- 若页面结构发生变化，解析可能失效
- 需要科学上网或确保网络可访问腾讯文档

## 贡献与反馈

如有建议或 bug，欢迎提交 issue 或 PR。
