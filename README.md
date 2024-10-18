# 网站批量访问与数据保存工具

## 项目概述

该项目是一个Python脚本，使用`Selenium`和`Openpyxl`库批量访问网站，抓取每个网站的标题、状态码、截图，并将这些信息保存到Excel文件中。程序可以自动调整Excel中的图片大小和行高，以确保显示美观。

## 功能

- **批量访问网站**：从文件中读取多个网址并批量访问。
- **抓取网站信息**：抓取网站的标题、状态码和当前时间。
- **网页截图**：自动截取网页的可视区域，并保存为Base64格式。
- **Excel保存**：将抓取的信息和截图插入到Excel文件中，并自动调整图片大小和行高以保证美观。

## 依赖环境

在运行此项目之前，请确保你的系统中安装了以下依赖库：

- Python 3.x
- `Selenium`（用于浏览器自动化）
- `aiohttp`（用于异步获取网站状态码）
- `Pillow`（用于图像处理）
- `Openpyxl`（用于Excel操作）
- `ChromeDriver`（与Chrome浏览器匹配的驱动）

### 安装依赖

你可以使用`pip`来安装所有需要的依赖：

```python
bash

复制代码
pip install selenium aiohttp pillow openpyxl
```

另外，请确保已经安装了`Google Chrome`浏览器，并将与其匹配的`ChromeDriver`添加到系统路径中。

## 使用说明

1. **准备网址文件**：首先，创建一个包含你要访问的网址的文本文件，每行一个网址，例如：

```python
arduino

复制代码
https://example.com
https://anotherwebsite.com
```

1. **运行程序**：

使用命令行运行脚本，并指定包含网址的文件路径。例如，假设网址文件为`urls.txt`：

```python
bash

复制代码
python script.py --urls urls.txt
```

1. **输出结果**：程序运行后，会生成一个名为`websites_data.xlsx`的Excel文件，其中包含每个网站的标题、状态码、截图和时间信息。

## 项目结构

```python
bash


复制代码
.
├── script.py          # 主程序文件
├── urls.txt           # 包含网址的文本文件（示例）
├── websites_data.xlsx # 运行结果保存的Excel文件
└── README.md          # 项目说明文件
```

## 示例

在程序运行后，生成的Excel文件将类似于以下内容：

| 序号 | URL                 | 标题title       | 状态码 | 截图   | 时间                |
| ---- | ------------------- | --------------- | ------ | ------ | ------------------- |
| 1    | https://example.com | Example Domain  | 200    | [图片] | 2024-10-18 10:30:00 |
| 2    | https://another.com | Another Website | 200    | [图片] | 2024-10-18 10:32:00 |

## 注意事项

- 确保`ChromeDriver`的版本与已安装的`Google Chrome`浏览器兼容。
- 当批量访问大量网站时，请考虑网络延迟或部分网站加载缓慢的情况。
