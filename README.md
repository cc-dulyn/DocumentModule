# docx2json

一个将DOCX文档转换为JSON格式的工具，保留文档结构并提取图像。

## 功能特性
- 提取文档结构（封面、目录、正文内容）
- 将文本、标题、表格和图像转换为JSON格式
- 支持多种提取库（python-docx和unstructured）
- 图像提取和保存

## 安装方法

### 使用pip安装
```bash
pip install download_path/docx2json-0.1.0-py3-none-any.whl
```

## 使用方法

### 作为Python库
```python
from docx2json import extract_document

docs_path = "输出文件.docx"
json_structure = extract_document(docs_path)
```

### func extract_document

```
extract_document(
	input_source: str | file,
	save_to_file: bool = True, 
    return_json: bool = True
) -> dict | str
```

**输入参数**

| 参数名         | 类型   | 默认值 | 说明                                                         |
| -------------- | ------ | ------ | ------------------------------------------------------------ |
| `input_source` | 多类型 | 无     | 文档输入源，支持两种类型： <br />- 文件对象（含`read()`方法的类文件对象） <br />- 字符串（文档本地路径，如`"data/report.docx"`） |
| `save_to_file` | `bool` | `True` | 控制是否保存结果到本地文件： <br />- `True`：在`output`目录生成 JSON 文件（命名规则：原文件名 +`_get.json`） <br />- `False`：不生成文件，仅内存处理 |
| `return_json`  | `bool` | `True` | 控制返回数据格式： <br />- `True`：返回 JSON 格式字符串 <br />- `False`：返回 Python 字典（`dict`） |

**输出结果**

| 输出类型              | 说明                                                         |
| --------------------- | ------------------------------------------------------------ |
| JSON 字符串（`str`）  | 当`return_json=True`时返回，包含文档结构化数据（标题、层级、内容等） |
| Python 字典（`dict`） | 当`return_json=False`时返回，直接存储文档结构化数据          |

## 依赖项
- python-docx：用于解析DOCX文档
- unstructured：用于备用文档解析
- Pillow：用于图像处理

## 许可证
MIT许可证