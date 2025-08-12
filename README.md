# DocumentModule 文档处理项目

## 项目简介
DocumentModule是一个文档处理工具，主要用于将docx文档转换为结构化的JSON格式，并提取文档中的图片。该项目支持文档结构分析、元素提取和图片保存等功能。

## 功能特点
- 将docx文档转换为结构化JSON
- 提取文档中的图片并保存到本地文件夹
- 支持文档章节识别和结构化组织
- 处理文档中的段落、表格和图片等元素
- 保持文档元素的原始顺序

## 安装说明

### 前提条件
- Python 3.7 或更高版本

### 安装依赖
1. 克隆或下载项目到本地
2. 进入项目根目录
3. 安装依赖库：
   ```
   pip install -r requirements.txt
   ```

## 使用方法

### 基本使用
1. 将需要处理的docx文件放入`DocModule/docx`目录

2. 运行主程序：
   ```
   python DocModule/docx2json.py
   ```

3. 处理结果将保存在`DocModule/output`目录下

4. 提取的图片将保存在`DocModule/temp_images`目录下

   

### 进阶使用

1. 参照`Demo_img.py`与`Demo_img.py`的内容
2. 导入相关模块，设定需处理文档路径

3. 调用`extract_document`方法以处理文档

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

### 自定义配置

- 可以修改`docx2json.py`中的`docs_path`变量来指定不同的文档路径
- 可以调整代码中的图片过滤参数来控制提取的图片质量和大小

## 项目结构
```
DocumentModule/
├── .idea/
│   └── ...               # IDE配置文件
├── DocModule/
│   ├── __pycache__/
│   │   └── ...           # 编译缓存文件
│   ├── docx/
│   │   ├── 合同管理.docx
│   │   └── 测试.docx     # 示例文档
│   ├── output/
│   │   └── 合同管理_get.json  # 处理结果示例
│   ├── temp_images/
│   │   └── ...           # 提取的图片
│   ├── Demo.py           # 演示脚本
│   ├── Demo_img.py       # 保存图片演示脚本
│   ├── docx2json.py      # 主程序文件
│   └── docx2json_img.py  # 保存图片主程序文件
├── PPTModule/
│   └── ...               # PPT处理模块（预留）
├── main.py               # 项目主入口
├── requirements.txt      # 项目依赖库
└── README.md             # 项目说明文档
```

## 依赖库
- Pillow==11.3.0：用于图片处理
- unstructured==0.18.11：用于文档元素提取

## JSON样式

文档转换后的JSON结构包含文档的完整层次结构和元素信息，以下是JSON格式的详细说明：

### 整体结构
```json
{
  "type": "document",
  "cover": [
    // 封皮部分元素
  ],
  "toc": [
    // 目录部分元素
  ],
  "content": [
    // 正文部分元素
  ]
}
```

### 元素类型
JSON中的元素主要分为四大类，每种元素都有对应的`type`字段：

1. **标题元素 (heading)**
```json
{
  "type": "heading",
  "index": 元素编号,
  "level": 标题级别 (1-6),
  "text": "标题文本",
  "children": [
    // 子元素数组
  ]
}
```

2. **段落元素 (paragraph)**
```json
{
  "type": "paragraph",
  "index": 元素编号,
  "text": "段落文本内容"
}
```

3. **图片元素 (image)**
```json
{
  "type": "image",
  "index": 元素编号,
  "original_name": "原始图片名称",
  "dimensions": "图片尺寸 (格式: wxh)",
  "file_size": 图片大小 (字节)
}
```

4. **表格元素 (table)**
```json
{
  "type": "table",
  "index": 元素编号,
  "row": 行数,
  "col": 列数,
  "data": [
    ["单元格1", "单元格2", ...],
    // 更多行
  ]
}
```

### 示例
以下是一个简化的JSON输出示例：
```json
{
  "type": "document",
  "cover": [
    {
      "type": "heading",
      "index": 1,
      "level": 1,
      "text": "文档标题",
      "children": []
    }
  ],
  "toc": [
    {
      "type": "heading",
      "index": 2,
      "level": 1,
      "text": "目录",
      "children": [
        {
          "type": "heading",
          "index": 3,
          "level": 2,
          "text": "第一章 介绍",
          "children": []
        }
      ]
    }
  ],
  "content": [
    {
      "type": "heading",
      "index": 4,
      "level": 1,
      "text": "第一章 介绍",
      "children": [
        {
          "type": "paragraph",
          "index": 5,
          "text": "这是一个示例文档。"
        },
        {
          "type": "image",
          "index": 6,
          "original_name": "image1.png",
          "dimensions": "800x600",
          "file_size": 12345
        },
        {
          "type": "table",
          "index": 7,
          "row": 2,
          "col": 2,
          "data": [
            ["表头1", "表头2"],
            ["内容1", "内容2"]
          ]
        }
      ]
    }
  ]
}
```

## 注意事项
1. 确保docx文件格式正确，避免使用过于复杂的格式或特殊元素
2. 支持自定义标题样式，如果使用自定义标题样式，样式名要包含“标题”和数字，用于进行标题识别
3. 处理大文件时可能需要较长时间，请耐心等待
4. 图片提取功能会自动过滤过小的图片（像素长宽均小于100px），以提高处理效率
5. 每次运行程序会自动清理并重新创建图片目录

