import docx2json

# 1.通过文件路径，提取Word文档为JSON结构，不保存图片到文件夹
docs_path = "docx\合同管理.docx"
docx2json.extract_document(docs_path,False)

# 2.通过文件路径，提取Word文档为JSON结构，不保存图片到文件夹,返回JSON字符串
with open("docx\合同管理.docx", 'rb') as f:
    json_str = docx2json.extract_document(f)
print(type(json_str))

# 3. 通过文件路径，提取Word文档为JSON结构，不保存图片到文件夹,返回JSON字典
with open("docx\合同管理.docx", 'rb') as f:
    structure = docx2json.extract_document(f, return_json=False)
print(type(structure))
