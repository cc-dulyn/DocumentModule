import docx2json

# 测试导入和使用DocumentStructureExtractor类
try:
    # 尝试创建一个简单的测试对象
    print("尝试导入DocumentStructureExtractor...")
    extractor_class = docx2json.DocumentStructureExtractor
    print("成功导入DocumentStructureExtractor!")
    print(f"DocumentStructureExtractor类型: {type(extractor_class)}")
    
    # 显示模块的所有公共属性，验证导出是否正确
    print("\ndocx2json模块的公共属性:")
    for attr in dir(docx2json):
        if not attr.startswith('_'):
            print(f"- {attr}")
            
    print("\n导入测试成功完成！")
except AttributeError as e:
    print(f"导入错误: {e}")