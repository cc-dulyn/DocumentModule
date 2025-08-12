import os
import tempfile
from docx2json import extract_document


def process_docx_fileobj(file_obj):
    """
    直接处理file对象并返回提取的JSON数据

    参数:
        file_obj: 文件对象，需要以二进制模式打开（rb）

    返回:
        str: 提取的JSON字符串
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    temp_file_path = None
    json_output_path = None

    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
            # 从file对象读取内容并写入临时文件
            content = file_obj.read()
            temp_file.write(content)
            temp_file_path = temp_file.name
            print(f"已创建临时文件: {temp_file_path}")

        # 调用提取函数处理临时文件
        print("开始提取文档内容...")
        extract_document(temp_file_path, True)

        # 构建输出JSON文件路径
        file_name = os.path.basename(temp_file_path).replace('.docx', '')
        json_output_path = os.path.join(
            current_dir, 'output', f'{file_name}_get.json'
        )
        print(f"JSON输出路径: {json_output_path}")

        # 读取并返回JSON内容
        if os.path.exists(json_output_path):
            with open(json_output_path, 'r', encoding='utf-8') as f:
                json_data = f.read()
                print("文档内容提取成功")
                return json_data
        else:
            print("提取失败，未找到输出的JSON文件")
            return None

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        return None

    finally:
        # 清理临时文件
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)
            print(f"已删除临时文件: {temp_file_path}")

        # 清理输出的JSON文件
        if json_output_path and os.path.exists(json_output_path):
            os.remove(json_output_path)
            print(f"已删除JSON文件: {json_output_path}")


# 演示如何使用
if __name__ == "__main__":
    # 示例：打开本地文件作为file对象并处理
    test_docx_path = "docx\合同管理.docx"

    if os.path.exists(test_docx_path):
        # 以二进制模式打开文件，获取file对象
        with open(test_docx_path, 'rb') as file_obj:
            # 直接传入file对象进行处理
            result = process_docx_fileobj(file_obj)

        # 打印结果
        if result:
            print("\n提取的JSON内容预览:")
            print(result[:500] + "..." if len(result) > 500 else result)
    else:
        print(f"测试文件不存在: {test_docx_path}")
