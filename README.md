
使用方法：

安装依赖：
pip install python-docx openpyxl
修改文件路径：
word_file = 'mockup.docx'  # 你的 Word 文件
output_file = 'output.xlsx'  # 输出的 Excel 文件

word_file = 'mockup.docx'  # 你的 Word 文件
output_file = 'output.xlsx'  # 输出的 Excel 文件
运行脚本：
python extract_mockup.py
代码逻辑：

✅ 读取 Word 所有段落和表格
✅ 识别标题级别（14.1 → Title1；14.1.1 → Title2；14.1.1.1 → Title3）
✅ 提取 OutputName（如 t_14_1_1_1）
✅ 识别脚注（以 (1)、(2) 开头）
✅ 处理 % → \u37\
✅ 生成两个 Sheet 的 Excel
