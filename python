from docx import Document
from docx.table import _Cell
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import re
from collections import defaultdict

class MockupExtractor:
    def __init__(self, word_file_path):
        self.doc = Document(word_file_path)
        self.titles = []  # 存储 Title 数据
        self.footnotes = []  # 存储 Footnote 数据
        self.current_title3 = None  # 当前表的 Title3
        self.current_output_name = None  # 当前表的 OutputName
        
    def extract_number_from_title(self, text):
        """从标题中提取编号，如 '14.1.1.1 受试者筛选情况' -> '14.1.1.1'"""
        match = re.match(r'^([\d.]+)\s+', text.strip())
        if match:
            return match.group(1).strip()
        return None
    
    def get_output_name(self, number_str):
        """将编号转换为 OutputName，如 '14.1.1.1' -> 't_14_1_1_1'"""
        if number_str:
            return 't_' + number_str.replace('.', '_')
        return None
    
    def determine_title_level(self, number_str):
        """判断标题级别，通过编号的点数
        14.1 -> Title1 (2 level)
        14.1.1 -> Title2 (3 level)
        14.1.1.1 -> Title3 (4 level)
        """
        if not number_str:
            return 0
        level = number_str.count('.') + 1
        return level
    
    def replace_percent(self, text):
        """将 % 替换为 \\u37\\"""
        if text is None:
            return text
        return text.replace('%', '\\u37\\')
    
    def extract_paragraphs_text(self):
        """从 Word 中提取所有段落的文本"""
        all_paragraphs = []
        for element in self.doc.element.body:
            if element.tag.endswith('p'):  # 段落
                para = element.getparent().find(element)
                from docx.oxml import parse_xml
                # 使用 docx 库的段落解析
                for para_obj in self.doc.paragraphs:
                    if para_obj._element == element:
                        all_paragraphs.append(para_obj.text)
                        break
        return all_paragraphs
    
    def extract_all_text_from_doc(self):
        """提取 Word 文档中的所有文本"""
        all_text = []
        
        # 提取段落
        for para in self.doc.paragraphs:
            text = para.text.strip()
            if text:
                all_text.append(text)
        
        # 提取表格中的文本
        for table in self.doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_text.append(cell_text)
                if row_text:
                    all_text.append('\t'.join(row_text))
        
        return all_text
    
    def parse_titles_and_footnotes(self):
        """解析所有标题和脚注"""
        all_text = self.extract_all_text_from_doc()
        
        title_hierarchy = defaultdict(str)  # 存储各级标题
        current_title1 = None
        current_title2 = None
        current_title3 = None
        current_footnotes = []
        
        i = 0
        while i < len(all_text):
            line = all_text[i].strip()
            
            # 检查是否是标题行
            number = self.extract_number_from_title(line)
            
            if number:
                level = self.determine_title_level(number)
                
                if level == 2:  # Title1
                    # 保存前一个表的信息
                    if current_title3:
                        self._save_title_and_footnotes(
                            current_title1, current_title2, current_title3, current_footnotes
                        )
                    current_title1 = line
                    current_title2 = None
                    current_title3 = None
                    current_footnotes = []
                    
                elif level == 3:  # Title2
                    # 保存前一个表的信息
                    if current_title3:
                        self._save_title_and_footnotes(
                            current_title1, current_title2, current_title3, current_footnotes
                        )
                    current_title2 = line
                    current_title3 = None
                    current_footnotes = []
                    
                elif level == 4:  # Title3
                    # 保存前一个表的信息
                    if current_title3:
                        self._save_title_and_footnotes(
                            current_title1, current_title2, current_title3, current_footnotes
                        )
                    current_title3 = line
                    current_footnotes = []
                    self.current_output_name = self.get_output_name(number)
            
            # 检查是否是脚注行（以 (1)、(2) 等开头）
            elif re.match(r'^\(\d+\)', line):
                current_footnotes.append(line)
            
            i += 1
        
        # 保存最后一个表
        if current_title3:
            self._save_title_and_footnotes(
                current_title1, current_title2, current_title3, current_footnotes
            )
    
    def _save_title_and_footnotes(self, title1, title2, title3, footnotes):
        """保存一个表的 Title 和 Footnote"""
        if not title3:
            return
        
        number = self.extract_number_from_title(title3)
        output_name = self.get_output_name(number)
        
        # 保存 Title
        title_row = {
            'OutputName': output_name,
            'Title1': self.replace_percent(title1) if title1 else '',
            'Title2': self.replace_percent(title2) if title2 else '',
            'Title3': self.replace_percent(title3) if title3 else ''
        }
        self.titles.append(title_row)
        
        # 保存 Footnotes
        for idx, footnote_text in enumerate(footnotes):
            style = '\\cufi-0' if idx == 0 else '\\par\\cufi-0'
            footnote_row = {
                'OutputName': output_name,
                'Style': style,
                'FootNote': self.replace_percent(footnote_text)
            }
            self.footnotes.append(footnote_row)
    
    def to_excel(self, output_file_path):
        """导出到 Excel 文件"""
        # 解析数据
        self.parse_titles_and_footnotes()
        
        # 创建 Excel 工作簿
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # 删除默认 sheet
        
        # Sheet1: Title
        ws_title = wb.create_sheet('Title', 0)
        ws_title.append(['OutputName', 'Title1', 'Title2', 'Title3'])
        for title_row in self.titles:
            ws_title.append([
                title_row['OutputName'],
                title_row['Title1'],
                title_row['Title2'],
                title_row['Title3']
            ])
        
        # Sheet2: Footnote
        ws_footnote = wb.create_sheet('Footnote', 1)
        ws_footnote.append(['OutputName', 'Style', 'FootNote'])
        for footnote_row in self.footnotes:
            ws_footnote.append([
                footnote_row['OutputName'],
                footnote_row['Style'],
                footnote_row['FootNote']
            ])
        
        # 保存文件
        wb.save(output_file_path)
        print(f"Excel 文件已生成: {output_file_path}")
        print(f"Title 行数: {len(self.titles)}")
        print(f"Footnote 行数: {len(self.footnotes)}")


# 使用示例
if __name__ == '__main__':
    # 替换为你的 Word 文件路径
    word_file = 'mockup.docx'
    output_file = 'output.xlsx'
    
    extractor = MockupExtractor(word_file)
    extractor.to_excel(output_file)
