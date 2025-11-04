import streamlit as st
import re
import pandas as pd
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import os
import tempfile
from io import BytesIO

# 设置页面配置
st.set_page_config(
    page_title="图号匹配工具",
    page_icon="🔍",
    layout="wide"
)

# 添加错误处理
try:
    from docx import Document
    from docx.shared import RGBColor
    from docx.enum.text import WD_COLOR_INDEX

    DOCX_AVAILABLE = True
except ImportError as e:
    st.error(f"导入docx模块失败: {e}")
    DOCX_AVAILABLE = False


class FigureNumberProcessor:
    def __init__(self):
        self.found_figures = []
        self.matched_figures = []
        self.unmatched_figures = []
        self.bom_list = []

    def extract_figure_numbers_from_docx(self, docx_content):
        """从Word文档中提取图号"""
        if not DOCX_AVAILABLE:
            st.error("docx模块不可用，无法处理Word文档")
            return []

        try:
            # 创建临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(docx_content)
                tmp_file_path = tmp_file.name

            doc = Document(tmp_file_path)
            figures = []

            # 定义匹配模式
            patterns = [
                r'\b[A-Z]{1,4}\d{1,3}-[A-Z0-9]{4,10}(?:-[A-Z0-9]{1,3})?\b',
                r'\b[A-Z]?\d{5,6}-[A-Z0-9]{6,7}\b',
                r'\b[A-Z0-9]{4,8}-[A-Z0-9]{3,8}\b',
            ]

            # 提取段落中的图号
            for para_idx, paragraph in enumerate(doc.paragraphs):
                text = paragraph.text
                if text.strip():
                    for pattern in patterns:
                        matches = re.finditer(pattern, text)
                        for match in matches:
                            figure = match.group()
                            if (any(c.isalpha() for c in figure) and
                                    any(c.isdigit() for c in figure) and
                                    '-' in figure):
                                figures.append({
                                    'text': figure,
                                    'type': 'paragraph',
                                    'para_idx': para_idx,
                                    'position': match.start(),
                                    'original_text': text
                                })

            # 提取表格中的图号
            for table_idx, table in enumerate(doc.tables):
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, paragraph in enumerate(cell.paragraphs):
                            text = paragraph.text
                            if text.strip():
                                for pattern in patterns:
                                    matches = re.finditer(pattern, text)
                                    for match in matches:
                                        figure = match.group()
                                        if (any(c.isalpha() for c in figure) and
                                                any(c.isdigit() for c in figure) and
                                                '-' in figure):
                                            figures.append({
                                                'text': figure,
                                                'type': 'table',
                                                'table_idx': table_idx,
                                                'row_idx': row_idx,
                                                'cell_idx': cell_idx,
                                                'para_idx': para_idx,
                                                'position': match.start(),
                                                'original_text': text
                                            })

            # 清理临时文件
            try:
                os.unlink(tmp_file_path)
            except:
                pass

            # 去重
            unique_figures = []
            seen = set()
            for fig in figures:
                if fig['text'] not in seen:
                    seen.add(fig['text'])
                    unique_figures.append(fig)

            self.found_figures = unique_figures
            return unique_figures

        except Exception as e:
            st.error(f"读取Word文档时出错: {e}")
            return []

    def load_bom_from_excel(self, excel_content, bom_column='BOM', sheet_name=0):
        """从Excel文件加载BOM列表"""
        try:
            # 读取Excel文件
            df = pd.read_excel(BytesIO(excel_content), sheet_name=sheet_name, engine='openpyxl')

            # 查找BOM列
            bom_data = []
            if bom_column in df.columns:
                bom_series = df[bom_column].dropna()
                bom_data = bom_series.astype(str).tolist()
            else:
                # 如果没有找到指定列，尝试查找包含'BOM'的列
                found_column = None
                for col in df.columns:
                    if 'bom' in str(col).lower():
                        found_column = col
                        break

                if found_column:
                    bom_series = df[found_column].dropna()
                    bom_data = bom_series.astype(str).tolist()
                else:
                    st.error("未找到BOM列，请检查Excel文件结构")
                    return []

            # 清理数据
            cleaned_bom_data = []
            for item in bom_data:
                item = str(item).strip()
                if item and item != 'nan' and item != 'None':
                    cleaned_bom_data.append(item)

            self.bom_list = list(set(cleaned_bom_data))
            return self.bom_list

        except Exception as e:
            st.error(f"读取Excel文件时出错: {e}")
            return []

    def match_figures_with_bom(self):
        """将提取的图号与BOM列表进行匹配（子串匹配）"""
        if not self.found_figures or not self.bom_list:
            st.error("错误：请先加载Word文档和Excel文件")
            return

        self.matched_figures = []
        self.unmatched_figures = []

        for fig_info in self.found_figures:
            figure = fig_info['text']
            matched = False

            # 子串匹配
            for bom_item in self.bom_list:
                if figure in bom_item:
                    matched = True
                    break

            if matched:
                self.matched_figures.append(fig_info)
            else:
                self.unmatched_figures.append(fig_info)

    def mark_unmatched_in_word(self, original_content):
        """在Word文档中标记不匹配的图号"""
        if not DOCX_AVAILABLE:
            st.error("docx模块不可用，无法生成标记文档")
            return None

        try:
            # 创建临时输入文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_input:
                tmp_input.write(original_content)
                tmp_input_path = tmp_input.name

            doc = Document(tmp_input_path)

            # 标记段落中的不匹配图号
            for fig_info in self.unmatched_figures:
                if fig_info['type'] == 'paragraph':
                    para_idx = fig_info['para_idx']
                    figure = fig_info['text']

                    if para_idx < len(doc.paragraphs):
                        paragraph = doc.paragraphs[para_idx]
                        original_text = paragraph.text

                        paragraph.clear()
                        start_pos = original_text.find(figure)
                        if start_pos != -1:
                            if start_pos > 0:
                                paragraph.add_run(original_text[:start_pos])

                            run = paragraph.add_run(figure)
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.font.highlight_color = WD_COLOR_INDEX.GRAY_50

                            if start_pos + len(figure) < len(original_text):
                                paragraph.add_run(original_text[start_pos + len(figure):])

            # 保存到临时输出文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_output:
                doc.save(tmp_output.name)
                with open(tmp_output.name, 'rb') as f:
                    marked_content = f.read()

            # 清理临时文件
            try:
                os.unlink(tmp_input_path)
                os.unlink(tmp_output.name)
            except:
                pass

            return marked_content

        except Exception as e:
            st.error(f"标记Word文档时出错: {e}")
            return None


def main():
    st.title("🔍 图号匹配工具")
    st.markdown("---")

    # 检查依赖
    if not DOCX_AVAILABLE:
        st.error("⚠️ 系统缺少必要的依赖包，请确保以下包已安装：")
        st.code("pip install docx pandas openpyxl")
        return

    # 初始化session state
    if 'processor' not in st.session_state:
        st.session_state.processor = FigureNumberProcessor()
    if 'processed' not in st.session_state:
        st.session_state.processed = False

    # 侧边栏说明
    with st.sidebar:
        st.header("使用说明")
        st.markdown("""
        1. 上传Word文档（包含图号）
        2. 上传Excel文件（包含BOM数据）  
        3. 配置参数
        4. 开始处理

        **图号格式示例:**
        - Y24-1205100-19
        - S01200-3900750  
        - YCFP-011
        """)

    # 文件上传
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 上传Word文档")
        word_file = st.file_uploader("选择Word文档", type=['docx'])

    with col2:
        st.subheader("📊 上传Excel文件")
        excel_file = st.file_uploader("选择Excel文件", type=['xlsm', 'xlsx'])

    # 配置参数
    st.subheader("⚙️ 配置参数")
    bom_column = st.text_input("BOM列名称", value="BOM")
    sheet_name = st.text_input("工作表名称", value="0")

    # 处理按钮
    if st.button("🚀 开始处理", type="primary"):
        if not word_file or not excel_file:
            st.error("请先上传Word文档和Excel文件！")
            return

        processor = st.session_state.processor

        with st.spinner("正在处理Word文档..."):
            word_content = word_file.getvalue()
            figures = processor.extract_figure_numbers_from_docx(word_content)

        if not figures:
            st.error("未找到符合要求的图号！")
            return

        with st.spinner("正在处理Excel文件..."):
            excel_content = excel_file.getvalue()
            try:
                sheet = int(sheet_name) if sheet_name.isdigit() else sheet_name
            except:
                sheet = 0

            bom_list = processor.load_bom_from_excel(excel_content, bom_column, sheet)

        if not bom_list:
            st.error("未找到BOM数据！")
            return

        with st.spinner("正在匹配图号..."):
            processor.match_figures_with_bom()

        st.session_state.processed = True
        st.success("处理完成！")

    # 显示结果
    if st.session_state.processed:
        processor = st.session_state.processor

        st.markdown("---")
        st.subheader("📊 处理结果")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("提取图号总数", len(processor.found_figures))
        with col2:
            st.metric("匹配图号", len(processor.matched_figures))
        with col3:
            st.metric("不匹配图号", len(processor.unmatched_figures))

        # 显示图号列表
        if processor.found_figures:
            st.subheader("📋 图号列表")
            for i, fig in enumerate(processor.found_figures, 1):
                status = "✅" if fig in processor.matched_figures else "❌"
                st.write(f"{status} {i}. {fig['text']}")

        # 下载标记文档
        if processor.unmatched_figures:
            with st.spinner("生成标记文档..."):
                marked_content = processor.mark_unmatched_in_word(word_content)

                if marked_content:
                    st.download_button(
                        label="📄 下载标记文档",
                        data=marked_content,
                        file_name=f"标记版_{word_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )


if __name__ == "__main__":
    main()