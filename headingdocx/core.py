import re
from typing import List, Optional, Tuple

from docx import Document


def get_headings(doc_path: str) -> List[Tuple[str, Optional[int]]]:
    """
    获取文档中所有标题目录，返回 [(标题文本, 级别), ...]
    """
    doc = Document(doc_path)
    headings = []
    for paragraph in doc.paragraphs:
        style = getattr(paragraph, "style", None)
        if (
            style is not None
            and getattr(style, "name", None)
            and style.name.startswith("Heading")
        ):
            try:
                level = int(style.name.split()[-1])
                headings.append((paragraph.text, level))
            except Exception:
                headings.append((paragraph.text, None))
    return headings


def rebuild_doc_by_headings(doc_path: str, heading_texts: List[str], output_path: str):
    """
    根据给定标题文本列表（heading_texts）重组文档，生成新文档。
    heading_texts: 标题文本列表，按新顺序排列（可增删）
    output_path: 新文档保存路径
    """
    doc = Document(doc_path)
    new_doc = Document()
    # 提取所有标题及其内容块
    content_blocks = {}
    current_block = []
    current_heading = None
    for paragraph in doc.paragraphs:
        style = getattr(paragraph, "style", None)
        is_heading = (
            style is not None
            and getattr(style, "name", None)
            and style.name.startswith("Heading")
        )
        if is_heading:
            if current_heading:
                content_blocks[current_heading] = current_block
            current_heading = paragraph.text
            current_block = [paragraph]
        else:
            if current_heading:
                current_block.append(paragraph)
    if current_heading:
        content_blocks[current_heading] = current_block
    # 按新顺序组合
    for title in heading_texts:
        if title in content_blocks:
            for para in content_blocks[title]:
                new_doc.element.body.append(para._p)
    new_doc.save(output_path)


def get_paragraph_xml(doc_path: str) -> List[str]:
    """
    获取文档所有段落的XML字符串列表
    """
    doc = Document(doc_path)
    xml_list = []
    for paragraph in doc.paragraphs:
        xml_list.append(paragraph._p.xml)
    return xml_list


def regex_replace_in_xml(xml_str: str, pattern: str, repl: str) -> str:
    """
    对xml字符串进行正则替换
    pattern: 正则表达式
    repl: 替换内容
    返回替换后的xml字符串
    """
    return re.sub(pattern, repl, xml_str)
