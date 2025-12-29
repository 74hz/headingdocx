import re
import zipfile
from io import StringIO
from typing import Iterator, List, Optional, Tuple

from lxml import etree

from .heading_utils import (
    get_outline_level_xml,
    is_bold_and_large_xml,
    is_bold_and_numbered_xml,
    is_bold_xml,
    is_heading_like,
    is_large_xml,
    match_heading_patterns,
)

NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def iter_paragraphs(doc_path: str) -> Iterator[etree._Element]:
    """流式遍历 docx 的所有段落（<w:p>）"""
    with zipfile.ZipFile(doc_path) as z:
        with z.open("word/document.xml") as f:
            context = etree.iterparse(f, events=("end",), tag="{%s}p" % NAMESPACE["w"])
            for event, elem in context:
                yield elem
                elem.clear()


def get_headings(doc_path: str) -> List[Tuple[str, Optional[int]]]:
    """
    流式获取文档中所有标题目录，返回 [(标题文本, 级别), ...]
    """
    headings = []
    for p in iter_paragraphs(doc_path):
        is_heading, level = is_heading_like(p)
        if is_heading:
            text = "".join(p.xpath(".//w:t/text()", namespaces=NAMESPACE)).strip()
            headings.append((text, level))
    return headings


def rebuild_doc_by_headings(doc_path: str, heading_texts: List[str], output_path: str):
    """
    根据给定标题文本列表（heading_texts）重组文档，生成新文档。
    只保留指定标题及其内容，顺序可调整。
    """
    # 1. 收集所有标题及其内容块
    heading_blocks = {}  # {标题: [段落XML, ...]}
    current_heading = None
    current_block = []
    for p in iter_paragraphs(doc_path):
        text = "".join(p.xpath(".//w:t/text()", namespaces=NAMESPACE)).strip()
        is_heading, _ = is_heading_like(p)
        # 新标题开始
        if is_heading:
            if current_heading and current_block:
                heading_blocks[current_heading] = list(current_block)
            current_heading = text
            current_block = [etree.tostring(p, encoding="unicode")]
        else:
            if current_heading:
                current_block.append(etree.tostring(p, encoding="unicode"))
    if current_heading and current_block:
        heading_blocks[current_heading] = list(current_block)

    # 打印调试：heading_blocks 的 key 和传入的 heading_texts
    print("DEBUG heading_blocks keys:", list(heading_blocks.keys()))
    print("DEBUG input heading_texts:", heading_texts)

    # 2. 读取原 document.xml 的头部和尾部
    with zipfile.ZipFile(doc_path) as z:
        with z.open("word/document.xml") as f:
            doc_xml = f.read().decode("utf-8")
    # 获取<w:body>前的头部和</w:body>后的尾部
    body_start = doc_xml.find("<w:body")
    body_end = doc_xml.rfind("</w:body>")
    head = doc_xml[:body_start]
    tail = doc_xml[body_end + len("</w:body>") :]

    # 3. 组装新 document.xml
    out = StringIO()
    # 写入头部和<w:body>标签
    body_tag_start = doc_xml[body_start : doc_xml.find(">", body_start) + 1]
    out.write(head)
    out.write(body_tag_start)
    # 写入重组后的段落
    for title in heading_texts:
        for para_xml in heading_blocks.get(title, []):
            out.write(para_xml)
    # 写入</w:body>和尾部
    out.write("</w:body>")
    out.write(tail)
    new_doc_xml = out.getvalue()

    # 4. 打包新 docx
    with zipfile.ZipFile(doc_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w") as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))


def get_paragraph_xml(doc_path: str) -> Iterator[str]:
    """流式获取所有段落的XML字符串"""
    for p in iter_paragraphs(doc_path):
        yield etree.tostring(p, encoding="unicode")


def regex_replace_in_docx(doc_path: str, pattern: str, repl: str, output_path: str):
    """只对 docx 正文（word/document.xml）做正则替换，并保存为新文件"""
    with zipfile.ZipFile(doc_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    text = data.decode("utf-8")
                    text = re.sub(pattern, repl, text)
                    data = text.encode("utf-8")
                zout.writestr(item, data)
