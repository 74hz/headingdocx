import re
from typing import Optional

from docx.shared import Pt
from docx.text.paragraph import Paragraph


def get_outline_level(paragraph: Paragraph) -> Optional[int]:
    """
    获取段落的 outlineLvl 属性（如果有），返回 int，否则 None
    """
    try:
        pPr = paragraph._element.pPr
        if pPr is not None:
            outlineLvl = pPr.find(qn("w:outlineLvl"))
            if outlineLvl is not None:
                val = outlineLvl.get(qn("w:val"))
                if val is not None:
                    return int(val)
    except Exception:
        pass
    return None


def is_bold_and_large(paragraph: Paragraph, min_size: int) -> bool:
    """
    判断段落是否有加粗且字号大于 min_size（磅值）
    """
    for run in paragraph.runs:
        if run.bold and run.font.size:
            try:
                if run.font.size.pt >= min_size:
                    return True
            except Exception:
                continue
    return False


def is_bold(paragraph: Paragraph) -> bool:
    """
    判断段落是否有加粗文本
    """
    return any(run.bold for run in paragraph.runs)


def is_large(paragraph: Paragraph, min_size: int) -> bool:
    """
    判断段落是否有字号大于 min_size（磅值）的文本
    """
    for run in paragraph.runs:
        if run.font.size:
            try:
                if run.font.size.pt >= min_size:
                    return True
            except Exception:
                continue
    return False


def is_bold_and_numbered(paragraph: Paragraph) -> bool:
    """
    判断段落是否加粗且有编号（numPr）
    """
    if not is_bold(paragraph):
        return False
    try:
        pPr = paragraph._element.pPr
        if pPr is not None and pPr.find(qn("w:numPr")) is not None:
            return True
    except Exception:
        pass
    return False


def qn(tag: str) -> str:
    """
    快捷获取带命名空间的标签名
    """
    from docx.oxml.ns import qn as _qn

    return _qn(tag)


def match_heading_patterns(text: str) -> bool:
    """
    判断文本是否匹配常见章节标题模式
    """
    patterns = [
        r"^第[一二三四五六七八九十百千]+[章节部分条款篇]",
        r"^[一二三四五六七八九十]+[、\s]",
        r"^[(（][一二三四五六七八九十]+[)）]",
        r"^\d+\.\d+(\.\d+)*[\s]",
        r"^\d+[、\s]",
    ]
    return any(re.match(p, text) for p in patterns)
