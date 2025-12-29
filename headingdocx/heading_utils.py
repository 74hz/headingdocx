import re

NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def get_outline_level_xml(p):
    """获取段落 outlineLvl 属性"""
    lvl_elem = p.xpath("./w:pPr/w:outlineLvl", namespaces=NAMESPACE)
    if lvl_elem:
        try:
            return int(lvl_elem[0].get("{%s}val" % NAMESPACE["w"]))
        except Exception:
            return None
    return None


def is_bold_xml(p):
    """判断段落是否加粗"""
    return bool(p.xpath(".//w:rPr/w:b", namespaces=NAMESPACE))


def is_large_xml(p, min_size=32):
    """判断段落字号是否大于 min_size（单位：半磅）"""
    sz_elems = p.xpath(".//w:rPr/w:sz", namespaces=NAMESPACE)
    for sz in sz_elems:
        try:
            val = int(sz.get("{%s}val" % NAMESPACE["w"]))
            if val >= min_size:
                return True
        except Exception:
            continue
    return False


def is_bold_and_large_xml(p, min_size=32):
    return is_bold_xml(p) and is_large_xml(p, min_size)


def is_bold_and_numbered_xml(p):
    """判断是否加粗且带编号（简单实现）"""
    text = "".join(p.xpath(".//w:t/text()", namespaces=NAMESPACE)).strip()
    return is_bold_xml(p) and bool(re.match(r"^\d+[\.\、]", text))


def match_heading_patterns(text):
    """可直接复用原有实现"""
    return bool(re.match(r"^第[一二三四五六七八九十]+章", text))


def is_heading_like(p):
    """
    统一的标题判定逻辑，供 get_headings 和 rebuild_doc_by_headings 共用。
    返回 (is_heading: bool, level: Optional[int])
    """
    text = "".join(p.xpath(".//w:t/text()", namespaces=NAMESPACE)).strip()
    if not text:
        return False, None
    style_elems = p.xpath("./w:pPr/w:pStyle", namespaces=NAMESPACE)
    style_name = style_elems[0].get("{%s}val" % NAMESPACE["w"]) if style_elems else ""
    level = None
    if style_name:
        m1 = re.match(r"^Heading\s*(\d+)$", style_name, re.I)
        m2 = re.match(r"^标题\s*(\d+)$", style_name)
        if m1:
            level = int(m1.group(1))
        elif m2:
            level = int(m2.group(1))
        elif style_name.isdigit():
            num = int(style_name)
            if 1 <= num <= 9:
                level = num
    # outlineLvl 属性
    if level is None:
        outline_lvl = get_outline_level_xml(p)
        if outline_lvl is not None:
            level = outline_lvl + 1
    # 格式特征
    if level is None and text and len(text) < 200:
        if is_bold_and_large_xml(p, min_size=44):
            level = 1
        elif is_bold_and_large_xml(p, min_size=36):
            level = 2
        elif is_bold_and_large_xml(p, min_size=32):
            level = 3
        elif is_bold_and_numbered_xml(p):
            level = 2
        elif match_heading_patterns(text) and is_bold_xml(p):
            if is_large_xml(p, min_size=32):
                level = 1
            else:
                level = 2
    return (level is not None), level
