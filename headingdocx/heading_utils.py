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
