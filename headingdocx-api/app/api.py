import os
import tempfile

from fastapi import APIRouter, File, Form, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from headingdocx.core import (
    get_headings,
    get_paragraph_xml,
    rebuild_doc_by_headings,
    regex_replace_in_docx,
    regex_replace_in_xml,
)

router = APIRouter()


@router.post("/headings")
async def extract_headings(file: UploadFile = File(...)):
    """
    上传 docx 文件，返回所有标题及级别
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    headings = get_headings(tmp_path)
    os.remove(tmp_path)
    return {"headings": headings}


@router.post("/rebuild")
async def rebuild_doc(
    file: UploadFile = File(...),
    heading_texts: str = Form(...),  # 逗号分隔标题
):
    """
    上传 docx 文件和标题顺序，返回重组后的 docx 文件
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    output_path = tmp_path + "_rebuild.docx"
    heading_list = [h.strip() for h in heading_texts.split(",")]
    rebuild_doc_by_headings(tmp_path, heading_list, output_path)
    return FileResponse(output_path, filename="rebuild.docx")


@router.post("/paragraph_xml")
async def paragraph_xml(file: UploadFile = File(...)):
    """
    上传 docx 文件，返回所有段落的 XML 列表
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    xml_list = get_paragraph_xml(tmp_path)
    os.remove(tmp_path)
    return JSONResponse({"xml_list": xml_list})


@router.post("/regex_replace")
async def regex_replace(
    xml: str = Form(...),
    pattern: str = Form(...),
    repl: str = Form(...),
):
    """
    对 XML 字符串进行正则替换
    """
    result = regex_replace_in_xml(xml, pattern, repl)
    return {"result": result}


@router.post("/regex_replace_docx")
async def regex_replace_docx(
    file: UploadFile = File(...),
    pattern: str = Form(...),
    repl: str = Form(...),
):
    """
    对 docx 文件所有段落 XML 进行正则替换，并返回新 docx 文件
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    output_path = tmp_path + "_replaced.docx"
    regex_replace_in_docx(tmp_path, pattern, repl, output_path)
    return FileResponse(output_path, filename="replaced.docx")
