# headingdocx-api

基于 headingdocx 工具包的 REST API 服务，支持 docx 标题提取、重组、段落 XML 获取、正则替换等功能。

## 主要功能

- **/headings**：上传 docx 文件，返回所有标题及级别
- **/rebuild**：上传 docx 文件和标题顺序，返回重组后的文档
- **/paragraph_xml**：上传 docx 文件，返回所有段落的 XML
- **/regex_replace**：对 XML 字符串进行正则替换

## 快速开始

1. 安装依赖（假设 headingdocx 已在本地或 PyPI 可用）：

   ```sh
   pip install -e .
   ```

2. 启动服务：

   ```sh
   uvicorn app.main:app --reload
   ```

3. 访问接口文档：

   打开浏览器访问 [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)

## 依赖

- fastapi
- uvicorn
- python-docx
- headingdocx（本地或 PyPI）

## 目录结构

```
headingdocx-api/
├── app/
│   ├── __init__.py
│   ├── main.py
│   └── api.py
├── README.md
├── .gitignore
├── pyproject.toml
```

## License

MIT
