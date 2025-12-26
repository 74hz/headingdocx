# headingdocx

A lightweight Python toolkit for manipulating Word (`.docx`) documents based on heading structure.

## Features

- **Extract all headings**: Get a list of all headings (title and level) from a Word document.
- **Rebuild documents by heading order**: Rearrange, add, or remove sections based on headings and generate a new document.
- **Paragraph XML access**: Retrieve the raw XML of each paragraph for advanced processing.
- **Regex replace in XML**: Perform regex-based replacements directly on XML strings.

## Installation

First, ensure you have `python-docx` installed.  
If you use a virtual environment (recommended):

```bash
pip install python-docx
```

Then, install this package (in your project root):

```bash
pip install -e .
```

## Usage

### 1. Extract all headings

```python
from headingdocx import get_headings

headings = get_headings("input.docx")
print(headings)  # [('Chapter 1', 1), ('1.1 Background', 2), ...]
```

### 2. Rebuild document by heading order

```python
from headingdocx import rebuild_doc_by_headings, get_headings

# Get current headings
headings = get_headings("input.docx")
# Reorder, add, or remove as needed
new_order = [h[0] for h in headings[::-1]]  # Example: reverse order

rebuild_doc_by_headings("input.docx", new_order, "output.docx")
```

### 3. Get all paragraph XML

```python
from headingdocx import get_paragraph_xml

xml_list = get_paragraph_xml("input.docx")
for xml in xml_list:
    print(xml)
```

### 4. Regex replace in XML

```python
from headingdocx import regex_replace_in_xml

xml = '<w:t>Hello World</w:t>'
new_xml = regex_replace_in_xml(xml, r'World', 'Docx')
print(new_xml)  # <w:t>Hello Docx</w:t>
```

## License

MIT

## Author

Your Name
