# -*- coding: utf-8 -*-
from docx.enum.style import WD_STYLE_TYPE
from docx import *

document = Document()
styles = document.styles

#生成所有段落样式
for s in styles:
    if s.type == WD_STYLE_TYPE.PARAGRAPH:
        document.add_paragraph('Paragraph style is : '+ s.name, style = s)

document.save('para_style.docx')