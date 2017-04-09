# encoding=utf8
import sys
import math

from docx import Document
from docx.enum.shape import WD_INLINE_SHAPE
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from lxml import etree
reload(sys)
sys.setdefaultencoding('utf8')

def check_section_margins(sections):
    top_margin = 0
    bottom_margin = 0
    left_margin = 0
    right_margin = 0
    for section in sections:
        print "Top: ", round(section.top_margin.cm, 2)
        print "Bottom: ", round(section.bottom_margin.cm, 2)
        print "Left: ", round(section.left_margin.cm, 2)
        print "Right: ", round(section.right_margin.cm, 2)

document = Document('data/internet.docx')

# print "core_properties"
# print document.core_properties.author
# print document.core_properties.category
# print document.core_properties.comments
# print document.core_properties.content_status
# print document.core_properties.created
# print document.core_properties.identifier
# print document.core_properties.keywords
# print document.core_properties.language
# print document.core_properties.last_modified_by
# print document.core_properties.last_printed
# print document.core_properties.modified
# print document.core_properties.revision
# print document.core_properties.subject
# print document.core_properties.title
# print document.core_properties.version
last_style = ''


tables = document.tables

# for sec in document.sections:
#     print sec.header.body

# for row in tables[0].rows:
#     for cell in row.cells:
#         for paragraph in cell.paragraphs:
            # print paragraph.text, paragraph.alignment


# for row in tables[0].rows:
#     print row.cells[0].paragraphs[0].text + row.cells[2].paragraphs[0].text

for p in document.paragraphs:
#     run = p.add_run()
    print p.text, p.style.paragraph_format.space_after
#     if p.style.name != "Normal":
#         if "Заголовок" in p.style.name:
#             print p.style.name, p.text
#             print '\n'
#         elif "Основной" in p.style.name:
#             last_style = p.style.name
#             if last_style == p.style.name:
#                 print p.style.name, p.text
#             else:
#                 # p.style.name, p.text
#                 last_style = ''
#
#         elif "toc" in p.style.name:
#             print "Оглавление: ", p.style.name, p.text
#         elif "Указатель" in p.style.name:
#             print p.style.name, p.text

# body_element = document._body._body
# print(body_element.xml)
#
# ps = body_element.xpath('./w:hyperlink/w:r')
#
# paragraphs = [Paragraph(p, None) for p in ps]
#



# print "!!!", ps

# styles = document.styles
# for style in styles:
#     if "ФИО" in style.name:
#         try:
#             print style.font.name
#         except:
#             print 'Родительский стиль: ', style.base_style.name
#         print '\n'

#style.builtin #Стиль сделан пользователем
#style.paragraph_format.first_line_indent.cm # Отступ первой строки
#style.paragraph_format.space_before.pt # Отступ до
#style.paragraph_format.space_after.pt # Отступ после
#style.paragraph_format.line_spacing # Межстрочный интервал