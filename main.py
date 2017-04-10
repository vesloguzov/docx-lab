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

file_analyzes = {}

def get_general_properties(document):
    general_properties = {}
    general_properties["author"] = document.core_properties.author # Автор документа
    general_properties["created"] = document.core_properties.created # Дата создания документа
    general_properties["last_modified_by"] = document.core_properties.last_modified_by # Пользователь, который менял последний
    general_properties["modified"] = document.core_properties.modified # Время изменения
    general_properties["title"] = document.core_properties.title #Название
    file_analyzes["general_properties"] = general_properties

def get_document_margins(document):
    margins = {}
    sections = document.sections
    margins["top"] = round(sections[0].top_margin.cm, 2)
    margins["bottom"] = round(sections[0].bottom_margin.cm, 2)
    margins["left"] = round(sections[0].left_margin.cm, 2)
    margins["right"] = round(sections[0].right_margin.cm, 2)
    file_analyzes["margins"] = margins

def get_custom_styles(document):
    custom_styles = {}
    all_docx_styles = document.styles
    for s in all_docx_styles:
        if s.builtin == False and "ФИО" in s.name:
            if "Заголовок" in s.name:
                header_style = {}
                custom_head_style = all_docx_styles[s.base_style.name]
                header_style["name"] = s.name
                header_style["font_name"] = s.font.name
                header_style["font_size"] = s.font.size
                header_style["font_italic"] = s.font.italic
                header_style["font_bold"] = s.font.bold
                header_style["line_spacing"] = s.paragraph_format.line_spacing
                header_style["first_line_indent"] = s.paragraph_format.first_line_indent
                header_style["space_before"] = s.paragraph_format.space_before
                header_style["space_after"] = s.paragraph_format.space_after
                header_style["alignment"] = s.paragraph_format.alignment
                custom_styles["header_style"] = header_style
            else:
                paragraph_style = {}
                custom_paragraph_style = all_docx_styles[s.base_style.name]
                header_style["name"] = s.name
                header_style["font_name"] = s.font.name

                print s.paragraph_format.alignment

    file_analyzes["custom_styles"] = custom_styles


document = Document('data/internet.docx')

get_general_properties(document)
get_document_margins(document)
get_custom_styles(document)

print file_analyzes


#
# tables = document.tables
#
# for sec in document.sections:
#     print sec.header.body

# for row in tables[0].rows:
#     for cell in row.cells:
#         for paragraph in cell.paragraphs:
            # print paragraph.text, paragraph.alignment


# for row in tables[0].rows:
#     print row.cells[0].paragraphs[0].text + row.cells[2].paragraphs[0].text

# get_section_margins(document.sections)
# print  file_analyzes



# for p in document.paragraphs:
#     run = p.add_run()
#     print p.text, p.style.paragraph_format.space_after
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