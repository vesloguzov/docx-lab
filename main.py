# encoding=utf8
import sys
from docx import Document

reload(sys)
sys.setdefaultencoding('utf8')






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

for p in document.paragraphs:
    if p.style.name != "Normal":
        if "Заголовок" in p.style.name:
            print '\n'
            print p.style.name, p.text
            print '\n'
        elif "Основной" in p.style.name:
            last_style = p.style.name
            if last_style == p.style.name:
                print p.text
            else:
                p.style.name, p.text
                last_style = ''

        elif "toc" in p.style.name:
            print "Оглавление: ", p.text
        elif "Указатель" in p.style.name:
            print p.style.name, p.text
