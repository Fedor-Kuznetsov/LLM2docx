import re
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import latex2mathml.converter
from mathml2omml import convert


def latex_to_omml(latex_str):
    """Конвертирует строку LaTeX в формат OMML (родные формулы Word)."""
    try:
        # 1. LaTeX -> MathML
        mathml = latex2mathml.converter.convert(latex_str)
        
        # Убеждаемся, что пространство имен MathML присутствует
        if 'xmlns="http://www.w3.org/1998/Math/MathML"' not in mathml:
            mathml = mathml.replace('<math>', '<math xmlns="http://www.w3.org/1998/Math/MathML">')
            
        # 2. MathML -> OMML
        omml = convert(mathml)
        
        # ВАЖНО: Добавляем пространство имен Word для математики, иначе parse_xml упадет
        ns = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
        if 'xmlns:m=' not in omml:
            omml = omml.replace('<m:oMath>', f'<m:oMath {ns}>')
            omml = omml.replace('<m:oMathPara>', f'<m:oMathPara {ns}>')
            
        return parse_xml(omml)
    except Exception as e:
        print(f"[ОШИБКА ФОРМУЛЫ] Не удалось конвертировать: {latex_str}")
        print(f"Причина: {e}")
        return None

def add_rich_text(paragraph, text):
    """Добавляет в параграф текст, разбивая его на обычный текст, жирный текст и формулы."""
    # Ищем вхождения формул $...$ ИЛИ жирного текста **...**
    parts = re.split(r'(\$.*?\$|\*\*.*?\*\*)', text)
    
    for part in parts:
        if not part:
            continue
            
        # Обработка выключных формул
        if part.startswith('$') and part.endswith('$') and len(part) > 2:
            latex_str = part[1:-1].strip()
            math_el = latex_to_omml(latex_str)
            if math_el is not None:
                paragraph._p.append(math_el)
            else:
                paragraph.add_run(part) # Если ошибка, пишем как есть
                
        # Обработка жирного текста
        elif part.startswith('**') and part.endswith('**') and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
            
        # Обычный текст
        else:
            paragraph.add_run(part)

def process_txt_to_docx(input_txt_path, output_docx_path):
    doc = Document()

    # Настройка полей страницы по ГОСТ
    sections = doc.sections
    for section in sections:
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)

    # Установка базового шрифта
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)

    with open(input_txt_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Разбиваем файл на блоки по двойному переносу строки
    blocks = content.split('\n\n')

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        # Заголовок 1
        if block.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(block[2:])
            run.bold = True
            run.font.size = Pt(16)
            p.paragraph_format.space_after = Pt(12)
            
        # Заголовок 2
        elif block.startswith('## '):
            p = doc.add_paragraph()
            run = p.add_run(block[3:])
            run.bold = True
            run.font.size = Pt(14)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)

        # Списки
        elif block.startswith('- '):
            lines = block.split('\n')
            for line in lines:
                if line.strip().startswith('- '):
                    p = doc.add_paragraph(style='List Bullet')
                    add_rich_text(p, line.strip()[2:])

        # Выключная (отдельная) формула
        elif block.startswith('$$') and block.endswith('$$'):
            latex_str = block[2:-2].strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            math_el = latex_to_omml(latex_str)
            if math_el is not None:
                p._p.append(math_el)
            else:
                p.add_run(latex_str)

        # Обычный текст (абзац)
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25) # Красная строка
            clean_block = block.replace('\n', ' ')
            add_rich_text(p, clean_block)

    doc.save(output_docx_path)
    print(f"Успех! Документ сохранен как: {output_docx_path}")

if __name__ == "__main__":
    INPUT_FILE = "report.txt"

    # Формируем имя: отчет_ген_ДД-ММ-ГГГГ_ЧЧ-ММ-СС.docx
    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    OUTPUT_FILE = f"отчет_ген_{timestamp}.docx"

    if not os.path.exists(INPUT_FILE):
        print(f"[ОШИБКА] Файл '{INPUT_FILE}' не найден в папке со скриптом.")
        print("Положите рядом со скриптом файл report.txt и запустите снова.")
        raise SystemExit(1)

    process_txt_to_docx(INPUT_FILE, OUTPUT_FILE)