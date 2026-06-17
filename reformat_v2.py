# -*- coding: utf-8 -*-
"""
Версия 2: Не урезаем оригинальный контент.
Берём шаблон Карпеева, оставляем ТОЛЬКО титульную страницу (P0-P53),
после неё восстанавливаем оригинальное содержимое студентов из extracted text.
"""
import shutil
import re
from docx import Document
from lxml import etree
from copy import deepcopy

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def wtag(t):
    return '{' + W + '}' + t

TEMPLATE = r'C:\Users\smopo\Downloads\AyuGram Desktop\Отчет_по_летней_практике_Карпеев_Владислав_4210.docx'
EXTRACTED = r'C:\Users\smopo\.claude\projects\C--Users-smopo\11ae5f31-a760-418e-a59d-3fbf8eaa72ac\tool-results\b3cv1jdqb.txt'

STUDENTS = [
    {
        'key': 'akhmadullin',
        'start_marker': r'FILE: C:\Users\smopo\Desktop\Ахмадуллин (1).docx',
        'end_marker': r'FILE: C:\Users\smopo\Desktop\беловодский_практос.docx',
        'out_path': r'C:\Users\smopo\Desktop\Ахмадуллин (1).docx',
    },
    {
        'key': 'belovodsky',
        'start_marker': r'FILE: C:\Users\smopo\Desktop\беловодский_практос.docx',
        'end_marker': r'FILE: C:\Users\smopo\Desktop\Закиров летняя.docx',
        'out_path': r'C:\Users\smopo\Desktop\беловодский_практос.docx',
    },
    {
        'key': 'zakirov',
        'start_marker': r'FILE: C:\Users\smopo\Desktop\Закиров летняя.docx',
        'end_marker': None,  # до конца файла
        'out_path': r'C:\Users\smopo\Desktop\Закиров летняя.docx',
    },
]

# ----------------------------------------------------------------
# Парсим extracted text для одного студента
# Возвращаем список (style_name, text)
# ----------------------------------------------------------------
PARA_RE = re.compile(r'^\[P(\d+)\] \[([^\]]+)\] (.*)$', re.DOTALL)

def parse_student_paragraphs(student):
    with open(EXTRACTED, encoding='utf-8') as f:
        content = f.read()

    start_idx = content.index(student['start_marker'])
    if student['end_marker']:
        end_idx = content.index(student['end_marker'])
    else:
        end_idx = len(content)
    block = content[start_idx:end_idx]
    lines = block.split('\n')

    paragraphs = []
    current_text = None
    current_style = None
    current_idx = None

    for line in lines:
        m = PARA_RE.match(line)
        if m:
            # Сохраняем предыдущий параграф
            if current_text is not None:
                paragraphs.append((current_idx, current_style, current_text))
            current_idx = int(m.group(1))
            current_style = m.group(2)
            current_text = m.group(3)
        else:
            # Продолжение предыдущего параграфа (многострочный)
            if current_text is not None and line.strip():
                current_text += '\n' + line

    # Последний
    if current_text is not None:
        paragraphs.append((current_idx, current_style, current_text))

    return paragraphs

# ----------------------------------------------------------------
# Создаём параграф со стилем
# ----------------------------------------------------------------
def mkpara(text, style='Normal', bold=False):
    p = etree.Element(wtag('p'))
    pPr = etree.SubElement(p, wtag('pPr'))
    if style and style != 'Normal':
        ps = etree.SubElement(pPr, wtag('pStyle'))
        ps.set(wtag('val'), style)
    if text:  # Создаём run только если есть текст
        r = etree.SubElement(p, wtag('r'))
        rPr = etree.SubElement(r, wtag('rPr'))
        if bold:
            etree.SubElement(rPr, wtag('b'))
        t = etree.SubElement(r, wtag('t'))
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
    return p

# ----------------------------------------------------------------
# Обработка одного студента
# ----------------------------------------------------------------
def process(student):
    print(f"\n=== {student['key'].upper()} ===")

    # 1. Копируем шаблон
    out_path = student['out_path']
    shutil.copy(TEMPLATE, out_path)
    doc = Document(out_path)

    # 2. Удаляем всё после титульной страницы (P54 и далее)
    # Титульная страница: P0-P53 включительно (заканчивается "Казань 2026")
    # Находим параграф с "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ" — это P55, удаляем от P54 включительно
    paragraphs_list = doc.paragraphs

    # Найти индекс параграфа с "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ"
    indiv_idx = None
    for i, p in enumerate(paragraphs_list):
        if 'ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ' in (p.text or ''):
            indiv_idx = i
            break

    if indiv_idx is None:
        print('  [!] Не найдено ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ')
        return

    # Удаляем параграф P54 (пустой между Казань 2026 и ИЗ) и все после
    # Удаляем начиная с indiv_idx-1 (если P54 пустой) или с indiv_idx
    # Будем удалять начиная с indiv_idx (включая)
    body = paragraphs_list[indiv_idx]._element.getparent()
    elements_to_remove = []
    cur = paragraphs_list[indiv_idx]._element
    while cur is not None:
        elements_to_remove.append(cur)
        cur = cur.getnext()

    # Удаляем sectPr из последнего параграфа перед удалением (если он там есть, нужно перенести)
    # Также удаляем все таблицы которые идут после титульника
    # Получаем все дочерние элементы body
    title_end_para = paragraphs_list[indiv_idx - 1]._element if indiv_idx > 0 else None

    # Получаем sectPr из конца body (последний sectPr во всём документе)
    # Это финальный sectPr который должен остаться
    sect_prs = body.findall(wtag('sectPr'))
    final_sect_pr = None
    if sect_prs:
        final_sect_pr = sect_prs[-1]
        # Снимем его пока — добавим в конец нашего документа после вставки контента
        final_sect_pr.getparent().remove(final_sect_pr)

    # Удаляем всё после indiv_idx-1 параграфа
    # Используем элементы напрямую
    if indiv_idx > 0:
        ref_el = paragraphs_list[indiv_idx - 1]._element
    else:
        ref_el = None

    # Удаляем все следующие элементы (параграфы, таблицы)
    if ref_el is not None:
        next_el = ref_el.getnext()
    else:
        next_el = list(body)[0] if len(list(body)) > 0 else None

    deleted_count = 0
    while next_el is not None:
        # Также проверяем sectPr внутри pPr — это разрыв секции внутри параграфа
        # Просто удаляем элемент целиком
        nxt = next_el.getnext()
        body.remove(next_el)
        deleted_count += 1
        next_el = nxt

    print(f'  Удалено элементов после титульника: {deleted_count}')

    # 3. Загружаем оригинальный контент студента
    orig_paragraphs = parse_student_paragraphs(student)
    print(f'  Параграфов в оригинале: {len(orig_paragraphs)}')

    # Пропускаем первые параграфы оригинала которые — это шапка титульника
    # Ищем где начинается реальный контент: "Задания  по летней практике" или "Задание 1"
    skip_until = None
    for i, (idx, style, text) in enumerate(orig_paragraphs):
        if 'Задание' in text and 'летней практике' in text:
            skip_until = i
            break
        if 'Задание 1' in text:
            skip_until = i
            break

    if skip_until is None:
        skip_until = 0
    print(f'  Пропускаем первые {skip_until} параграфов (шапка титульника)')

    content_paras = orig_paragraphs[skip_until:]

    # 4. Добавляем оригинальный контент после титульной страницы
    # Сначала вставим разрыв страницы
    page_break_para = etree.Element(wtag('p'))
    page_break_run = etree.SubElement(page_break_para, wtag('r'))
    br = etree.SubElement(page_break_run, wtag('br'))
    br.set(wtag('type'), 'page')

    # Найдем где сейчас последний элемент body
    body.append(page_break_para)

    # Добавляем все параграфы оригинального контента
    for (idx, style, text) in content_paras:
        # Чистим стиль — наш доступный список стилей шаблона
        # Маппинг стилей
        if 'Heading' in style:
            p_el = mkpara(text, style=style, bold=True)
        elif style == 'p2':
            p_el = mkpara(text, style='p2')
        elif 'List' in style:
            p_el = mkpara(text, style='List Paragraph')
        else:
            p_el = mkpara(text, style='Normal')
        body.append(p_el)

    # Возвращаем sectPr в самый конец
    if final_sect_pr is not None:
        body.append(final_sect_pr)

    # 5. Сохраняем
    doc.save(out_path)
    print(f'  ✓ Сохранено: {out_path}')

# ----------------------------------------------------------------
for student in STUDENTS:
    process(student)

print("\nГотово!")
