#!/usr/bin/env python3
"""
Запускать на сервере в /opt/gnbot/:
  python3 setup_templates.py

Создаёт template_dogovor_new.docx и template_schet_new.docx
из существующих файлов на сервере.
"""
import os, sys

try:
    from docx import Document
    from docx.text.paragraph import Paragraph as DocP
except ImportError:
    print("Устанавливаю python-docx...")
    os.system("pip install python-docx --break-system-packages")
    from docx import Document
    from docx.text.paragraph import Paragraph as DocP

def replace_all(body, pairs):
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = DocP(child, body)
            full = ''.join(r.text for r in para.runs)
            nw = full
            for old, new in pairs:
                if old: nw = nw.replace(old, new)
            if nw != full and para.runs:
                para.runs[0].text = nw
                for r in para.runs[1:]: r.text = ''
        elif tag in ('tbl', 'tr', 'tc', 'body'):
            replace_all(child, pairs)

def fix_bik_standalone(body):
    """Заменяем 'БИК ' (пустой, заказчик) на маркер, не трогая 'БИК 044525974' исполнителя"""
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = DocP(child, body)
            if para.text.strip() in ('БИК', 'БИК '):
                if para.runs:
                    para.runs[0].text = 'БИК [[БИК]]'
                    for r in para.runs[1:]: r.text = ''
        elif tag in ('tbl', 'tr', 'tc', 'body'):
            fix_bik_standalone(child)

def fix_zak_standalone(body):
    """Заменяем строку содержащую ТОЛЬКО 'Заказчик' на маркер"""
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = DocP(child, body)
            if para.text.strip() == 'Заказчик':
                if para.runs:
                    para.runs[0].text = '[[ЗАК]]'
                    for r in para.runs[1:]: r.text = ''
        elif tag in ('tbl', 'tr', 'tc', 'body'):
            fix_zak_standalone(child)

# ─── ДОГОВОР ───────────────────────────────────────────────────────────────
print("Создаю template_dogovor_new.docx...")

# Ищем последний сгенерированный договор как основу
import glob
candidates = glob.glob('Договор_*.docx')
if not candidates:
    print("ОШИБКА: Не найден ни один Договор_*.docx в /opt/gnbot/")
    print("Сначала запусти бота и создай хотя бы один тестовый договор,")
    print("потом снова запусти этот скрипт.")
    sys.exit(1)

src_dog = sorted(candidates)[-1]
print(f"  Использую {src_dog} как основу...")

doc = Document(src_dog)
pairs_dog = [
    # Исполнитель — исправляем если устаревший
    ('ИНН 032386861274', 'ИНН 032315540193'),
    ('р/с 40802810000009415686', 'р/с 40802810500007432200'),
    ('ochrvvv@mail.ru', 'msk@goodnight.show'),
    # Номер договора
    ('на оказание услуг № 11', 'на оказание услуг № [[НОМ]]'),
    # Дата создания (разбита по runs — объединяем через полный текст параграфа)
    ('«2» марта 2026', '«[[ДЕНЬ]]» [[МЕС_ГОД]]'),
    ('«2» сентября 2025', '«[[ДЕНЬ]]» [[МЕС_ГОД]]'),
    ('«2» октября 2025', '«[[ДЕНЬ]]» [[МЕС_ГОД]]'),
    # Директор заказчика — в шапке (подставляется ФИО из карточки)
    # Ищем любое "действующего на основании Устава" и берём всё до него
]

# Заменяем через полный текст параграфов
replace_all(doc.element.body, pairs_dog)

# Директор — ищем строку с "Генерального Директора"
import re
for para in doc.paragraphs:
    if 'Генерального Директора' in para.text and 'действующего' in para.text:
        full = ''.join(r.text for r in para.runs)
        new = re.sub(
            r'Генерального Директора\s+[А-ЯЁ][^\,]+,\s*действующего',
            'Генерального Директора [[ДИР]], действующего',
            full
        )
        if new != full and para.runs:
            para.runs[0].text = new
            for r in para.runs[1:]: r.text = ''
        break

# Динамические поля — заменяем через regex по параграфам
for para in doc.paragraphs:
    full = ''.join(r.text for r in para.runs)
    new = full

    # Длительность (вида "– 4 часа;")
    new = re.sub(r'(продолжительность оказания Услуг\s*[–\-]\s*)[\d,\.]+ час[а-я]*;',
                 r'\1[[ДЛИТ]];', new)

    # Срок (вида "27 марта 2026 года 16:00 - 20:00")
    new = re.sub(r'(Срок оказания Услуг\s*[–\-]\s*)\d+\s+[а-яё]+\s+\d{4}\s+года\s+[\d:]+\s*[-–]\s*[\d:]+;',
                 r'\1[[ДАТА_МЕР]];', new)

    # Место
    new = re.sub(r'(Место проведения:\s*)([^\n]+)',
                 r'\1[[МЕСТО]]', new)

    # Стоимость (число в скобках со словами)
    new = re.sub(r'(\bсоставляет\s+)[\d\s\xa0]+\([^)]+\)(\s+рублей)',
                 r'\1[[СУМ_Ц]] ([[СУМ_СЛ]])\2', new)

    if new != full and para.runs:
        para.runs[0].text = new
        for r in para.runs[1:]: r.text = ''

# Реквизиты заказчика в nested table
replace_all(doc.element.body, [
    ('Юридический адрес:', 'Юридический адрес:'),  # сохраняем метку
    ('Фактический адрес:', 'Фактический адрес:'),
    ('ОГРН ', 'ОГРН [[ОГРН]]'),
    ('ИНН / КПП /', 'ИНН [[ИНН]] / КПП [[КПП]]'),
    ('Расчетный счет №', 'Расчетный счет № [[РС]]'),
    ('к/с №', 'к/с № [[КС]]'),
])

# Адреса заказчика (убираем конкретные адреса прошлых клиентов)
import re
def replace_para_regex(body, pattern, repl):
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = DocP(child, body)
            full = ''.join(r.text for r in para.runs)
            new = re.sub(pattern, repl, full)
            if new != full and para.runs:
                para.runs[0].text = new
                for r in para.runs[1:]: r.text = ''
        elif tag in ('tbl', 'tr', 'tc', 'body'):
            replace_para_regex(child, pattern, repl)

replace_para_regex(doc.element.body,
    r'(Юридический адрес:\s*)[\d\s\w.,«»"]+',
    r'\1[[ЗАК_АДР]]')
replace_para_regex(doc.element.body,
    r'(Фактический адрес:\s*)[\d\s\w.,«»"]+',
    r'\1[[ЗАК_АДР]]')

fix_bik_standalone(doc.element.body)
fix_zak_standalone(doc.element.body)

# Подписи
replace_all(doc.element.body, [
    ('Генеральный директор Заказчик', 'Генеральный директор [[ЗАК]]'),
])
replace_para_regex(doc.element.body,
    r'(_+)([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.)$',
    r'\1[[ДИР_ИНИ]]')

# Заменяем ФИО директора в подписях (строка только из ФИО)
for para in doc.paragraphs:
    t = para.text.strip()
    if re.match(r'^[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+$', t):
        # Одиночная строка с ФИО — скорее всего подпись директора
        if para.runs:
            para.runs[0].text = '[[ДИР]]'
            for r in para.runs[1:]: r.text = ''

doc.save('template_dogovor_new.docx')
print("  template_dogovor_new.docx создан!")

# ─── СЧЁТ ────────────────────────────────────────────────────────────────────
print("Создаю template_schet_new.docx...")

if not os.path.exists('template_schet.docx'):
    print("ОШИБКА: template_schet.docx не найден в /opt/gnbot/")
    sys.exit(1)

doc_s = Document('template_schet.docx')
pairs_sch = [
    ('ИП ЭРДЫНЕЕВ ГЭСЭР БУЯНТУЕВИЧ', 'ИП Очирова Оксана Эдуардовна'),
    ('670011, РОССИЯ, РЕСП БУРЯТИЯ, Г УЛАН-УДЭ, МКР 142-Й, -, Д 4, КВ 18',
     '670031, РОССИЯ, РЕСП БУРЯТИЯ, Г УЛАН-УДЭ, ПР СТРОИТЕЛЕЙ, Д 62, КВ 49'),
    ('ЮЛ, ИНН, КПП, АДРЕС', '[[ЗАК_ПОЛН]]'),
    ('Счет на оплату № 304', 'Счет на оплату № [[НОМ]]'),
    ('от 10 сентября 2025 г.', 'от [[ДАТА_СЧЕТ]] г.'),
    ('Организация мероприятия 10.10.2025', 'Организация мероприятия [[ДАТА_МЕР]]'),
    ('1 000,00', '[[СУМ_Ц]]'),
    ('1\xa0000,00', '[[СУМ_Ц]]'),
    ('1000,00 руб.', '[[СУМ_Ц]] руб.'),
    ('сумму 1000,00', 'сумму [[СУМ_Ц]]'),
    ('Одна тысяча рублей 00 копеек', '[[СУМ_СЛ]]'),
]
replace_all(doc_s.element.body, pairs_sch)
doc_s.save('template_schet_new.docx')
print("  template_schet_new.docx создан!")

print("\nГотово! Перезапусти бот:")
print("  systemctl restart gnbot && systemctl status gnbot")
