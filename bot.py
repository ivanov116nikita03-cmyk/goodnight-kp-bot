import os
import re
import zipfile
import shutil
import tempfile
import subprocess
from datetime import date
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes, ConversationHandler
)
from telegram.error import BadRequest

TOKEN = os.environ.get("BOT_TOKEN", "")

TEMPLATE_BIG   = "template_big.pptx"
TEMPLATE_SMALL = "template_small.pptx"
TEMPLATE_VYEZD = "template_vyezd.pptx"
TEMPLATE_DOGOVOR = "template_dogovor.docx"
TEMPLATE_SCHET   = "template_schet.docx"
TEMPLATE_AKT     = "template_akt.docx"

# Реквизиты исполнителя (фиксированные)
ISPOLNITEL = {
    "name":    "ИП Очирова Оксана Эдуардовна",
    "inn":     "032315540193",
    "bank":    "АО «ТБанк»",
    "bik":     "044525974",
    "ks":      "30101810145250000974",
    "rs":      "40802810500007432200",
    "address": "г. Москва",
}

# Состояния КП
(KP_NAME, KP_LOC, KP_ADDR, KP_DATE, KP_TIME,
 KP_FMT, KP_PROG, KP_PRICE, KP_CONFIRM) = range(9)

# Состояния документов
(DOC_NUM, DOC_DATE_EVENT, DOC_TIME, DOC_DUR, DOC_ADDR,
 DOC_PRICE, DOC_PAY_DATE, DOC_CARD_CHOICE, DOC_CARD,
 DOC_DIRECTOR, DOC_CONFIRM) = range(11, 22)

FIXED_DUR = {"velkom": "20 мин", "break_": "10 мин"}
PROGRAM_BLOCKS = [
    ("velkom",  "Велком"),
    ("gn",      "Good Night"),
    ("break_",  "Перерыв"),
    ("kk",      "Karaoke Star"),
    ("bad",     "Bad Night 21+"),
    ("ktokogo", "Кто Кого"),
    ("arenda",  "Аренда студии"),
    ("disco",   "Дискотека с диджеем"),
    ("mafia",   "Мафия"),
]
DURATIONS = ["20 мин", "30 мин", "1 час", "1.5 часа", "2 часа"]
BASE_DUR = {
    "game":   {"default": "1.5 часа", "arenda": "1.5 часа"},
    "packet": {"default": "1 час",    "arenda": "30 мин"},
    "free":   {"default": "1 час",    "arenda": "1 час"},
}


# ─── Утилиты ─────────────────────────────────────────────────────────────────

async def safe_delete(bot, chat_id, message_id):
    try:
        await bot.delete_message(chat_id=chat_id, message_id=message_id)
    except BadRequest:
        pass

async def delete_tracked(ctx, chat_id):
    for mid in ctx.user_data.get('msg_ids', []):
        await safe_delete(ctx.bot, chat_id, mid)
    ctx.user_data['msg_ids'] = []

def track(ctx, *ids):
    ctx.user_data.setdefault('msg_ids', []).extend(ids)

def get_base_dur(bid, fmt):
    if bid in FIXED_DUR:
        return FIXED_DUR[bid]
    base = BASE_DUR.get(fmt, BASE_DUR["free"])
    return base.get(bid, base["default"])

def num_to_words(n):
    """Число прописью (рубли)"""
    ones = ['','один','два','три','четыре','пять','шесть','семь','восемь','девять',
            'десять','одиннадцать','двенадцать','тринадцать','четырнадцать','пятнадцать',
            'шестнадцать','семнадцать','восемнадцать','девятнадцать']
    tens = ['','','двадцать','тридцать','сорок','пятьдесят','шестьдесят','семьдесят','восемьдесят','девяносто']
    hundreds = ['','сто','двести','триста','четыреста','пятьсот','шестьсот','семьсот','восемьсот','девятьсот']
    thousands_f = ['','одна','две','три','четыре','пять','шесть','семь','восемь','девять',
                   'десять','одиннадцать','двенадцать','тринадцать','четырнадцать','пятнадцать',
                   'шестнадцать','семнадцать','восемнадцать','девятнадцать']

    try:
        n = int(str(n).replace(' ', '').replace(',', ''))
    except:
        return str(n)

    if n == 0:
        return 'ноль рублей 00 копеек'

    result = []
    if n >= 1000:
        th = n // 1000
        if th < 20:
            result.append(thousands_f[th])
        else:
            result.append(tens[th // 10])
            if th % 10:
                result.append(thousands_f[th % 10])
        if th % 100 in range(11, 20):
            result.append('тысяч')
        elif th % 10 == 1:
            result.append('тысяча')
        elif th % 10 in [2, 3, 4]:
            result.append('тысячи')
        else:
            result.append('тысяч')
        n = n % 1000

    if n >= 100:
        result.append(hundreds[n // 100])
        n = n % 100

    if n < 20:
        if n > 0:
            result.append(ones[n])
    else:
        result.append(tens[n // 10])
        if n % 10:
            result.append(ones[n % 10])

    words = ' '.join(w for w in result if w)
    words = words[0].upper() + words[1:] if words else ''
    return f"{words} рублей 00 копеек"

def make_genitive(name):
    """Склонение имени в родительный падеж"""
    n = name.strip()
    low = n.lower()
    # Популярные мужские имена на согласную (не склоняются в -а/-я форме)
    male_names = ['никита', 'андрей', 'алексей', 'сергей', 'дмитрий', 'максим',
                  'артём', 'артем', 'иван', 'михаил', 'александр', 'владимир',
                  'кирилл', 'роман', 'денис', 'евгений', 'игорь', 'олег',
                  'антон', 'виктор', 'геннадий', 'константин', 'юрий', 'павел',
                  'тимур', 'руслан', 'марат', 'данил', 'данила', 'дандар', 'гэсэр']
    # Правила
    if low.endswith('ия'): return n[:-2] + 'ии'
    if low.endswith('ья'): return n[:-2] + 'ьи'
    if low.endswith('ея'): return n[:-2] + 'еи'
    if low.endswith('я'):  return n[:-1] + 'и'
    if low.endswith('а'):
        if low[-2] in 'гкхжшщч':
            return n[:-1] + 'и'
        return n[:-1] + 'ы'
    if low.endswith('ь'):  return n[:-1] + 'и'
    # Мужские имена на согласную — добавляем -а
    for mn in male_names:
        if low == mn:
            return n + 'а'
    # По умолчанию — добавляем -а если заканчивается на согласную
    vowels = 'аеёиоуыэюяaeiouy'
    if low and low[-1] not in vowels:
        return n + 'а'
    return name

def parse_card(text):
    """Парсит карточку предприятия из текста"""
    data = {}
    patterns = {
        'name':      r'(?:ООО|ОАО|ЗАО|АО|ИП|ОБЩЕСТВО[^,\n]*)[^\n]*',
        'ogrn':      r'ОГРН[:\s]*(\d+)',
        'inn':       r'ИНН[:\s]*(\d+)',
        'kpp':       r'КПП[:\s]*(\d+)',
        'bank':      r'(?:Банк|БАНК)[:\s]*([^\n,]+)',
        'bik':       r'БИК[:\s]*(\d+)',
        'ks':        r'[кК]/[сС][:\s№]*(\d+)',
        'rs':        r'(?:Сч\.|счёт|счет|р/с)[:\s№]*(\d+)',
        'director':  r'(?:директор|Директор|ДИРЕКТОР)[^\n]*?([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
    }
    lines = text.strip().split('\n')
    for line in lines[:3]:
        line = line.strip()
        if re.match(r'(ООО|ОАО|ЗАО|АО|ИП)\s', line, re.I):
            data['name'] = line
            break
    for key, pattern in patterns.items():
        if key == 'name':
            continue
        m = re.search(pattern, text, re.I)
        if m:
            data[key] = m.group(1).strip() if m.lastindex else m.group(0).strip()
    return data


# ─── Клавиатуры ──────────────────────────────────────────────────────────────

def kb_main():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Создать КП",       callback_data="menu_kp")],
        [InlineKeyboardButton("📄 Документы",        callback_data="menu_docs")],
    ])

def kb_docs():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 Договор + Счёт + Акт", callback_data="menu_all_docs")],
        [InlineKeyboardButton("◀ Назад",                 callback_data="menu_back")],
    ])

def kb_location():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏢 Большая студия", callback_data="loc_big")],
        [InlineKeyboardButton("🏠 Малая студия",   callback_data="loc_small")],
        [InlineKeyboardButton("🚗 Выезд",          callback_data="loc_vyezd")],
        [InlineKeyboardButton("❌ Отмена",         callback_data="cancel")],
    ])

def kb_format():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🎮 Игра (1.5 ч)",    callback_data="fmt_game")],
        [InlineKeyboardButton("📦 Пакет (1 ч)",     callback_data="fmt_packet")],
        [InlineKeyboardButton("✏️ Свободный выбор", callback_data="fmt_free")],
        [InlineKeyboardButton("❌ Отмена",          callback_data="cancel")],
    ])

def kb_program(sel, fmt, dur_mode=None):
    rows = []
    n = 1
    for bid, bname in PROGRAM_BLOCKS:
        is_on = bid in sel
        is_fixed = bid in FIXED_DUR
        icon = "✅" if is_on else "☐"
        num = str(n) if is_on else " "
        dur = sel.get(bid, get_base_dur(bid, fmt)) if is_on else ""
        label = f"{icon} {num}. {bname}" + (f" — {dur}" if dur else "")
        if is_on: n += 1
        if not is_fixed and is_on and dur_mode == bid:
            rows.append([InlineKeyboardButton(
                ("▶ " if d == sel.get(bid) else "") + d,
                callback_data=f"dur_{bid}_{d}"
            ) for d in DURATIONS])
        row = [InlineKeyboardButton(label, callback_data=f"tog_{bid}")]
        if not is_fixed and is_on and dur_mode != bid:
            row.append(InlineKeyboardButton("⏱", callback_data=f"editdur_{bid}"))
        rows.append(row)
    rows.append([
        InlineKeyboardButton("✔ Готово",  callback_data="prog_done"),
        InlineKeyboardButton("❌ Отмена", callback_data="cancel"),
    ])
    return InlineKeyboardMarkup(rows)

def kb_confirm():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Создать КП",    callback_data="confirm_yes")],
        [InlineKeyboardButton("🔄 Начать заново", callback_data="confirm_no")],
        [InlineKeyboardButton("❌ Отмена",        callback_data="cancel")],
    ])

def kb_card_choice():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📎 Загрузить txt файл", callback_data="card_file")],
        [InlineKeyboardButton("✏️ Вставить текстом",  callback_data="card_text")],
        [InlineKeyboardButton("❌ Отмена",            callback_data="cancel")],
    ])

def kb_doc_confirm():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Создать документы",  callback_data="doc_confirm_yes")],
        [InlineKeyboardButton("🔄 Начать заново",      callback_data="doc_confirm_no")],
        [InlineKeyboardButton("❌ Отмена",             callback_data="cancel")],
    ])


# ─── Генерация файлов ─────────────────────────────────────────────────────────

def replace_shape_text(xml_str, shape_name, new_text, sz="1600"):
    def replacer(m):
        sp = m.group(0)
        para = (
            f'<a:p><a:r>'
            f'<a:rPr lang="ru-RU" sz="{sz}" b="1" dirty="0">'
            f'<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
            f'<a:latin typeface="Georgia"/>'
            f'</a:rPr>'
            f'<a:t xml:space="preserve">{new_text}</a:t>'
            f'</a:r></a:p>'
        )
        sp = re.sub(r'<a:p>.*?</a:p>', para, sp, flags=re.DOTALL)
        return sp
    return re.sub(
        rf'<p:sp>(?:(?!<p:sp>).)*?name="{re.escape(shape_name)}".*?</p:sp>',
        replacer, xml_str, flags=re.DOTALL
    )

def convert_to_pdf(src_path, tmp_dir):
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir, src_path],
            timeout=60, check=True, capture_output=True
        )
        base = os.path.splitext(os.path.basename(src_path))[0]
        pdf = os.path.join(tmp_dir, base + ".pdf")
        return pdf if os.path.exists(pdf) else None
    except Exception:
        return None

def build_kp(data):
    loc = data['location']
    name_gen = make_genitive(data['name'])
    date_str = data['date']
    time_str = data['time']
    prog = data['program_lines']
    price = data['price']
    address = data.get('address', '')
    is_vyezd = loc == 'vyezd'

    template = {'big': TEMPLATE_BIG, 'small': TEMPLATE_SMALL, 'vyezd': TEMPLATE_VYEZD}[loc]
    loc_label = {'big': 'Большая студия', 'small': 'Малая студия', 'vyezd': 'Выезд'}[loc]
    safe_date = date_str.replace('/', '-').replace('.', '-')
    fname_base = f"KP_{data['name']}_{safe_date}_{loc_label}"

    tmp_dir = tempfile.mkdtemp()
    pptx_path = os.path.join(tmp_dir, fname_base + ".pptx")
    work_dir = os.path.join(tmp_dir, 'work')
    os.makedirs(work_dir)

    with zipfile.ZipFile(template, 'r') as z:
        z.extractall(work_dir)

    # Слайд 1
    s1 = open(os.path.join(work_dir, 'ppt/slides/slide1.xml'), encoding='utf-8').read()
    s1 = re.sub(r'Программа для [А-Яа-яёЁ]+', f'Программа для {name_gen}', s1)
    s1 = s1.replace('Программа для имя', f'Программа для {name_gen}')
    open(os.path.join(work_dir, 'ppt/slides/slide1.xml'), 'w', encoding='utf-8').write(s1)

    # Слайд 3 — все поля 24pt
    s3 = open(os.path.join(work_dir, 'ppt/slides/slide3.xml'), encoding='utf-8').read()
    s3 = replace_shape_text(s3, 'TextBox_new_51', f'Дата: {date_str}  |  Время: {time_str}')
    s3 = replace_shape_text(s3, 'TextBox_new_54', prog)
    price_label = f'{price} руб (общая стоимость)' if is_vyezd else f'{price} руб/чел'
    s3 = replace_shape_text(s3, 'TextBox_new_55', price_label)
    addr = address if is_vyezd else 'Денисовский переулок 30, стр. 1'
    s3 = replace_shape_text(s3, 'TextBox 13', addr)
    open(os.path.join(work_dir, 'ppt/slides/slide3.xml'), 'w', encoding='utf-8').write(s3)

    with zipfile.ZipFile(pptx_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(work_dir):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, work_dir))

    # Конвертируем в PDF
    pdf = convert_to_pdf(pptx_path, tmp_dir)
    if pdf:
        return pdf, fname_base + ".pdf", tmp_dir
    return pptx_path, fname_base + ".pptx", tmp_dir

def docx_replace(xml, old, new):
    return xml.replace(old, new)

def build_docs(data):
    """Генерирует договор (docx), счёт (pdf), акт (pdf)"""
    card = data['card']
    company_name = card.get('name', 'Заказчик')
    inn = card.get('inn', '')
    kpp = card.get('kpp', '')
    ogrn = card.get('ogrn', '')
    bank_zak = card.get('bank', '')
    bik_zak = card.get('bik', '')
    ks_zak = card.get('ks', '')
    rs_zak = card.get('rs', '')
    director = data.get('director', 'Директор')

    doc_num = data['doc_num']
    today = data['today']
    date_event = data['date_event']
    time_event = data['time_event']
    duration = data['duration']
    address = data['address']
    price = data['price']
    pay_date = data['pay_date']
    price_words = num_to_words(price.replace(' ', ''))

    tmp_dir = tempfile.mkdtemp()
    results = []

    # ─── ДОГОВОР (docx) ───
    work_dir = os.path.join(tmp_dir, 'dogovor_work')
    os.makedirs(work_dir)
    with zipfile.ZipFile(TEMPLATE_DOGOVOR, 'r') as z:
        z.extractall(work_dir)

    doc_xml_path = os.path.join(work_dir, 'word/document.xml')
    with open(doc_xml_path, encoding='utf-8') as f:
        xml = f.read()

    # Убираем жёлтые выделения
    xml = xml.replace('<w:highlight w:val="yellow"/>', '')
    # Замены переменных
    replacements_doc = [
        ('136', doc_num),
        ('сентября 2025', _month_year(today)),
        ('«2»', f'«{today.split(".")[0]}»'),
        ('Общество с ограниченной ответственностью «ВК»', company_name),
        ('Общество с ограниченной ответственностью «Гарда Технологии»', company_name),
        ('Генерального Директора Багудиной Елены Геннадьевны', f'Генерального Директора {director}'),
        ('4 часа', duration),
        ('27 сентября 2025 года 16:00 - 20:00', f'{date_event} года {time_event}'),
        ('г. Красногорск, Московская область, Яхт-клуб «Парк Рублёво»', address),
        ('76 280 (семьдесят шесть тысяч двести восемьдесят)', f'{price} ({price_words})'),
        ('1027739850962', ogrn),
        ('7743001840 / 997750001', f'{inn} / {kpp}'),
        ('044525823', bik_zak),
        ('БАНК ГПБ (АО)', bank_zak),
        ('40702810100000003759', rs_zak),
        ('30101810200000000823', ks_zak),
        ('Генеральный директор управляющей организации ООО «Управляющая компания ВК»',
         f'Генеральный директор {company_name}'),
        ('Багудина Елена Геннадьевна', director),
        ('Багудина Е.Г.', _initials(director)),
    ]
    for old, new in replacements_doc:
        xml = xml.replace(old, new)

    with open(doc_xml_path, 'w', encoding='utf-8') as f:
        f.write(xml)

    dogovor_path = os.path.join(tmp_dir, f'Договор_{doc_num}_{company_name[:20]}.docx')
    with zipfile.ZipFile(dogovor_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(work_dir):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, work_dir))
    results.append((dogovor_path, f'Договор_{doc_num}.docx'))

    # ─── СЧЁТ (docx → pdf) ───
    work_dir2 = os.path.join(tmp_dir, 'schet_work')
    os.makedirs(work_dir2)
    with zipfile.ZipFile(TEMPLATE_SCHET, 'r') as z:
        z.extractall(work_dir2)

    schet_xml = os.path.join(work_dir2, 'word/document.xml')
    with open(schet_xml, encoding='utf-8') as f:
        xml2 = f.read()

    replacements_schet = [
        ('ИП ЭРДЫНЕЕВ ГЭСЭР БУЯНТУЕВИЧ', ISPOLNITEL['name']),
        ('032315540193', ISPOLNITEL['inn']),
        ('АО «ТБанк»', ISPOLNITEL['bank']),
        ('044525974', ISPOLNITEL['bik']),
        ('30101810145250000974', ISPOLNITEL['ks']),
        ('40802810500007432200', ISPOLNITEL['rs']),
        ('670011, РОССИЯ, РЕСП БУРЯТИЯ, Г УЛАН-УДЭ, МКР 142-Й, -, Д 4, КВ 18', ISPOLNITEL['address']),
        ('304 от 10 сентября 2025 г.', f'{doc_num} от {_full_date(today)} г.'),
        ('ЮЛ, ИНН, КПП, АДРЕС', f'{company_name}, ИНН {inn}, КПП {kpp}'),
        ('Организация мероприятия', 'Организация мероприятия'),
        ('10.10.2025', date_event),
        ('1 000,00', _fmt_price(price)),
        ('1\xa0000,00', _fmt_price(price)),
        ('1000,00', price.replace(' ', '')+',00'),
        ('Одна тысяча рублей 00 копеек', price_words),
    ]
    for old, new in replacements_schet:
        xml2 = xml2.replace(old, new)

    with open(schet_xml, 'w', encoding='utf-8') as f:
        f.write(xml2)

    schet_docx = os.path.join(tmp_dir, f'Счёт_{doc_num}.docx')
    with zipfile.ZipFile(schet_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(work_dir2):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, work_dir2))

    schet_pdf = convert_to_pdf(schet_docx, tmp_dir)
    if schet_pdf:
        results.append((schet_pdf, f'Счёт_{doc_num}.pdf'))
    else:
        results.append((schet_docx, f'Счёт_{doc_num}.docx'))

    # ─── АКТ (docx → pdf) ───
    work_dir3 = os.path.join(tmp_dir, 'akt_work')
    os.makedirs(work_dir3)
    with zipfile.ZipFile(TEMPLATE_AKT, 'r') as z:
        z.extractall(work_dir3)

    akt_xml = os.path.join(work_dir3, 'word/document.xml')
    with open(akt_xml, encoding='utf-8') as f:
        xml3 = f.read()

    replacements_akt = [
        ('N 151 от «10» октября', f'N {doc_num} от «{date_event.split(".")[0]}» {_month_only(date_event)}'),
        ('2025г._________________________________________________',
         f'{date_event.split(".")[-1]}г._________________________________________________'),
        ('ЮЛ, ИНН', f'{ISPOLNITEL["name"]}, ИНН {ISPOLNITEL["inn"]}'),
        ('Заказчик: ИП Эрдынеев Гэсэр Буянтуевич, ИНН 032315540193',
         f'Заказчик: {company_name}, ИНН {inn}'),
        ('Организация мероприятия', 'Организация мероприятия'),
        ('10.10.2025', date_event),
        ('1 000,00', _fmt_price(price)),
        ('1\xa0000,00', _fmt_price(price)),
        (': 1 000,00', f': {_fmt_price(price)}'),
        ('Одна тысяча рублей 00 копеек', price_words),
    ]
    for old, new in replacements_akt:
        xml3 = xml3.replace(old, new)

    with open(akt_xml, 'w', encoding='utf-8') as f:
        f.write(xml3)

    akt_docx = os.path.join(tmp_dir, f'Акт_{doc_num}.docx')
    with zipfile.ZipFile(akt_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(work_dir3):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, work_dir3))

    akt_pdf = convert_to_pdf(akt_docx, tmp_dir)
    if akt_pdf:
        results.append((akt_pdf, f'Акт_{doc_num}.pdf'))
    else:
        results.append((akt_docx, f'Акт_{doc_num}.docx'))

    return results, tmp_dir


def _month_year(date_str):
    months = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
              'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    try:
        parts = date_str.split('.')
        return f'{months[int(parts[1])]} {parts[2]}'
    except:
        return date_str

def _month_only(date_str):
    months = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
              'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    try:
        return months[int(date_str.split('.')[1])]
    except:
        return date_str

def _full_date(date_str):
    months = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
              'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    try:
        d, m, y = date_str.split('.')
        return f'{int(d)} {months[int(m)]} {y}'
    except:
        return date_str

def _fmt_price(price_str):
    try:
        n = int(price_str.replace(' ', ''))
        return f"{n:,}".replace(',', '\xa0') + ',00'
    except:
        return price_str + ',00'

def _initials(fio):
    parts = fio.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio


# ─── КП Handlers ─────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    msg = await update.message.reply_text("Главное меню:", reply_markup=kb_main())
    track(ctx, msg.message_id)

async def menu_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "menu_kp":
        ctx.user_data.clear()
        await query.edit_message_text("Как зовут клиента?")
        return KP_NAME
    elif query.data == "menu_docs":
        await query.edit_message_text("Раздел документов:", reply_markup=kb_docs())
    elif query.data == "menu_all_docs":
        ctx.user_data.clear()
        await query.edit_message_text("Номер договора (например: 11):")
        return DOC_NUM
    elif query.data == "menu_back":
        await query.edit_message_text("Главное меню:", reply_markup=kb_main())

async def cmd_kp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    msg = await update.message.reply_text("Как зовут клиента?")
    track(ctx, msg.message_id)
    return KP_NAME

async def kp_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['name'] = update.message.text.strip().capitalize()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Выбери локацию:", reply_markup=kb_location())
    track(ctx, msg.message_id)
    return KP_LOC

async def kp_location(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    if q.data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END
    loc = q.data.replace('loc_', '')
    ctx.user_data['location'] = loc
    if loc == 'vyezd':
        await q.edit_message_text("Адрес выезда:")
        return KP_ADDR
    await q.edit_message_text("Дата мероприятия? (например: 15.06.2026)")
    return KP_DATE

async def kp_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['address'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Дата мероприятия? (например: 15.06.2026)")
    track(ctx, msg.message_id)
    return KP_DATE

async def kp_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['date'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Время начала?")
    track(ctx, msg.message_id)
    return KP_TIME

async def kp_time(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['time'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    if ctx.user_data['location'] == 'vyezd':
        ctx.user_data['fmt'] = 'free'
        ctx.user_data['selected'] = {}
        ctx.user_data['dur_mode'] = None
        msg = await update.message.reply_text("Программа (⏱ меняет время):", reply_markup=kb_program({}, 'free'))
        track(ctx, msg.message_id)
        return KP_PROG
    msg = await update.message.reply_text("Формат:", reply_markup=kb_format())
    track(ctx, msg.message_id)
    return KP_FMT

async def kp_format(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    if q.data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END
    fmt = q.data.replace('fmt_', '')
    ctx.user_data.update({'fmt': fmt, 'selected': {}, 'dur_mode': None})
    hints = {'game': "🎮 Игра — 1.5 ч:", 'packet': "📦 Пакет — 1 ч, аренда 30 мин:", 'free': "✏️ Свободный:"}
    await q.edit_message_text(hints[fmt], reply_markup=kb_program({}, fmt))
    return KP_PROG

async def kp_program(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    sel = ctx.user_data.get('selected', {})
    fmt = ctx.user_data.get('fmt', 'free')
    data = q.data

    if data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END
    if data.startswith('tog_'):
        bid = data[4:]
        if bid in sel: del sel[bid]
        else: sel[bid] = get_base_dur(bid, fmt)
        if bid not in FIXED_DUR: ctx.user_data['dur_mode'] = None
    elif data.startswith('editdur_'):
        ctx.user_data['dur_mode'] = data[8:]
    elif data.startswith('dur_'):
        _, bid, dur = data.split('_', 2)
        sel[bid] = dur; ctx.user_data['dur_mode'] = None
    elif data == 'prog_done':
        if not sel:
            await q.edit_message_text("Выбери хотя бы один блок!", reply_markup=kb_program(sel, fmt))
            return KP_PROG
        ctx.user_data['selected'] = sel
        await q.edit_message_text("Стоимость? (руб/чел или общая для выезда)")
        return KP_PRICE

    await q.edit_message_reply_markup(reply_markup=kb_program(sel, fmt, ctx.user_data.get('dur_mode')))
    return KP_PROG

async def kp_price(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().replace(' ', '')
    try:
        price = f"{int(''.join(filter(str.isdigit, raw))):,}".replace(',', ' ')
    except: price = raw
    ctx.user_data['price'] = price
    track(ctx, update.message.message_id)

    sel = ctx.user_data['selected']
    fmt = ctx.user_data.get('fmt', 'free')
    n = 1; lines = []
    for bid, bname in PROGRAM_BLOCKS:
        if bid in sel:
            lines.append(f"{n}) {bname} — {sel[bid]}"); n += 1
    ctx.user_data['program_lines'] = '\n'.join(lines)

    loc = {'big': 'Большая студия', 'small': 'Малая студия', 'vyezd': 'Выезд'}[ctx.user_data['location']]
    is_vyezd = ctx.user_data['location'] == 'vyezd'
    summary = (
        f"Проверь КП:\n\nКлиент: {ctx.user_data['name']}\nЛокация: {loc}\n"
        f"Адрес: {ctx.user_data.get('address', 'Денисовский переулок 30, стр. 1')}\n"
        f"Дата: {ctx.user_data['date']}  Время: {ctx.user_data['time']}\n"
        f"Программа:\n{ctx.user_data['program_lines']}\n"
        f"Стоимость: {price} {'руб (общая)' if is_vyezd else 'руб/чел'}"
    )
    msg = await update.message.reply_text(summary, reply_markup=kb_confirm())
    track(ctx, msg.message_id)
    return KP_CONFIRM

async def kp_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    chat_id = q.message.chat_id
    if q.data in ('confirm_no', 'cancel'):
        await delete_tracked(ctx, chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END
    await q.edit_message_text("Готовлю КП...")
    try:
        path, fname, tmp_dir = build_kp(ctx.user_data)
        with open(path, 'rb') as f:
            await q.message.reply_document(document=f, filename=fname)
        shutil.rmtree(tmp_dir, ignore_errors=True)
        await delete_tracked(ctx, chat_id)
        await safe_delete(ctx.bot, chat_id, q.message.message_id)
    except Exception as e:
        await q.message.reply_text(f"Ошибка: {e}")
    return ConversationHandler.END


# ─── Документы Handlers ───────────────────────────────────────────────────────

async def doc_num(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['doc_num'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Дата мероприятия? (например: 23.05.2026)")
    track(ctx, msg.message_id)
    return DOC_DATE_EVENT

async def doc_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['date_event'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Время мероприятия? (например: 19:00 — 20:30)")
    track(ctx, msg.message_id)
    return DOC_TIME

async def doc_time(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['time_event'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Длительность? (например: 1,5 часа)")
    track(ctx, msg.message_id)
    return DOC_DUR

async def doc_dur(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['duration'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Адрес проведения:")
    track(ctx, msg.message_id)
    return DOC_ADDR

async def doc_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['address'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("Стоимость (рублей, цифрами):")
    track(ctx, msg.message_id)
    return DOC_PRICE

async def doc_price(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().replace(' ', '')
    try:
        ctx.user_data['price'] = f"{int(''.join(filter(str.isdigit, raw))):,}".replace(',', ' ')
    except:
        ctx.user_data['price'] = raw
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text(
        f"Сегодняшняя дата (для счёта)? Сегодня: {date.today().strftime('%d.%m.%Y')}"
    )
    track(ctx, msg.message_id)
    return DOC_PAY_DATE

async def doc_pay_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['today'] = update.message.text.strip()
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text(
        "Карточка предприятия заказчика:",
        reply_markup=kb_card_choice()
    )
    track(ctx, msg.message_id)
    return DOC_CARD_CHOICE

async def doc_card_choice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    if q.data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END
    ctx.user_data['card_method'] = q.data
    if q.data == "card_file":
        await q.edit_message_text("Загрузи txt файл с карточкой предприятия:")
    else:
        await q.edit_message_text("Вставь текст карточки предприятия:")
    return DOC_CARD

async def doc_card_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    ctx.user_data['card'] = parse_card(text)
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("ФИО генерального директора заказчика (полностью):")
    track(ctx, msg.message_id)
    return DOC_DIRECTOR

async def doc_card_file(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    tmp = tempfile.mktemp(suffix='.txt')
    await file.download_to_drive(tmp)
    with open(tmp, encoding='utf-8', errors='ignore') as f:
        text = f.read()
    os.remove(tmp)
    ctx.user_data['card'] = parse_card(text)
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text("ФИО генерального директора заказчика (полностью):")
    track(ctx, msg.message_id)
    return DOC_DIRECTOR

async def doc_director(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['director'] = update.message.text.strip()
    track(ctx, update.message.message_id)

    card = ctx.user_data.get('card', {})
    summary = (
        f"Проверь данные документов:\n\n"
        f"Договор №{ctx.user_data['doc_num']}\n"
        f"Дата мероприятия: {ctx.user_data['date_event']}\n"
        f"Время: {ctx.user_data['time_event']}\n"
        f"Длительность: {ctx.user_data['duration']}\n"
        f"Адрес: {ctx.user_data['address']}\n"
        f"Стоимость: {ctx.user_data['price']} руб\n"
        f"Дата счёта: {ctx.user_data['today']}\n"
        f"Заказчик: {card.get('name', '?')}\n"
        f"ИНН: {card.get('inn', '?')}\n"
        f"Директор: {ctx.user_data['director']}"
    )
    msg = await update.message.reply_text(summary, reply_markup=kb_doc_confirm())
    track(ctx, msg.message_id)
    return DOC_CONFIRM

async def doc_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    chat_id = q.message.chat_id
    if q.data in ('doc_confirm_no', 'cancel'):
        await delete_tracked(ctx, chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END

    await q.edit_message_text("Готовлю документы...")
    try:
        files, tmp_dir = build_docs(ctx.user_data)
        for fpath, fname in files:
            with open(fpath, 'rb') as f:
                await q.message.reply_document(document=f, filename=fname)
        shutil.rmtree(tmp_dir, ignore_errors=True)
        await delete_tracked(ctx, chat_id)
        await safe_delete(ctx.bot, chat_id, q.message.message_id)
    except Exception as e:
        await q.message.reply_text(f"Ошибка: {e}")
    return ConversationHandler.END


async def cancel_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    await delete_tracked(ctx, update.effective_chat.id)
    await update.message.reply_text("Отменено. Напиши /start")
    return ConversationHandler.END


async def post_init(app):
    await app.bot.set_my_commands([
        BotCommand("start",  "Главное меню"),
        BotCommand("kp",     "Создать КП"),
        BotCommand("cancel", "Отменить"),
    ])


def main():
    app = Application.builder().token(TOKEN).post_init(post_init).build()

    kp_conv = ConversationHandler(
        entry_points=[
            CommandHandler("kp", cmd_kp),
            CallbackQueryHandler(menu_cb, pattern=r'^menu_kp$'),
        ],
        states={
            KP_NAME:    [MessageHandler(filters.TEXT & ~filters.COMMAND, kp_name)],
            KP_LOC:     [CallbackQueryHandler(kp_location, pattern=r'^(loc_|cancel)')],
            KP_ADDR:    [MessageHandler(filters.TEXT & ~filters.COMMAND, kp_addr)],
            KP_DATE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, kp_date)],
            KP_TIME:    [MessageHandler(filters.TEXT & ~filters.COMMAND, kp_time)],
            KP_FMT:     [CallbackQueryHandler(kp_format,   pattern=r'^(fmt_|cancel)')],
            KP_PROG:    [CallbackQueryHandler(kp_program)],
            KP_PRICE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, kp_price)],
            KP_CONFIRM: [CallbackQueryHandler(kp_confirm,  pattern=r'^(confirm_|cancel)')],
        },
        fallbacks=[CommandHandler("cancel", cancel_handler)],
    )

    doc_conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(menu_cb, pattern=r'^menu_all_docs$'),
        ],
        states={
            DOC_NUM:         [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_num)],
            DOC_DATE_EVENT:  [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_date)],
            DOC_TIME:        [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_time)],
            DOC_DUR:         [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_dur)],
            DOC_ADDR:        [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_addr)],
            DOC_PRICE:       [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_price)],
            DOC_PAY_DATE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_pay_date)],
            DOC_CARD_CHOICE: [CallbackQueryHandler(doc_card_choice, pattern=r'^(card_|cancel)')],
            DOC_CARD:        [
                MessageHandler(filters.TEXT & ~filters.COMMAND, doc_card_text),
                MessageHandler(filters.Document.ALL, doc_card_file),
            ],
            DOC_DIRECTOR:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_director)],
            DOC_CONFIRM:     [CallbackQueryHandler(doc_confirm, pattern=r'^(doc_confirm_|cancel)')],
        },
        fallbacks=[CommandHandler("cancel", cancel_handler)],
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CallbackQueryHandler(menu_cb, pattern=r'^(menu_docs|menu_back)$'))
    app.add_handler(kp_conv)
    app.add_handler(doc_conv)

    print("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()
