import os
from generate_kp import make_kp_pdf
import re
import zipfile
import shutil
import tempfile
import subprocess
import json
from datetime import date
from docx import Document as DocxDocument
from docx.text.paragraph import Paragraph as DocxParagraph
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes, ConversationHandler
)
from telegram.error import BadRequest

TOKEN = os.environ.get("BOT_TOKEN", "")

# ── Контроль доступа ──────────────────────────────────────────
ADMIN_ID = 1179401257
USERS_FILE = "/opt/gnbot/allowed_users.json"

def load_allowed() -> set:
    if not os.path.exists(USERS_FILE):
        return {ADMIN_ID}
    with open(USERS_FILE, "r") as f:
        return set(json.load(f))

def save_allowed(users: set):
    with open(USERS_FILE, "w") as f:
        json.dump(list(users), f)

def is_allowed(user_id: int) -> bool:
    return user_id in load_allowed()
# ─────────────────────────────────────────────────────────────

TEMPLATE_BIG   = "template_big.pptx"
TEMPLATE_SMALL = "template_small.pptx"
TEMPLATE_VYEZD = "template_vyezd.pptx"
TEMPLATE_DOGOVOR = "template_dogovor_new.docx"
TEMPLATE_SCHET   = "template_schet_new.docx"
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
DOC_MENU = 10  # menu /docs
(DOC_NUM, DOC_DATE_EVENT, DOC_TIME, DOC_ADDR,
 DOC_PRICE, DOC_PAY_DATE, DOC_CARD_CHOICE, DOC_CARD,
 DOC_DIRECTOR, DOC_CONFIRM) = range(11, 21)  # DOC_DUR убран — считается автоматически

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


# Утилиты

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
    male_names = ['никита', 'андрей', 'алексей', 'сергей', 'дмитрий', 'максим',
                  'артём', 'артем', 'иван', 'михаил', 'александр', 'владимир',
                  'кирилл', 'роман', 'денис', 'евгений', 'игорь', 'олег',
                  'антон', 'виктор', 'геннадий', 'константин', 'юрий', 'павел',
                  'тимур', 'руслан', 'марат', 'данил', 'данила', 'дандар', 'гэсэр']
    if low.endswith('ия'): return n[:-2] + 'ии'
    if low.endswith('ья'): return n[:-2] + 'ьи'
    if low.endswith('ея'): return n[:-2] + 'еи'
    if low.endswith('я'):  return n[:-1] + 'и'
    if low.endswith('а'):
        if low[-2] in 'гкхжшщч':
            return n[:-1] + 'и'
        return n[:-1] + 'ы'
    if low.endswith('ь'):  return n[:-1] + 'и'
    for mn in male_names:
        if low == mn:
            return n + 'а'
    vowels = 'аеёиоуыэюяaeiouy'
    if low and low[-1] not in vowels:
        return n + 'а'
    return name

def parse_card(text):
    """Парсит карточку предприятия из текста.
    Обрабатывает как txt-файлы, так и вставленный текст.
    Работает даже без пробелов между метками и значениями."""
    data = {}
    t = text.strip()

    # Название организации — несколько стратегий по убыванию точности:
    # 1) Аббревиатура в кавычках: ООО "АМБУШСТОР" / ООО «Ромашка»
    m = re.search(
        r'((?:ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ|АКЦИОНЕРНОЕ ОБЩЕСТВО|'
        r'ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО|ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ|'
        r'ООО|ОАО|ЗАО|АО|ПАО|ИП)\s*[«""]?[А-ЯЁA-Za-z0-9«""\'«][^»""\n]{0,80}[»""]?)',
        t, re.I
    )
    if m:
        data['name'] = m.group(1).strip().rstrip(',. ')
    else:
        # 2) Первая непустая строка целиком — если она длиннее 3 слов и не похожа на метку
        first_line = t.split('\n')[0].strip()
        if (len(first_line) > 5
                and not re.match(r'^(ИНН|КПП|ОГРН|БИК|р/?с|к/?с|Банк|Счёт|Счет|Адрес)', first_line, re.I)):
            data['name'] = first_line.rstrip(',. ')

    # ИНН (10 или 12 цифр)
    m = re.search(r'ИНН\D{0,3}(\d{10,12})', t, re.I)
    if m:
        data['inn'] = m.group(1)

    # КПП (9 цифр)
    m = re.search(r'КПП\D{0,3}(\d{9})', t, re.I)
    if m:
        data['kpp'] = m.group(1)

    # ОГРН (13 или 15 цифр)
    m = re.search(r'ОГРН\D{0,5}(\d{13,15})', t, re.I)
    if m:
        data['ogrn'] = m.group(1)

    # БИК (9 цифр)
    m = re.search(r'БИК\D{0,3}(\d{9})', t, re.I)
    if m:
        data['bik'] = m.group(1)

    # Расчётный счёт (20 цифр)
    m = re.search(r'(?:расчетный\s+счет|р/?с|р\.\s*сч|Сч\.|Счёт|Счет)\D{0,8}(\d{20})', t, re.I)
    if m:
        data['rs'] = m.group(1)

    # Корреспондентский счёт (20 цифр)
    m = re.search(r'(?:к/?с|корр\w*)\D{0,8}(\d{20})', t, re.I)
    if m:
        data['ks'] = m.group(1)

    # Банк — сначала по метке, потом по вхождению слова банк
    m = re.search(r'(?:Наименование\s+банка|Банк\s+получателя)\s*[:\n]?\s*(.+?)(?:\n|БИК|$)', t, re.I | re.DOTALL)
    if m:
        data['bank'] = m.group(1).strip()[:100]
    else:
        m = re.search(r'([^\n]*[Бб]анк[^\n]{3,80})', t)
        if m:
            data['bank'] = m.group(1).strip()[:100]

    # Генеральный директор — несколько паттернов, включая без пробела и следующую строку
    for pat in [
        r'(?:Генеральный\s+директор|Ген\.?\s*дир\.?)\s*[:/\n]?\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:Генеральный\s+директор|Ген\.?\s*дир\.?)([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:Генеральный\s+директор|Директор)[^\n]*\n\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)',
        r'(?:Директор|директор)\s*/[^/]*/\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:^|\n)\s*Директор\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
    ]:
        m = re.search(pat, t, re.I | re.MULTILINE)
        if m:
            data['director'] = m.group(1).strip()
            break

    # Адрес (юридический или фактический)
    for addr_pat in [
        r'(?:Юридический\s+адрес|Юр\.?\s*адрес)[:\s]+([^\n]{10,120})',
        r'(?:^|\n)\s*Адрес[:\s]+([^\n]{10,120})',
        r'(\d{6},?\s*(?:г\.|город|г)\s*[А-ЯЁ][а-яё]+[^\n]{5,100})',
    ]:
        m = re.search(addr_pat, t, re.I | re.MULTILINE)
        if m:
            data['address'] = m.group(1).strip().rstrip(',. ')
            break

    return data


# Клавиатуры

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
    # Только текст — .doc и .txt не поддерживаются (бинарный формат)
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✏️ Вставить текст карточки", callback_data="card_text")],
        [InlineKeyboardButton("❌ Отмена",                  callback_data="cancel")],
    ])

def kb_doc_confirm():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Создать документы",  callback_data="doc_confirm_yes")],
        [InlineKeyboardButton("🔄 Начать заново",      callback_data="doc_confirm_no")],
        [InlineKeyboardButton("❌ Отмена",             callback_data="cancel")],
    ])


# Генерация файлов

def replace_shape_text(xml_str, shape_name, new_text, sz="2400"):
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


def normalize_docx_xml(xml):
    """Объединяет текст всех w:t элементов в первый внутри каждого w:p.
    Это позволяет надёжно делать str.replace даже если текст разбит по разным runs."""
    t_re = re.compile(r'<w:t(\s[^>]*)?>(.*?)</w:t>', re.DOTALL)
    result = []
    pos = 0
    search_from = 0
    para_re = re.compile(r'<w:p[ >]')
    end_re   = re.compile(r'</w:p>')
    while True:
        ms = para_re.search(xml, search_from)
        if not ms:
            break
        me = end_re.search(xml, ms.end())
        if not me:
            break
        p_start, p_end = ms.start(), me.end()
        para = xml[p_start:p_end]
        ts = list(t_re.finditer(para))
        if len(ts) > 1:
            combined = ''.join(m.group(2) for m in ts)
            for i in range(len(ts)-1, -1, -1):
                m = ts[i]
                if i == 0:
                    sp = ' xml:space="preserve"' if combined != combined.strip() or '  ' in combined else ''
                    nt = f'<w:t{sp}>{combined}</w:t>'
                else:
                    nt = '<w:t/>'
                para = para[:m.start()] + nt + para[m.end():]
            result.append(xml[pos:p_start])
            result.append(para)
            pos = p_end
        search_from = p_end
    result.append(xml[pos:])
    return ''.join(result)


def num_to_words_short(price_str):
    """Число прописью без суффикса 'рублей 00 копеек' — для вставки в скобках."""
    full = num_to_words(price_str.replace(' ', ''))
    for sfx in [' рублей 00 копеек', ' рублей', ' 00 копеек']:
        if full.endswith(sfx):
            return full[:-len(sfx)]
    return full


def calc_duration(time_str):
    """Считает продолжительность из строки вида '12:00-22:00' или '12:00 — 20:30'."""
    times = re.findall(r'(\d{1,2}):(\d{2})', time_str)
    if len(times) >= 2:
        start = int(times[0][0]) * 60 + int(times[0][1])
        end   = int(times[1][0]) * 60 + int(times[1][1])
        diff  = end - start
        if diff > 0:
            h, m = diff // 60, diff % 60
            if m == 0:
                sfx = 'час' if h == 1 else 'часа' if h in [2, 3, 4] else 'часов'
                return f'{h} {sfx}'
            total = diff / 60
            sfx = 'часа' if total < 5 else 'часов'
            return f'{total:.1f} {sfx}'.replace('.', ',')
    return None


def _date_word(date_str):
    """'23.06.2026' → '23 июня 2026'"""
    months = ['','января','февраля','марта','апреля','мая','июня',
              'июля','августа','сентября','октября','ноября','декабря']
    try:
        d, mo, y = date_str.split('.')
        return f'{int(d)} {months[int(mo)]} {y}'
    except:
        return date_str

def build_kp(data):
    loc = data['location']
    name_gen = data['name']
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

    s1 = open(os.path.join(work_dir, 'ppt/slides/slide1.xml'), encoding='utf-8').read()
    s1 = re.sub(r'Программа для [А-Яа-яёЁ]+', f'Программа для {name_gen}', s1)
    s1 = s1.replace('Программа для имя', f'Программа для {name_gen}')
    open(os.path.join(work_dir, 'ppt/slides/slide1.xml'), 'w', encoding='utf-8').write(s1)

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

    pdf = convert_to_pdf(pptx_path, tmp_dir)
    if pdf:
        return pdf, fname_base + ".pdf", tmp_dir
    return pptx_path, fname_base + ".pptx", tmp_dir

def docx_replace(xml, old, new):
    return xml.replace(old, new)

def ensure_docx(template_path, tmp_dir):
    """Гарантирует что файл — валидный docx (zip). Если .doc — конвертирует через LibreOffice."""
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Шаблон не найден: {template_path}")
    try:
        with zipfile.ZipFile(template_path, 'r'):
            pass
        return template_path
    except zipfile.BadZipFile:
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'docx', '--outdir', tmp_dir, template_path],
            timeout=60, capture_output=True
        )
        base = os.path.splitext(os.path.basename(template_path))[0]
        converted = os.path.join(tmp_dir, base + '.docx')
        if os.path.exists(converted):
            return converted
        raise Exception(f"Не удалось конвертировать {template_path}: {result.stderr.decode()}")

def find_doc_xml(work_dir):
    """Ищет document.xml в любом месте внутри распакованного docx"""
    for root, dirs, files in os.walk(work_dir):
        if 'document.xml' in files:
            return os.path.join(root, 'document.xml')
    raise FileNotFoundError(f"document.xml не найден в {work_dir}")


def post_process_docx_xml(xml):
    """Убирает только жёлтую подсветку текста."""
    xml = re.sub(r'<w:highlight[^/]*/>', '', xml)
    xml = re.sub(r'<w:highlight\b[^>]*/>', '', xml)
    return xml


def _apply_replacements(body, pairs):
    """Применяет замены во всех параграфах, включая вложенные таблицы.
    Объединяет все runs параграфа, делает замену, кладёт в первый run."""
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = DocxParagraph(child, body)
            full = ''.join(r.text for r in para.runs)
            nw = full
            for old, new in pairs:
                if old:
                    nw = nw.replace(old, new)
            if nw != full and para.runs:
                para.runs[0].text = nw
                for r in para.runs[1:]:
                    r.text = ''
        elif tag in ('tbl', 'tr', 'tc', 'body'):
            _apply_replacements(child, pairs)

def build_docs(data):
    """Генерирует договор, счёт и акт используя python-docx с [[МАРКЕРАМИ]]."""
    card         = data.get('card', {})
    company_name = card.get('name', '')
    inn          = card.get('inn', '')
    kpp          = card.get('kpp', '')
    ogrn         = card.get('ogrn', '')
    bik_zak      = card.get('bik', '')
    ks_zak       = card.get('ks', '')
    rs_zak       = card.get('rs', '')
    address_zak  = card.get('address', '')
    director     = data.get('director', '')

    doc_num    = data['doc_num']
    # Дата договора — всегда автоматически сегодняшняя
    today      = date.today().strftime('%d.%m.%Y')
    date_event = data['date_event']
    time_event = data['time_event']
    duration   = data.get('duration', '—')
    address    = data['address']
    price      = data['price']

    price_words       = num_to_words(price.replace(' ', ''))
    price_short       = num_to_words_short(price)
    today_day         = today.split('.')[0].lstrip('0')   # «2» вместо «02»
    today_month_year  = _month_year(today)
    date_srок         = f'{_date_word(date_event)} года {time_event}'
    date_schet_word   = _date_word(today)
    price_fmt         = _fmt_price(price)

    # Полная строка заказчика для счёта: "ООО Название, ИНН ..., КПП ..., адрес"
    zak_polnaya_parts = [p for p in [company_name,
                                      f'ИНН {inn}' if inn else '',
                                      f'КПП {kpp}' if kpp else '',
                                      address_zak] if p]
    zak_polnaya = ', '.join(zak_polnaya_parts) if zak_polnaya_parts else '—'

    tmp_dir = tempfile.mkdtemp()
    results = []

    # ── ДОГОВОР (XML-подход для надёжной замены разбитых runs) ─────────────────
    try:
        wdir1 = os.path.join(tmp_dir, 'dog_work')
        os.makedirs(wdir1)
        with zipfile.ZipFile(TEMPLATE_DOGOVOR, 'r') as z:
            z.extractall(wdir1)
        xml_path1 = find_doc_xml(wdir1)
        with open(xml_path1, encoding='utf-8') as f:
            xml1 = f.read()
        xml1 = normalize_docx_xml(xml1)
        xml1 = post_process_docx_xml(xml1)
        bank_zak = card.get('bank', '')
        bank_zak = card.get('bank', '')
        dog_pairs = [
            # Маркеры шаблона — только подстановка, форматирование не трогаем
            ('[[НОМ]]',       doc_num),
            ('[[ДЕНЬ]]',      today_day),
            ('[[МЕС_ГОД]]',   today_month_year),
            ('[[ДИР]]',       director),
            ('[[ДЛИТ]]',      duration),
            ('[[ДАТА_МЕР]]',  date_srок),
            ('[[МЕСТО]]',     address),
            ('[[СУМ_Ц]]',     price),
            ('[[СУМ_СЛ]]',    price_short),
            ('[[ЗАК]]',       company_name),
            ('[[ОГРН]]',      ogrn or '—'),
            ('[[ИНН]]',       inn),
            ('[[КПП]]',       kpp),
            ('[[БИК]]',       bik_zak),
            ('[[БАНК_ЗАК]]',  bank_zak or '—'),
            ('[[РС]]',        rs_zak),
            ('[[КС]]',        ks_zak),
            ('[[ДИР_ИНИ]]',   _initials(director)),
        ]
        for old, new in dog_pairs:
            if old:
                xml1 = xml1.replace(old, new)
        # Адрес: если нет — удаляем строки целиком
        if address_zak:
            xml1 = xml1.replace('[[ЗАК_АДР]]', address_zak)
        else:
            xml1 = re.sub(r'<w:p[ >](?:(?!</w:p>).)*\[\[ЗАК_АДР\]\](?:(?!</w:p>).)*</w:p>',
                          '', xml1, flags=re.DOTALL)
        if ogrn:
            xml1 = re.sub(re.escape(ogrn) + r'\d+', ogrn, xml1)
        with open(xml_path1, 'w', encoding='utf-8') as f:
            f.write(xml1)
        dog_path = os.path.join(tmp_dir, f'Договор_{doc_num}.docx')
        with zipfile.ZipFile(dog_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, dirs, files in os.walk(wdir1):
                for fn in files:
                    fp = os.path.join(root, fn)
                    z.write(fp, os.path.relpath(fp, wdir1))
        results.append((dog_path, f'Договор_{doc_num}.docx'))
    except Exception as e:
        raise Exception(f"Ошибка договора: {e}")

    # ── СЧЁТ (XML-подход для надёжной замены) ──────────────────────────────────
    try:
        wdir2 = os.path.join(tmp_dir, 'sch_work')
        os.makedirs(wdir2)
        with zipfile.ZipFile(TEMPLATE_SCHET, 'r') as z:
            z.extractall(wdir2)
        xml_path2 = find_doc_xml(wdir2)
        with open(xml_path2, encoding='utf-8') as f:
            xml2 = f.read()
        xml2 = normalize_docx_xml(xml2)
        xml2 = post_process_docx_xml(xml2)
        sch_pairs = [
            ('[[НОМ]]',       doc_num),
            ('[[ДАТА_СЧЕТ]]', date_schet_word),
            ('[[ДАТА_МЕР]]',  date_event),
            ('[[ЗАК_ПОЛН]]',  zak_polnaya),
            # Отдельные маркеры на случай если шаблон использует их раздельно
            ('[[ЗАК]]',       company_name),
            ('[[ИНН]]',       inn),
            ('[[КПП]]',       kpp),
            ('[[СУМ_Ц]]',     price_fmt),
            ('[[СУМ_СЛ]]',    price_words),
            # Старый адрес исполнителя в шаблоне счёта
            ('670011, РОССИЯ, РЕСП БУРЯТИЯ, Г УЛАН-УДЭ, МКР 142-Й, -, Д 4, КВ 18',
             '670031, РОССИЯ, РЕСП БУРЯТИЯ, Г УЛАН-УДЭ, ПР СТРОИТЕЛЕЙ, Д 62, КВ 49'),
        ]
        for old, new in sch_pairs:
            if old:
                xml2 = xml2.replace(old, new)
        with open(xml_path2, 'w', encoding='utf-8') as f:
            f.write(xml2)
        sch_path = os.path.join(tmp_dir, f'Счёт_{doc_num}.docx')
        with zipfile.ZipFile(sch_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, dirs, files in os.walk(wdir2):
                for fn in files:
                    fp = os.path.join(root, fn)
                    z.write(fp, os.path.relpath(fp, wdir2))
        sch_pdf = convert_to_pdf(sch_path, tmp_dir)
        results.append((sch_pdf, f'Счёт_{doc_num}.pdf') if sch_pdf
                       else (sch_path, f'Счёт_{doc_num}.docx'))
    except Exception as e:
        raise Exception(f"Ошибка счёта: {e}")

    # ── АКТ (XML подход) ────────────────────────────────────────────────────────
    try:
        wdir3 = os.path.join(tmp_dir, 'akt_work')
        os.makedirs(wdir3)
        akt_src = ensure_docx(TEMPLATE_AKT, tmp_dir)
        with zipfile.ZipFile(akt_src, 'r') as z:
            z.extractall(wdir3)
        xml_path3 = find_doc_xml(wdir3)
        with open(xml_path3, encoding='utf-8') as f:
            xml3 = f.read()
        xml3 = normalize_docx_xml(xml3)
        xml3 = post_process_docx_xml(xml3)

        # Строки исполнителя — фиксированные
        isp_str = f'{ISPOLNITEL["name"]}, ИНН {ISPOLNITEL["inn"]}'
        # Строка заказчика — из карточки
        zak_str = company_name
        if inn:
            zak_str += f', ИНН {inn}'

        akt_repl = [
            # Маркеры (если шаблон уже обновлён setup_templates)
            ('[[ЗАК]]',   company_name),
            ('[[ИНН]]',   inn),
            # Хардкод-строки старых шаблонов (обратная совместимость)
            ('N 151 от «10» октября',
             f'N {doc_num} от «{date_event.split(".")[0]}» {_month_only(date_event)}'),
            ('2025г.',    f'{date_event.split(".")[-1]}г.'),
            # Исполнитель (старый вариант с другим ИП)
            ('ИП Эрдынеев Гэсэр Буянтуевич, ИНН 032315540193', isp_str),
            ('ЮЛ, ИНН',  isp_str),
            # Заказчик — подставляем company_name + ИНН
            ('Заказчик, ИНН 9725189078', zak_str),
            ('Заказчик, ИНН',            zak_str),
            # Если заказчик = исполнитель (ошибка шаблона) — заменяем
            (f'Заказчик: {ISPOLNITEL["name"]}, ИНН {ISPOLNITEL["inn"]}', f'Заказчик: {zak_str}'),
            # Дата мероприятия и суммы
            ('10.10.2025',  date_event),
            ('12.12.2012',  date_event),
            ('1 000,00',    price_fmt),
            ('1\xa0000,00', price_fmt),
            ('12 000,00',   price_fmt),
            (': 1 000,00',  f': {price_fmt}'),
            ('Одна тысяча рублей 00 копеек',     price_words),
            ('Двенадцать тысяч рублей 00 копеек', price_words),
        ]
        for old, new in akt_repl:
            if old and new:
                xml3 = xml3.replace(old, new)
        with open(xml_path3, 'w', encoding='utf-8') as f:
            f.write(xml3)
        akt_docx = os.path.join(tmp_dir, f'Акт_{doc_num}.docx')
        with zipfile.ZipFile(akt_docx, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, dirs, files in os.walk(wdir3):
                for fn in files:
                    fp = os.path.join(root, fn)
                    z.write(fp, os.path.relpath(fp, wdir3))
        akt_pdf = convert_to_pdf(akt_docx, tmp_dir)
        results.append((akt_pdf, f'Акт_{doc_num}.pdf') if akt_pdf
                       else (akt_docx, f'Акт_{doc_num}.docx'))
    except Exception as e:
        raise Exception(f"Ошибка акта: {e}")

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


# КП Handlers

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    uid = user.id
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)

    if not is_allowed(uid):
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("Одобрить", callback_data=f"approve_{uid}"),
            InlineKeyboardButton("Отказать", callback_data=f"deny_{uid}"),
        ]])
        await ctx.bot.send_message(
            chat_id=ADMIN_ID,
            text=(
                f"Запрос доступа к боту\n\n"
                f"Имя: {user.full_name}\n"
                f"Username: @{user.username or 'нет'}\n"
                f"ID: {uid}"
            ),
            reply_markup=keyboard,
        )
        await update.message.reply_text(
            "Запрос на доступ отправлен администратору.\n"
            "Ожидайте одобрения — напишите /start после получения уведомления."
        )
        return

    msg = await update.message.reply_text("Главное меню:", reply_markup=kb_main())
    track(ctx, msg.message_id)


async def access_decision(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    action, uid_str = query.data.split("_", 1)
    uid = int(uid_str)
    if action == "approve":
        users = load_allowed()
        users.add(uid)
        save_allowed(users)
        await query.edit_message_text(f"Пользователь {uid} одобрен.")
        await ctx.bot.send_message(
            chat_id=uid,
            text="Доступ одобрен! Напишите /start чтобы начать работу."
        )
    elif action == "deny":
        await query.edit_message_text(f"Пользователь {uid} отклонён.")
        await ctx.bot.send_message(
            chat_id=uid,
            text="В доступе отказано. Обратитесь к администратору."
        )


async def cmd_users(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    users = load_allowed()
    text = "Допущенные пользователи:\n" + "\n".join(str(u) for u in sorted(users))
    await update.message.reply_text(text)


async def cmd_remove(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    if not ctx.args:
        await update.message.reply_text("Укажите ID: /remove 12345")
        return
    uid = int(ctx.args[0])
    users = load_allowed()
    users.discard(uid)
    save_allowed(users)
    await update.message.reply_text(f"Пользователь {uid} удалён.")


async def menu_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "menu_kp":
        ctx.user_data.clear()
        await query.edit_message_text("Для кого КП? (например: Компании Ромашка)")
        # ФИХ 2: трекаем сообщение "Как зовут клиента?" чтобы оно удалилось в конце
        track(ctx, query.message.message_id)
        return KP_NAME
    elif query.data == "menu_docs":
        await query.edit_message_text("Раздел документов:", reply_markup=kb_docs())
        track(ctx, query.message.message_id)
        return DOC_MENU
    elif query.data == "menu_all_docs":
        saved_mid = query.message.message_id   # сохраняем до clear()
        ctx.user_data.clear()
        await query.edit_message_text("Номер договора (например: 11):")
        track(ctx, saved_mid)                  # трекаем заново
        return DOC_NUM
    elif query.data == "menu_back":
        await query.edit_message_text("Главное меню:", reply_markup=kb_main())

async def cmd_kp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    msg = await update.message.reply_text("Для кого КП? (например: Компании Ромашка)")
    track(ctx, msg.message_id)
    return KP_NAME

async def kp_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['name'] = update.message.text.strip()
    # ФИХ 2: сразу пытаемся удалить сообщение пользователя с именем
    # В личном чате Telegram не позволяет удалять чужие сообщения — safe_delete молча игнорирует
    # В группе с правами администратора сообщение будет удалено
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
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
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Дата мероприятия? (например: 15.06.2026)")
    track(ctx, msg.message_id)
    return KP_DATE

async def kp_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['date'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Время начала?")
    track(ctx, msg.message_id)
    return KP_TIME

async def kp_time(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['time'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
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
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)

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
        d = ctx.user_data
        loc = d['location']
        name_gen = d['name']
        path, fname, tmp_dir = make_kp_pdf(
            loc, name_gen, d['date'], d['time'],
            d['program_lines'], d['price'],
            d.get('address', '')
        )
        with open(path, 'rb') as f:
            await q.message.reply_document(document=f, filename=fname)
        shutil.rmtree(tmp_dir, ignore_errors=True)
        await delete_tracked(ctx, chat_id)
        await safe_delete(ctx.bot, chat_id, q.message.message_id)
    except Exception as e:
        await q.message.reply_text(f"Ошибка: {e}")
    return ConversationHandler.END


# Документы Handlers

# ФИХ 3: cmd_docs теперь возвращает DOC_MENU и является entry_point для doc_conv
async def cmd_docs(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    msg = await update.message.reply_text("Раздел документов:", reply_markup=kb_docs())
    track(ctx, msg.message_id)
    return DOC_MENU

async def doc_num(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['doc_num'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Дата мероприятия? (например: 23.05.2026)")
    track(ctx, msg.message_id)
    return DOC_DATE_EVENT

async def doc_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['date_event'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Время мероприятия? (например: 19:00 — 20:30)")
    track(ctx, msg.message_id)
    return DOC_TIME

async def doc_time(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    time_str = update.message.text.strip()
    ctx.user_data['time_event'] = time_str
    # Считаем длительность автоматически из времени
    dur = calc_duration(time_str)
    ctx.user_data['duration'] = dur if dur else '—'
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Адрес проведения:")
    track(ctx, msg.message_id)
    return DOC_ADDR



async def doc_addr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['address'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Стоимость (рублей, цифрами):")
    track(ctx, msg.message_id)
    return DOC_PRICE

async def doc_price(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().replace(' ', '')
    try:
        ctx.user_data['price'] = f"{int(''.join(filter(str.isdigit, raw))):,}".replace(',', ' ')
    except:
        ctx.user_data['price'] = raw
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text(
        "Карточка предприятия заказчика:",
        reply_markup=kb_card_choice()
    )
    ctx.user_data['today'] = __import__('datetime').date.today().strftime('%d.%m.%Y')
    track(ctx, msg.message_id)
    return DOC_CARD_CHOICE

async def doc_pay_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['today'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
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
    await q.edit_message_text(
        "Вставь текст карточки предприятия (скопируй из Word или любого документа):"
    )
    return DOC_CARD

async def doc_card_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    ctx.user_data['card'] = parse_card(text)
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _after_card_parsed(update, ctx)

async def doc_card_file(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    tmp = tempfile.mktemp(suffix='.txt')
    await file.download_to_drive(tmp)
    with open(tmp, encoding='utf-8', errors='ignore') as f:
        text = f.read()
    os.remove(tmp)
    ctx.user_data['card'] = parse_card(text)
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _after_card_parsed(update, ctx)

async def _after_card_parsed(update, ctx):
    """После разбора карточки: если директор найден — пропускаем вопрос."""
    card = ctx.user_data.get('card', {})
    if card.get('director'):
        ctx.user_data['director'] = card['director']
        return await _show_doc_summary(update, ctx)
    msg = await update.message.reply_text(
        "ФИО генерального директора заказчика (полностью):"
    )
    track(ctx, msg.message_id)
    return DOC_DIRECTOR


async def _show_doc_summary(update, ctx):
    """Показывает итоговую сводку перед созданием документов."""
    card = ctx.user_data.get('card', {})
    dur = ctx.user_data.get('duration', '—')
    summary = (
        f"Проверь данные документов:\n\n"
        f"Договор №{ctx.user_data['doc_num']}\n"
        f"Дата мероприятия: {ctx.user_data['date_event']}\n"
        f"Время: {ctx.user_data['time_event']}\n"
        f"Длительность: {dur}\n"
        f"Адрес: {ctx.user_data['address']}\n"
        f"Стоимость: {ctx.user_data['price']} руб\n"
        f"Дата счёта: {ctx.user_data.get('today', __import__('datetime').date.today().strftime('%d.%m.%Y'))}\n"
        f"Заказчик: {card.get('name', '?')}\n"
        f"ИНН: {card.get('inn', '?')}\n"
        f"Директор: {ctx.user_data.get('director', '?')}"
    )
    msg = await update.message.reply_text(summary, reply_markup=kb_doc_confirm())
    track(ctx, msg.message_id)
    return DOC_CONFIRM


async def doc_director(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['director'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _show_doc_summary(update, ctx)

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
        BotCommand("docs",   "Документы"),
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

    # ФИХ 3: doc_conv теперь включает:
    # - CommandHandler("docs", cmd_docs) как entry_point
    # - DOC_MENU как первое состояние (меню с кнопками)
    # - menu_cb обрабатывает переходы внутри конверсации
    doc_conv = ConversationHandler(
        entry_points=[
            CommandHandler("docs", cmd_docs),
            CallbackQueryHandler(menu_cb, pattern=r'^menu_docs$'),
        ],
        states={
            DOC_MENU: [
                CallbackQueryHandler(menu_cb, pattern=r'^(menu_docs|menu_all_docs|menu_back|cancel)$'),
            ],
            DOC_NUM:         [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_num)],
            DOC_DATE_EVENT:  [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_date)],
            DOC_TIME:        [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_time)],
            DOC_ADDR:        [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_addr)],
            DOC_PRICE:       [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_price)],
            DOC_PAY_DATE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_pay_date)],
            DOC_CARD_CHOICE: [CallbackQueryHandler(doc_card_choice, pattern=r'^(card_text|cancel)')],
            DOC_CARD:        [
                MessageHandler(filters.TEXT & ~filters.COMMAND, doc_card_text),
            ],
            DOC_DIRECTOR:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_director)],
            DOC_CONFIRM:     [CallbackQueryHandler(doc_confirm, pattern=r'^(doc_confirm_|cancel)')],
        },
        fallbacks=[CommandHandler("cancel", cancel_handler)],
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("users", cmd_users))
    app.add_handler(CommandHandler("remove", cmd_remove))
    app.add_handler(CallbackQueryHandler(access_decision, pattern=r"^(approve|deny)_"))
    # Кнопка "Документы" из главного меню теперь обрабатывается внутри doc_conv
    app.add_handler(kp_conv)
    app.add_handler(doc_conv)
    # Кнопка "Назад" из меню документов (вне активной конверсации)
    app.add_handler(CallbackQueryHandler(menu_cb, pattern=r'^menu_back$'))

    print("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()
