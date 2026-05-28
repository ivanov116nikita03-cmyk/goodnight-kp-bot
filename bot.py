# Файл bot.py — полная версия с экраном редактирования карточки

import os
import traceback
from generate_kp import make_kp_pdf
import re
import zipfile
import shutil
import tempfile
import subprocess
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

TEMPLATE_BIG   = "template_big.pptx"
TEMPLATE_SMALL = "template_small.pptx"
TEMPLATE_VYEZD = "template_vyezd.pptx"
TEMPLATE_DOGOVOR     = "template_dogovor_new.docx"
TEMPLATE_DOGOVOR_IP  = "template_dogovor_new___IP.docx"
TEMPLATE_DOGOVOR_FIZ = "template_dogovor_new___Fiz_Liz.docx"
TEMPLATE_SCHET   = "template_schet_new.docx"
TEMPLATE_AKT     = "template_akt.docx"

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
DOC_MENU = 10
(DOC_NUM, DOC_DATE_EVENT, DOC_TIME, DOC_ADDR,
 DOC_PRICE, DOC_PAY_DATE, DOC_CARD_CHOICE, DOC_CARD,
 DOC_DIRECTOR, DOC_CONFIRM) = range(11, 21)
DOC_TYPE       = 21
DOC_FIZ_FIO    = 22
DOC_FIZ_PASS   = 23
DOC_FIZ_ISSUED = 24
DOC_FIZ_CODE   = 25
DOC_FREE_INPUT = 26   # свободный ввод данных мероприятия одним сообщением
DOC_FORM       = 26   # алиас для совместимости
DOC_FORM_EDIT  = 27   # не используется

# НОВЫЕ состояния для анкеты карточки
DOC_CARD_REVIEW = 30   # показываем анкету с кнопками
DOC_CARD_EDIT   = 31   # пользователь вводит новое значение поля

# Метки полей карточки (ключ: человекочитаемое название)
CARD_FIELDS = [
    ('name',     'Название'),
    ('inn',      'ИНН'),
    ('kpp',      'КПП'),
    ('ogrn',     'ОГРН'),
    ('bik',      'БИК'),
    ('rs',       'Р/счёт'),
    ('ks',       'К/счёт'),
    ('bank',     'Банк'),
    ('director', 'Директор'),
    ('address',  'Адрес'),
]

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
    ("vedenie", "Ведение"),
]
DURATIONS = ["20 мин", "30 мин", "1 час", "1.5 часа", "2 часа"]
BASE_DUR = {
    "game":   {"default": "1.5 часа", "arenda": "1.5 часа", "vedenie": "1 час"},
    "packet": {"default": "1 час",    "arenda": "30 мин",   "vedenie": "1 час"},
    "free":   {"default": "1 час",    "arenda": "1 час",    "vedenie": "1 час"},
}

# Фиксированная стоимость блока "Ведение" для КП
VEDENIE_PRICE_PER_HOUR = 14000


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
    if n >= 1000000:
        # Миллионы
        mn = n // 1000000
        result.extend(num_to_words(mn * 1000).replace(' тысяч рублей 00 копеек','').replace(' тысяча рублей 00 копеек','').replace(' тысячи рублей 00 копеек','').split())
        # Пересчитываем без миллионов
        result = []
        mn = n // 1000000
        if mn == 1:
            result.append('один миллион')
        elif mn in [2,3,4]:
            result.append(ones[mn] + ' миллиона')
        elif mn < 20:
            result.append(ones[mn] + ' миллионов')
        else:
            result.append(tens[mn // 10])
            if mn % 10:
                result.append(ones[mn % 10] + ' миллионов')
            else:
                result.append('миллионов')
        n = n % 1000000
    if n >= 1000:
        th = n // 1000
        if th < 20:
            result.append(thousands_f[th])
        elif th < 100:
            result.append(tens[th // 10])
            if th % 10:
                result.append(thousands_f[th % 10])
        else:
            # Сотни тысяч (100_000 — 999_000)
            result.append(hundreds[th // 100])
            th_rem = th % 100
            if th_rem < 20:
                if th_rem > 0:
                    result.append(thousands_f[th_rem])
            else:
                result.append(tens[th_rem // 10])
                if th_rem % 10:
                    result.append(thousands_f[th_rem % 10])
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

def make_genitive_fio(fio):
    """Склоняет полное ФИО (Фамилия Имя Отчество) в родительный падеж."""
    parts = fio.strip().split()
    if len(parts) < 2:
        return fio

    # Определяем пол по отчеству (3-е слово) или окончанию фамилии
    gender = 'unknown'
    if len(parts) >= 3:
        ot = parts[2].lower()
        if ot.endswith('ович') or ot.endswith('евич'):
            gender = 'm'
        elif ot.endswith('овна') or ot.endswith('евна') or ot.endswith('инична'):
            gender = 'f'
    if gender == 'unknown':
        fam = parts[0].lower()
        if fam.endswith('ов') or fam.endswith('ев') or fam.endswith('ин') or fam.endswith('ын'):
            gender = 'm'
        elif fam.endswith('ова') or fam.endswith('ева') or fam.endswith('ина'):
            gender = 'f'

    result = []
    for i, part in enumerate(parts):
        low = part.lower()

        # ОТЧЕСТВО
        if i == 2:
            if low.endswith('ович') or low.endswith('евич'):
                result.append(part + 'а'); continue
            if low.endswith('овна') or low.endswith('евна') or low.endswith('инична'):
                result.append(part[:-1] + 'ы'); continue
            result.append(part); continue

        # ИМЯ — расширенный список с неправильными формами
        if i == 1:
            irregulars = {
                'лев': 'Льва', 'пётр': 'Петра', 'павел': 'Павла',
                'семён': 'Семёна', 'семен': 'Семёна',
            }
            if low in irregulars:
                result.append(irregulars[low]); continue
            # Мужские на -ий -> -ия (Дмитрий, Василий, Георгий)
            if gender == 'm' and low.endswith('ий'):
                result.append(part[:-2] + 'ия'); continue
            result.append(make_genitive(part)); continue

        # ФАМИЛИЯ
        # -ский/-цкий (муж -> -ского, жен -> -ской)
        if low.endswith('ский') or low.endswith('цкий'):
            if gender == 'f':
                result.append(part[:-2] + 'ой')
            else:
                result.append(part[:-2] + 'ого')
            continue
        if low.endswith('ская') or low.endswith('цкая'):
            result.append(part[:-2] + 'ой'); continue
        # Мужские на -ов/-ев/-ин/-ын/-ан/-он
        if re.search(r'(?:ов|ев|ин|ын|он|ан)$', low) and gender != 'f':
            result.append(part + 'а'); continue
        # Женские на -ова/-ева/-ина -> -овой/-евой/-иной
        if re.search(r'(?:ова|ева|ина|ына)$', low):
            result.append(part[:-1] + 'ой'); continue
        # Общий случай через make_genitive
        result.append(make_genitive(part))

    return ' '.join(result)

def make_genitive(name):
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
    """
    Парсит карточку предприятия. Использует точные форматы полей:
    ИНН: 10 цифр (ЮЛ) или 12 цифр (ИП/ФЛ)
    КПП: ровно 9 цифр
    ОГРН: 13 цифр (ЮЛ) или 15 цифр (ИП)
    БИК: ровно 9 цифр, ВСЕГДА начинается с 04
    Р/с: ровно 20 цифр, начинается с 405/406/407/408/423/655
    К/с: ровно 20 цифр, ВСЕГДА начинается с 301
    """
    data = {}
    t = text.strip()

    # НАЗВАНИЕ
    m = re.search(
        r'((?:ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ|АКЦИОНЕРНОЕ ОБЩЕСТВО|'
        r'ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО|ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ|'
        r'ООО|ОАО|ЗАО|АО|ПАО|ИП)\s*[«""]?[А-ЯЁA-Za-z0-9][^»""\n]{0,80}[»""]?)',
        t, re.I
    )
    if m:
        data['name'] = m.group(1).strip().rstrip(',. ')
    else:
        first_line = t.split('\n')[0].strip()
        if (len(first_line) > 5
                and not re.match(r'^(ИНН|КПП|ОГРН|БИК|р/?с|к/?с|Банк|Счёт|Счет|Адрес)', first_line, re.I)):
            data['name'] = first_line.rstrip(',. ')

    # ИНН: 10 или 12 цифр, НЕ берём ИНН банка
    for m in re.finditer(r'(?<!\d)ИНН\s*[:/]?\s*(\d{10}|\d{12})(?!\d)', t, re.I):
        ctx_before = t[max(0, m.start()-30):m.start()].lower()
        if 'банк' not in ctx_before:
            data['inn'] = m.group(1)
            break

    # КПП: ровно 9 цифр
    m = re.search(r'КПП\s*[:/]?\s*(\d{9})(?!\d)', t, re.I)
    if m:
        data['kpp'] = m.group(1)
    if not data.get('kpp'):
        m = re.search(r'(?:ИНН\s*\d{10,12}[,/\s]+КПП\s*[:/]?\s*|/КПП[:/]?\s*)(\d{9})(?!\d)', t, re.I)
        if m:
            data['kpp'] = m.group(1)

    # ОГРН: 13 цифр (ЮЛ) или 15 цифр (ИП)
    m = re.search(r'ОГРН\w*\s*[:/]?\s*(\d{13}|\d{15})(?!\d)', t, re.I)
    if m:
        data['ogrn'] = m.group(1)

    # БИК: ровно 9 цифр, ВСЕГДА начинается с 04
    m = re.search(r'БИК\s*[:/]?\s*(04\d{7})(?!\d)', t, re.I)
    if m:
        data['bik'] = m.group(1)
    if not data.get('bik'):
        # Слитное написание: "БИК044525104" или "БИКбанка044525104"
        m = re.search(r'БИК\s*(?:банка)?\s*(04\d{7})(?!\d)', t, re.I)
        if m:
            data['bik'] = m.group(1)
    if not data.get('bik'):
        # Без метки: 9 цифр начиная с 04, рядом слово "банк" или "бик"
        for mm in re.finditer(r'(?<!\d)(04\d{7})(?!\d)', t):
            ctx = t[max(0, mm.start()-30):mm.start()+12].lower()
            if 'бик' in ctx or 'банк' in ctx:
                data['bik'] = mm.group(1)
                break

    # Р/С: ровно 20 цифр, начинается с 405/406/407/408/423/655
    rs_pref = r'(?:407|408|405|406|423|655)'
    m = re.search(
        r'(?:расчет\w*\s*счет|р\s*/\s*с|р\.с\.|р/счет|р/сч)\s*[:/]?\s*(' + rs_pref + r'\d{17})(?!\d)',
        t, re.I
    )
    if m:
        data['rs'] = m.group(1)
    if not data.get('rs'):
        for mm in re.finditer(r'(?<!\d)(' + rs_pref + r'\d{17})(?!\d)', t):
            val = mm.group(1)
            ctx = t[max(0, mm.start()-40):mm.start()+22].lower()
            if any(x in ctx for x in ['р/с', 'р.с', 'расч', 'счет', 'сч.']):
                data['rs'] = val
                break
        if not data.get('rs'):
            m = re.search(r'(?:р\s*/\s*с|р\.с\.|расч\w*\s+счет)\s*[:/]?\s*(\d{20})(?!\d)', t, re.I)
            if m:
                data['rs'] = m.group(1)

    # К/С: ровно 20 цифр, ВСЕГДА начинается с 301
    m = re.search(
        r'(?:корр\w*\s*счет|к\s*/\s*с|к\.с\.|к/счет|к/сч)\s*[:/]?\s*(301\d{17})(?!\d)',
        t, re.I
    )
    if m:
        data['ks'] = m.group(1)
    if not data.get('ks'):
        for mm in re.finditer(r'(?<!\d)(301\d{17})(?!\d)', t):
            ctx = t[max(0, mm.start()-40):mm.start()+22].lower()
            if any(x in ctx for x in ['к/с', 'к.с', 'корр']):
                data['ks'] = mm.group(1)
                break
        if not data.get('ks'):
            m = re.search(r'(?:к\s*/\s*с|корр\w*\s+счет)\s*[:/]?\s*(\d{20})(?!\d)', t, re.I)
            if m:
                data['ks'] = m.group(1)

    # БАНК: строка рядом с меткой, не цифры
    # Порядок важен: сначала точные метки, потом общий "Банк"
    bank_patterns = [
        # "Наименование банка" / "Банк получателя" с явным разделителем
        r'(?:Наименование\s+банка|Банк\s+получателя)\s*[:\n]\s*([^\n]{3,100}?)(?=\n|ИНН|БИК|$)',
        # "Банк:" с двоеточием
        r'(?:^|\n)\s*Банк\s*:\s*([^\n]{3,100}?)(?=\n|$)',
        # "Банк Название банка ТОЧКА..." — пропускаем "Название банка"
        r'(?:^|\n|\s)Банк\s+(?:Название\s+банка\s+)?([А-ЯЁ«"ТОЧКА][^\n]{2,100}?)(?=\s+БИК|\s+ИНН|\s+Расч|\s*$)',
        # Просто "Банк" + известные аббревиатуры
        r'(?:^|\n)\s*Банк\s+((?:ПАО|АО|ОАО|ЗАО|НКО|КБ|ТОЧКА)\s*[«"]?[^\n«"]{2,80}[»"]?)(?=\n|БИК|$)',
    ]
    for bp in bank_patterns:
        m = re.search(bp, t, re.I | re.MULTILINE | re.DOTALL)
        if m:
            val = m.group(1).strip().rstrip(',. ')
            # Убираем "Название банка" если вдруг попало в начало
            val = re.sub(r'^(?:Название\s+банка\s+)', '', val, flags=re.I).strip()
            # Обрезаем если БИК попал в конец строки банка
            val = re.split(r'\s+БИК\b', val, flags=re.I)[0].strip().rstrip(',. ')
            if len(val) >= 3 and not re.match(r'^\d+$', val) and not val[:2].isdigit():
                data['bank'] = val[:100]
                break

    # ДИРЕКТОР
    for pat in [
        r'(?:Генеральный\s+директор|Ген\.?\s*дир\.?)\s*[:/]?\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:Генеральный\s+директор|Ген\.?\s*дир\.?)([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:Генеральный\s+директор|Директор)[^\n]*\n\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?)',
        r'(?:Директор|директор)\s*/[^/]*/\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:^|\n)\s*Директор\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
        r'(?:Индивидуальный предприниматель|ИП)\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
    ]:
        m = re.search(pat, t, re.I | re.MULTILINE)
        if m:
            data['director'] = m.group(1).strip()
            break

    # АДРЕС: собираем полный адрес: индекс + город + улица + дом/квартира
    # Стратегия: ищем блок начинающийся с индекса до номера дома/квартиры
    addr_found = False
    # Вариант 1: явная метка "Юридический адрес" или "Почтовый адрес"
    ADDR_STOP = r'\s+(?:ОГРН|ИНН|КПП|БИК|Банк|Счёт|Счет|р/с|к/с|Телефон|Email|Электронная|Директор|Руководитель|Фактический\s+адрес|Почтовый\s+адрес|Юридический\s+адрес)'
    for addr_pat in [
        r'(?:Почтовый\s+адрес|Юридический\s+адрес|Юр\.?\s*адрес)\s*[:\n]?\s*(\d{6}[^\n]{10,200})',
        r'(?:^|\n)\s*(?:Почтовый\s+адрес|Юридический\s+адрес)\D{0,5}(\d{6}[^\n]{10,200})',
    ]:
        m = re.search(addr_pat, t, re.I | re.MULTILINE)
        if m:
            val = re.split(ADDR_STOP, m.group(1), flags=re.I)[0].strip().rstrip(',. ')
            data['address'] = val
            addr_found = True
            break
    # Вариант 2: индекс + многострочный блок (индекс/город/улица/дом на разных строках)
    if not addr_found:
        m = re.search(
            r'(?:^|\n)\s*(?:Индекс|Почтовый\s+индекс)\s*[:\n]?\s*(\d{6})',
            t, re.I | re.MULTILINE
        )
        if m:
            # Берём текст начиная с индекса, собираем строки до пустой строки или нерелевантного поля
            start = m.start(1)
            block = t[start:start+400]
            # Убираем переносы строк — склеиваем в одну строку
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            addr_parts = []
            for line in lines:
                # Стоп-слова: явно не адрес
                if re.match(r'^(?:Телефон|Email|Электронная|ОГРН|ИНН|КПП|БИК|Банк|Счёт|Счет|р/с|к/с|ОКВЭД|Данные|Серия|Руководитель|Директор)', line, re.I):
                    break
                addr_parts.append(line)
                # Если нашли дом/квартиру — стоп
                if re.search(r'(?:д\.?\s*\d|дом\s*\d|кв\.?\s*\d|квартира\s*\d)', line, re.I):
                    break
            if addr_parts:
                data['address'] = ', '.join(addr_parts).strip().rstrip(',. ')
                addr_found = True
    # Вариант 3: одна строка с индексом и городом — обрезаем по стоп-словам
    if not addr_found:
        m = re.search(r'(\d{6},?\s*(?:РОССИЯ|Россия|г\.|город|Г\.?\s|ул\.|пр-кт)[^\n]{10,200})', t, re.I)
        if m:
            val = m.group(1).strip()
            # Обрезаем по стоп-словам которые не относятся к адресу
            val = re.split(ADDR_STOP, val, flags=re.I)[0].strip().rstrip(',. ')
            data['address'] = val

    return data


# ── Клавиатуры ─────────────────────────────────────────────────────────────

def kb_main():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Создать КП",  callback_data="menu_kp")],
        [InlineKeyboardButton("📄 Документы",   callback_data="menu_docs")],
        [InlineKeyboardButton("🧮 Калькулятор", callback_data="menu_calc")],
    ])

def kb_docs():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 Договор + Счёт + Акт", callback_data="menu_all_docs")],
        [InlineKeyboardButton("◀ Назад",                  callback_data="menu_back")],
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
    """
    sel = OrderedDict: {bid: dur} в порядке выбора пользователя.
    Выбранные блоки показываются в порядке выбора со стрелками ↑↓.
    Невыбранные — внизу списка.
    """
    rows = []
    # Сначала выбранные — в порядке выбора
    sel_list = list(sel.keys())
    for n, bid in enumerate(sel_list, 1):
        bname = next(bn for b, bn in PROGRAM_BLOCKS if b == bid)
        is_fixed = bid in FIXED_DUR
        dur = sel[bid]
        label = f"✅ {n}. {bname} — {dur}"
        if not is_fixed and dur_mode == bid:
            rows.append([InlineKeyboardButton(
                ("▶ " if d == dur else "") + d,
                callback_data=f"dur_{bid}_{d}"
            ) for d in DURATIONS])
        row = [InlineKeyboardButton(label, callback_data=f"tog_{bid}")]
        if not is_fixed and dur_mode != bid:
            row.append(InlineKeyboardButton("⏱", callback_data=f"editdur_{bid}"))
        # Стрелки перемещения (только если нет dur_mode)
        if not is_fixed and dur_mode != bid and len(sel_list) > 1:
            up   = InlineKeyboardButton("↑", callback_data=f"up_{bid}")
            down = InlineKeyboardButton("↓", callback_data=f"dn_{bid}")
            if n == 1:
                row.append(down)
            elif n == len(sel_list):
                row.append(up)
            else:
                row.append(up)
                row.append(down)
        rows.append(row)
    # Потом невыбранные
    for bid, bname in PROGRAM_BLOCKS:
        if bid not in sel:
            rows.append([InlineKeyboardButton(f"☐  {bname}", callback_data=f"tog_{bid}")])
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
        [InlineKeyboardButton("✏️ Вставить текст карточки", callback_data="card_text")],
        [InlineKeyboardButton("❌ Отмена",                   callback_data="cancel")],
    ])

def kb_doc_type():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏢 ООО / АО",  callback_data="type_ooo")],
        [InlineKeyboardButton("👤 ИП",         callback_data="type_ip")],
        [InlineKeyboardButton("🧑 Физ. лицо", callback_data="type_fiz")],
        [InlineKeyboardButton("❌ Отмена",     callback_data="cancel")],
    ])

def kb_doc_confirm():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Создать документы",  callback_data="doc_confirm_yes")],
        [InlineKeyboardButton("🔄 Начать заново",      callback_data="doc_confirm_no")],
        [InlineKeyboardButton("❌ Отмена",             callback_data="cancel")],
    ])

def kb_card_review(card):
    """
    Клавиатура анкеты карточки.
    Вся строка кликабельна: нажатие открывает редактирование поля.
    """
    rows = []
    for key, label in CARD_FIELDS:
        if key not in card:
            continue
        val = card.get(key, '')
        icon = '✅' if val else '❌'
        display = (val[:38] + '…') if val and len(val) > 38 else (val or 'не найдено')
        rows.append([
            InlineKeyboardButton(
                f"{icon} {label}: {display}",
                callback_data=f"cardedit_{key}"
            ),
        ])
    rows.append([
        InlineKeyboardButton("✅ Подтвердить", callback_data="card_ok"),
        InlineKeyboardButton("❌ Отмена",      callback_data="cancel"),
    ])
    return InlineKeyboardMarkup(rows)


# ── Генерация файлов ────────────────────────────────────────────────────────

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
    t_re = re.compile(r'<w:t(\s[^>]*)?>(.*?)</w:t>', re.DOTALL)
    result = []
    pos = 0
    search_from = 0
    para_re = re.compile(r'<w:p[ >]')
    end_re   = re.compile(r'</w:p>')
    while True:
        ms = para_re.search(xml, search_from)
        if not ms: break
        me = end_re.search(xml, ms.end())
        if not me: break
        p_start, p_end = ms.start(), me.end()
        para = xml[p_start:p_end]
        # Не объединяем runs если параграф содержит таб-стоп — он разделяет runs намеренно
        if '<w:tab/>' in para:
            search_from = p_end
            continue
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
    full = num_to_words(price_str.replace(' ', ''))
    for sfx in [' рублей 00 копеек', ' рублей', ' 00 копеек']:
        if full.endswith(sfx):
            return full[:-len(sfx)]
    return full

def calc_duration(time_str):
    times = re.findall(r'(\d{1,2}):(\d{2})', time_str)
    if len(times) >= 2:
        start = int(times[0][0]) * 60 + int(times[0][1])
        end   = int(times[1][0]) * 60 + int(times[1][1])
        # Полночь (00:00) и переход через сутки
        if end == 0:
            end = 24 * 60
        elif end < start:
            end += 24 * 60
        diff = end - start
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
    for root, dirs, files in os.walk(work_dir):
        if 'document.xml' in files:
            return os.path.join(root, 'document.xml')
    raise FileNotFoundError(f"document.xml не найден в {work_dir}")

def post_process_docx_xml(xml):
    xml = re.sub(r'<w:highlight[^/]*/>', '', xml)
    xml = re.sub(r'<w:highlight\b[^>]*/>', '', xml)
    return xml

def _apply_replacements(body, pairs):
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
    doc_type   = data.get('doc_type', 'ooo')
    doc_num    = data['doc_num']
    today      = date.today().strftime('%d.%m.%Y')
    date_event = data['date_event']
    time_event = data['time_event']
    duration   = data.get('duration', '—')
    address    = data['address']
    price      = data['price']
    price_words       = num_to_words(price.replace(' ', ''))
    price_short       = num_to_words_short(price)
    today_day         = today.split('.')[0].lstrip('0')
    today_month_year  = _month_year(today)
    date_srок         = f'{_date_word(date_event)} года {time_event}'
    date_schet_word   = _date_word(today)
    price_fmt         = _fmt_price(price)
    zak_polnaya_parts = [p for p in [company_name,
                                      f'ИНН {inn}' if inn else '',
                                      f'КПП {kpp}' if kpp else '',
                                      address_zak] if p]
    zak_polnaya = ', '.join(zak_polnaya_parts) if zak_polnaya_parts else '—'
    tmp_dir = tempfile.mkdtemp()
    results = []
    tmpl_dog = {
        'ooo': TEMPLATE_DOGOVOR,
        'ip':  TEMPLATE_DOGOVOR_IP,
        'fiz': TEMPLATE_DOGOVOR_FIZ,
    }.get(doc_type, TEMPLATE_DOGOVOR)
    fiz_fio      = data.get('fiz_fio', '')
    fiz_passport = data.get('fiz_passport', '')
    fiz_issued   = data.get('fiz_issued', '')
    fiz_code     = data.get('fiz_code', '')
    zak_name = fiz_fio if doc_type == 'fiz' else company_name
    bank_zak = card.get('bank', '')
    try:
        wdir1 = os.path.join(tmp_dir, 'dog_work')
        os.makedirs(wdir1)
        with zipfile.ZipFile(tmpl_dog, 'r') as z:
            z.extractall(wdir1)
        xml_path1 = find_doc_xml(wdir1)
        with open(xml_path1, encoding='utf-8') as f:
            xml1 = f.read()
        xml1 = normalize_docx_xml(xml1)
        xml1 = post_process_docx_xml(xml1)
        dog_pairs = [
            ('[[НОМ]]',       doc_num),
            ('[[ДЕНЬ]]',      today_day),
            ('[[МЕС_ГОД]]',   today_month_year),
            ('[[ДИР]]',       make_genitive_fio(director)),
            ('[[ДИР_ИМ]]',    director),   # именительный падеж (для блока подписей)
            ('[[ДЛИТ]]',      duration),
            ('[[ДАТА_МЕР]]',  date_srок),
            ('[[МЕСТО]]',     address),
            ('[[СУМ_Ц]]',     price),
            ('[[СУМ_СЛ]]',    price_short),
            ('_________________[[ЗАК]]',
             f'_________________{ _initials(_strip_prefix(zak_name))}' if doc_type in ('ip','fiz')
             else f'_________________{ _initials(zak_name)}'),
            ('[[ ЗАК]]',
             _strip_prefix(zak_name) if doc_type in ('ip','fiz') else zak_name),
            ('[[ЗАК]]',       zak_name),
            ('[[ФИО]]',       fiz_fio),
            ('[[СЕР И НОМ]]', fiz_passport),
            ('[[КЕМ И КОГ]]', fiz_issued),
            ('[[КОД ПОД]]',   fiz_code),
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
        xml1 = xml1.replace('[[ЗАК_АДР]]', address_zak or '')
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
            ('_________________[[ЗАК]]',
             f'_________________{ _initials(_strip_prefix(zak_name))}' if doc_type in ('ip','fiz')
             else f'_________________{ _initials(zak_name)}'),
            ('[[ ЗАК]]',
             _strip_prefix(zak_name) if doc_type in ('ip','fiz') else zak_name),
            ('[[ЗАК]]',       zak_name),
            ('[[ФИО]]',       fiz_fio),
            ('[[СЕР И НОМ]]', fiz_passport),
            ('[[КЕМ И КОГ]]', fiz_issued),
            ('[[КОД ПОД]]',   fiz_code),
            ('[[ИНН]]',       inn),
            ('[[КПП]]',       kpp),
            ('[[СУМ_Ц]]',     price_fmt),
            ('[[СУМ_СЛ]]',    price_words),
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
        sch_final = sch_pdf if sch_pdf and os.path.exists(sch_pdf) else sch_path
        if os.path.exists(sch_final):
            results.append((sch_final, f'Счёт_{doc_num}.pdf' if sch_pdf else f'Счёт_{doc_num}.docx'))
    except Exception as e:
        raise Exception(f"Ошибка счёта: {e}")
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
        isp_str = f'{ISPOLNITEL["name"]}, ИНН {ISPOLNITEL["inn"]}'
        zak_str = company_name
        if inn:
            zak_str += f', ИНН {inn}'
        akt_repl = [
            ('[[ЗАК]]',   company_name),
            ('[[ИНН]]',   inn),
            ('N 151 от «10» октября',
             f'N {doc_num} от «{date_event.split(".")[0]}» {_month_only(date_event)}'),
            ('2025г.',    f'{date_event.split(".")[-1]}г.'),
            ('ИП Эрдынеев Гэсэр Буянтуевич, ИНН 032315540193', isp_str),
            ('ЮЛ, ИНН',  isp_str),
            ('Заказчик, ИНН 9725189078', zak_str),
            ('Заказчик, ИНН',            zak_str),
            (f'Заказчик: {ISPOLNITEL["name"]}, ИНН {ISPOLNITEL["inn"]}', f'Заказчик: {zak_str}'),
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
        akt_final = akt_pdf if akt_pdf and os.path.exists(akt_pdf) else akt_docx
        if os.path.exists(akt_final):
            results.append((akt_final, f'Акт_{doc_num}.pdf' if akt_pdf else f'Акт_{doc_num}.docx'))
    except Exception as e:
        raise Exception(f"Ошибка акта: {e}")
    if not results:
        raise Exception("Ни один документ не был создан. Проверьте шаблоны на сервере.")
    if doc_type == 'fiz':
        return [results[0]], tmp_dir
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

def _strip_prefix(name):
    return re.sub(
        r'^(индивидуальный\s+предприниматель|ИП|ООО|АО|ОАО|ЗАО|ПАО)\s+',
        '', name.strip(), flags=re.I
    ).strip()

def _initials(fio):
    parts = fio.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio


# ── КП Handlers ─────────────────────────────────────────────────────────────


# ── Прайс для калькулятора ────────────────────────────────────────────────────
PRICE_STUDIO = {
    ("велком", 1):   1800,
    ("велком", 1.5): 2200,
    ("велком", 2):   2800,
    ("велком", 3):   3200,
    ("лаунж",  1):   1800,
    ("лаунж",  1.5): 2200,
    ("лаунж",  2):   2800,
    ("лаунж",  3):   3200,
}
STUDIO_MIN = {
    ("велком", 1):   18000,
    ("велком", 1.5): 22000,
    ("лаунж",  1):   14400,
    ("лаунж",  1.5): 17600,
}
PRICE_VYEZD = {
    (1,   10): 28000, (1,   15): 33000, (1,   20): 39000,
    (1,   25): 44000, (1,   30): 49000, (1,   35): 54000, (1,   40): 59000,
    (1.5, 10): 32000, (1.5, 15): 37000, (1.5, 20): 43000,
    (1.5, 25): 48000, (1.5, 30): 53000, (1.5, 35): 58000, (1.5, 40): 63000,
    (2,   10): 37000, (2,   15): 42000, (2,   20): 48000,
    (2,   25): 54000, (2,   30): 58000, (2,   35): 63000, (2,   40): 68000,
}

def calc_vyezd(people: int, hours: float) -> int:
    brackets = [10, 15, 20, 25, 30, 35, 40]
    for b in brackets:
        if people <= b:
            return PRICE_VYEZD.get((hours, b), 0)
    return 0

def calc_studio(fmt: str, hours: float, people: int) -> dict:
    price_per = PRICE_STUDIO.get((fmt, hours), 0)
    total = price_per * people
    minimum = STUDIO_MIN.get((fmt, hours), 0)
    actual = max(total, minimum)
    return {"per_person": price_per, "raw": total, "minimum": minimum, "actual": actual}

CALC_FORMAT, CALC_HOURS, CALC_PEOPLE, CALC_RS = range(40, 44)

def kb_calc_format():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚗 Выезд",         callback_data="calc_vyezd")],
        [InlineKeyboardButton("🏢 Студия велком", callback_data="calc_velkom")],
        [InlineKeyboardButton("🏠 Студия лаунж",  callback_data="calc_lanzh")],
        [InlineKeyboardButton("❌ Отмена",         callback_data="cancel")],
    ])

def kb_calc_hours_vyezd():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("1 час",    callback_data="ch_1"),
         InlineKeyboardButton("1.5 часа", callback_data="ch_1.5"),
         InlineKeyboardButton("2 часа",   callback_data="ch_2")],
        [InlineKeyboardButton("❌ Отмена", callback_data="cancel")],
    ])

def kb_calc_hours_studio():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("1 час",    callback_data="ch_1"),
         InlineKeyboardButton("1.5 часа", callback_data="ch_1.5")],
        [InlineKeyboardButton("2 часа",   callback_data="ch_2"),
         InlineKeyboardButton("3 часа",   callback_data="ch_3")],
        [InlineKeyboardButton("❌ Отмена", callback_data="cancel")],
    ])

def kb_calc_rs():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("💳 Нет, обычная оплата",              callback_data="rs_no")],
        [InlineKeyboardButton("🏦 Да, по расчётному счёту (+10%)", callback_data="rs_yes")],
    ])

async def calc_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await safe_delete(ctx.bot, query.message.chat_id, query.message.message_id)
    msg = await ctx.bot.send_message(
        query.message.chat_id,
        "Выберите формат мероприятия:",
        reply_markup=kb_calc_format()
    )
    track(ctx, msg.message_id)
    return CALC_FORMAT

async def calc_got_format(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    fmt_map = {"calc_vyezd": "выезд", "calc_velkom": "велком", "calc_lanzh": "лаунж"}
    fmt = fmt_map.get(query.data, "выезд")
    ctx.user_data["calc_fmt"] = fmt
    kb = kb_calc_hours_vyezd() if fmt == "выезд" else kb_calc_hours_studio()
    await query.edit_message_text("Выберите длительность:", reply_markup=kb)
    return CALC_HOURS

async def calc_got_hours(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    hours = float(query.data.replace("ch_", ""))
    ctx.user_data["calc_hours"] = hours
    await query.edit_message_text("Введите количество человек:")
    return CALC_PEOPLE

async def calc_got_people(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        people = int(update.message.text.strip())
        if people < 1:
            raise ValueError
    except ValueError:
        msg = await update.message.reply_text("Введите число больше 0:")
        track(ctx, msg.message_id, update.message.message_id)
        return CALC_PEOPLE
    ctx.user_data["calc_people"] = people
    track(ctx, update.message.message_id)
    msg = await update.message.reply_text(
        "Оплата по расчётному счёту?",
        reply_markup=kb_calc_rs()
    )
    track(ctx, msg.message_id)
    return CALC_RS

async def calc_got_rs(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    fmt    = ctx.user_data.get("calc_fmt", "выезд")
    hours  = ctx.user_data.get("calc_hours", 1)
    people = ctx.user_data.get("calc_people", 10)
    rs     = query.data == "rs_yes"
    if fmt == "выезд":
        base = calc_vyezd(people, hours)
        if base == 0:
            await query.edit_message_text(
                "Более 40 человек рассчитывается индивидуально.\nСвяжитесь с менеджером.",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("◀ В меню", callback_data="menu_back")]])
            )
            return ConversationHandler.END
        total = int(base * 1.1) if rs else base
        rs_note = " (+10% РС)" if rs else ""
        text = (
            f"🧮 Расчёт стоимости\n\n"
            f"Формат: Выезд\n"
            f"Длительность: {hours} ч\n"
            f"Гостей: {people} чел\n"
            f"{'Оплата по РС' if rs else 'Обычная оплата'}\n\n"
            f"Стоимость{rs_note}: {total:,} ₽".replace(",", " ")
        )
    else:
        r = calc_studio(fmt, hours, people)
        if r["per_person"] == 0:
            await query.edit_message_text(
                f"Для формата {fmt} {hours} ч нет тарифа.\nПроверьте данные.",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("◀ В меню", callback_data="menu_back")]])
            )
            return ConversationHandler.END
        base = r["actual"]
        total = int(base * 1.1) if rs else base
        rs_note = " (+10% РС)" if rs else ""
        min_note = f"\n⚠️ Применена минималка: {r['minimum']:,} ₽".replace(",", " ") if r["raw"] < r["minimum"] else ""
        text = (
            f"🧮 Расчёт стоимости\n\n"
            f"Формат: Студия {fmt}\n"
            f"Длительность: {hours} ч\n"
            f"Гостей: {people} чел\n"
            f"Цена/чел: {r['per_person']:,} ₽\n".replace(",", " ") +
            f"{'Оплата по РС' if rs else 'Обычная оплата'}"
            f"{min_note}\n\n"
            f"Итого{rs_note}: {total:,} ₽".replace(",", " ")
        )
    await query.edit_message_text(
        text,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔄 Новый расчёт", callback_data="menu_calc")],
            [InlineKeyboardButton("◀ В меню",        callback_data="menu_back")],
        ])
    )
    return ConversationHandler.END

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    msg = await update.message.reply_text("Главное меню:", reply_markup=kb_main())
    track(ctx, msg.message_id)

async def menu_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "menu_kp":
        ctx.user_data.clear()
        await query.edit_message_text("Для кого КП? (например: Компании Ромашка)")
        track(ctx, query.message.message_id)
        return KP_NAME
    elif query.data == "menu_docs":
        await query.edit_message_text("Раздел документов:", reply_markup=kb_docs())
        track(ctx, query.message.message_id)
        return DOC_MENU
    elif query.data == "menu_all_docs":
        saved_mid = query.message.message_id
        ctx.user_data.clear()
        await query.edit_message_text("Тип заказчика:", reply_markup=kb_doc_type())
        track(ctx, saved_mid)
        return DOC_TYPE
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
        await q.edit_message_text(
            "Адрес выезда:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("❌ Отмена", callback_data="cancel")
            ]])
        )
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
        if bid in sel:
            del sel[bid]
        else:
            sel[bid] = get_base_dur(bid, fmt)
        if bid not in FIXED_DUR: ctx.user_data['dur_mode'] = None
    elif data.startswith('up_'):
        bid = data[3:]
        keys = list(sel.keys())
        i = keys.index(bid)
        if i > 0:
            keys[i], keys[i-1] = keys[i-1], keys[i]
            sel = {k: sel[k] for k in keys}
            ctx.user_data['selected'] = sel
        ctx.user_data['dur_mode'] = None
    elif data.startswith('dn_'):
        bid = data[3:]
        keys = list(sel.keys())
        i = keys.index(bid)
        if i < len(keys) - 1:
            keys[i], keys[i+1] = keys[i+1], keys[i]
            sel = {k: sel[k] for k in keys}
            ctx.user_data['selected'] = sel
        ctx.user_data['dur_mode'] = None
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
            dur = sel[bid]
            lines.append(f"{n}) {bname} — {dur}")
            n += 1
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
        try:
            await q.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        msg = await q.message.reply_text("Главное меню:", reply_markup=kb_main())
        track(ctx, msg.message_id)
        return ConversationHandler.END
    await q.edit_message_text("Готовлю КП...")
    try:
        d = ctx.user_data
        path, fname, tmp_dir = make_kp_pdf(
            d['location'], d['name'], d['date'], d['time'],
            d['program_lines'], d['price'],
            d.get('address', '')
        )
        with open(path, 'rb') as f:
            await q.message.reply_document(document=f, filename=fname)
        shutil.rmtree(tmp_dir, ignore_errors=True)
        await delete_tracked(ctx, chat_id)
        await safe_delete(ctx.bot, chat_id, q.message.message_id)
    except Exception as e:
        tb = traceback.format_exc()
        print(f"BUILD_KP ERROR: {tb}")
        await q.message.reply_text(f"Ошибка: {e}")
    return ConversationHandler.END


# ── Свободный ввод данных мероприятия ────────────────────────────────────────

FREE_INPUT_PROMPT = (
    "Напиши данные мероприятия одним сообщением.\n"
    "Разделяй каждый пункт символом \\  (обратный слеш)\n\n"
    "Порядок: номер \\ дата \\ время \\ адрес \\ стоимость\n\n"
    "Пример:\n"
    "12 \\ 15.06.2026 \\ 19:00-23:00 \\ Светланский д2 \\ 45000"
)

def _parse_free_input(text):
    """Парсит свободный ввод через запятую: номер, дата, время, адрес, стоимость."""
    result = {}
    t = text.strip()

    # Разбиваем по \ — ожидаем 5 частей
    parts = [p.strip() for p in t.split('\\')]

    if len(parts) >= 5:
        # Формат: номер, дата, время, адрес, стоимость
        result['doc_num']    = parts[0]
        result['date_event'] = parts[1].replace('/', '.')
        result['time_event'] = parts[2]
        # Адрес может содержать запятые — берём всё между временем и последней частью
        result['address']    = ', '.join(parts[3:-1]).strip()
        raw = parts[-1].replace(' ', '')
        try:
            result['price'] = f"{int(''.join(filter(str.isdigit, raw))):,}".replace(',', ' ')
        except:
            result['price'] = raw
    else:
        # Запятых мало — пробуем по паттернам как запасной вариант
        m = re.search(r'(\d{1,2}[./]\d{1,2}[./]\d{4})', t)
        if m: result['date_event'] = m.group(1).replace('/', '.')
        m = re.search(r'(\d{1,2}:\d{2})\s*[-—–]\s*(\d{1,2}:\d{2})', t)
        if m: result['time_event'] = f"{m.group(1)} — {m.group(2)}"
        m = re.search(r'(?:^|№)\s*(\d{1,4})(?=[,\s])', t)
        if m: result['doc_num'] = m.group(1)
        m = re.search(r'(\d{4,6})\s*(?:руб|₽|р\.?)?', t, re.I)
        if m:
            n = m.group(1)
            if not (2020 <= int(n) <= 2030):
                result['price'] = f"{int(n):,}".replace(',', ' ')

    return result

def _missing_fields(d):
    fields = [
        ('doc_num',    'номер договора'),
        ('date_event', 'дата мероприятия (дд.мм.гггг)'),
        ('time_event', 'время (например: 19:00 — 21:00)'),
        ('address',    'адрес проведения'),
        ('price',      'стоимость (рублей)'),
    ]
    return [(k, lbl) for k, lbl in fields if not d.get(k)]

async def doc_free_input(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает свободный ввод данных мероприятия."""
    text = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    parsed = _parse_free_input(text)

    # Обновляем только найденные поля
    for k, v in parsed.items():
        if v:
            ctx.user_data[k] = v

    missing = _missing_fields(ctx.user_data)
    if missing:
        miss_list = ', '.join(lbl for _, lbl in missing)
        msg = await update.message.reply_text(
            f"Почти готово! Не хватает: {miss_list}\n\nДопиши недостающее:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("❌ Отмена", callback_data="cancel")
            ]])
        )
        track(ctx, msg.message_id)
        return DOC_FREE_INPUT

    # Всё заполнено — идём дальше
    time_str = ctx.user_data.get('time_event', '')
    dur = calc_duration(time_str)
    ctx.user_data['duration'] = dur if dur else '—'
    ctx.user_data['today'] = __import__('datetime').date.today().strftime('%d.%m.%Y')

    if ctx.user_data.get('doc_type') == 'fiz':
        msg = await update.message.reply_text("ФИО заказчика (полностью):")
        track(ctx, msg.message_id)
        return DOC_FIZ_FIO
    msg = await update.message.reply_text("Вставь текст карточки предприятия заказчика:")
    track(ctx, msg.message_id)
    return DOC_CARD

# doc_form_cb и doc_form_edit оставлены как заглушки для совместимости
async def doc_form_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    if q.data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        try:
            await q.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        msg = await q.message.reply_text("Главное меню:", reply_markup=kb_main())
        track(ctx, msg.message_id)
        return ConversationHandler.END
    return DOC_FREE_INPUT

async def doc_form_edit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    return await doc_free_input(update, ctx)

# ── Документы Handlers ───────────────────────────────────────────────────────

async def doc_type(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    if q.data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END
    ctx.user_data['doc_type'] = q.data.replace('type_', '')
    ctx.user_data['today'] = __import__('datetime').date.today().strftime('%d.%m.%Y')
    await q.edit_message_text(
        FREE_INPUT_PROMPT,
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("❌ Отмена", callback_data="cancel")
        ]])
    )
    return DOC_FREE_INPUT

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
    ctx.user_data['today'] = __import__('datetime').date.today().strftime('%d.%m.%Y')
    if ctx.user_data.get('doc_type') == 'fiz':
        msg = await update.message.reply_text("ФИО заказчика (полностью):")
        track(ctx, msg.message_id)
        return DOC_FIZ_FIO
    msg = await update.message.reply_text(
        "Вставь текст карточки предприятия заказчика:"
    )
    track(ctx, msg.message_id)
    return DOC_CARD

async def doc_pay_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['today'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text(
        "Вставь текст карточки предприятия заказчика:"
    )
    track(ctx, msg.message_id)
    return DOC_CARD

async def doc_card_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    card = parse_card(text)
    if ctx.user_data.get('doc_type') == 'ip':
        card.pop('kpp', None)
    ctx.user_data['card'] = card
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _show_card_review(update, ctx)

async def doc_card_file(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    tmp = tempfile.mktemp(suffix='.txt')
    await file.download_to_drive(tmp)
    with open(tmp, encoding='utf-8', errors='ignore') as f:
        text = f.read()
    os.remove(tmp)
    card = parse_card(text)
    if ctx.user_data.get('doc_type') == 'ip':
        card.pop('kpp', None)
    ctx.user_data['card'] = card
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _show_card_review(update, ctx)


# ── НОВЫЕ ХЕНДЛЕРЫ: анкета карточки ─────────────────────────────────────────

def _card_review_text(card):
    """Текст анкеты для отображения над кнопками."""
    lines = ["Проверь данные карточки:\n"]
    for key, label in CARD_FIELDS:
        if key not in card:
            continue
        val = card.get(key, '')
        icon = '✅' if val else '❌'
        lines.append(f"{icon} {label}: {val or 'не найдено'}")
    lines.append("\nНажми на поле чтобы исправить, или подтверди.")
    return '\n'.join(lines)

async def _show_card_review(update_or_query, ctx, is_edit=False):
    """
    Показывает анкету карточки.
    is_edit=True: обновляем существующее сообщение (после исправления поля).
    """
    card = ctx.user_data.get('card', {})
    text = _card_review_text(card)
    kb = kb_card_review(card)
    if is_edit:
        q = update_or_query
        try:
            await q.edit_message_text(text, reply_markup=kb)
        except Exception:
            pass
    else:
        update = update_or_query
        msg = await update.message.reply_text(text, reply_markup=kb)
        track(ctx, msg.message_id)
    return DOC_CARD_REVIEW

async def doc_card_review_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатия в анкете карточки."""
    q = update.callback_query; await q.answer()
    data = q.data

    if data == "cancel":
        await delete_tracked(ctx, q.message.chat_id)
        await q.message.reply_text("Отменено.")
        return ConversationHandler.END

    if data == "card_ok":
        # Пользователь подтвердил карточку, идём дальше
        return await _after_card_confirmed(q.message, ctx)

    if data.startswith("cardedit_"):
        # Нажали на поле — запрашиваем новое значение
        field_key = data[len("cardedit_"):]
        ctx.user_data['editing_field'] = field_key
        # Находим метку поля
        label = next((lbl for k, lbl in CARD_FIELDS if k == field_key), field_key)
        current = ctx.user_data.get('card', {}).get(field_key, '')
        hint = f"Текущее: {current}" if current else "Поле пустое"
        try:
            await q.edit_message_text(
                f"Введи значение для поля «{label}»\n{hint}:"
            )
        except Exception:
            pass
        return DOC_CARD_EDIT

    return DOC_CARD_REVIEW

async def doc_card_edit_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Пользователь ввёл новое значение для поля карточки."""
    new_val = update.message.text.strip()
    field_key = ctx.user_data.get('editing_field', '')
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    if field_key:
        ctx.user_data.setdefault('card', {})[field_key] = new_val
        ctx.user_data.pop('editing_field', None)
    # Показываем анкету заново через новое сообщение
    card = ctx.user_data.get('card', {})
    text = _card_review_text(card)
    kb = kb_card_review(card)
    msg = await update.message.reply_text(text, reply_markup=kb)
    track(ctx, msg.message_id)
    return DOC_CARD_REVIEW

async def _after_card_confirmed(message, ctx):
    """После подтверждения карточки: проверяем директора."""
    card = ctx.user_data.get('card', {})
    doc_type = ctx.user_data.get('doc_type', 'ooo')
    if doc_type == 'ip':
        ctx.user_data['director'] = card.get('name', '')
        return await _show_doc_summary(message, ctx, via_message=False)
    if card.get('director'):
        ctx.user_data['director'] = card['director']
        return await _show_doc_summary(message, ctx, via_message=False)
    msg = await message.reply_text("ФИО генерального директора заказчика (полностью):")
    track(ctx, msg.message_id)
    return DOC_DIRECTOR


# ── Остальные хендлеры документов ───────────────────────────────────────────

async def doc_fiz_fio(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['fiz_fio'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Серия и номер паспорта (например: 4520 123456):")
    track(ctx, msg.message_id)
    return DOC_FIZ_PASS

async def doc_fiz_pass(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['fiz_passport'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Кем и когда выдан (например: ОВД Москва 01.01.2020):")
    track(ctx, msg.message_id)
    return DOC_FIZ_ISSUED

async def doc_fiz_issued(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['fiz_issued'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    msg = await update.message.reply_text("Код подразделения (например: 770-001):")
    track(ctx, msg.message_id)
    return DOC_FIZ_CODE

async def doc_fiz_code(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['fiz_code'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _show_fiz_summary(update, ctx)

async def _show_fiz_summary(update, ctx):
    d = ctx.user_data
    summary = (
        f"Проверь данные:\n\n"
        f"Договор №{d['doc_num']}\n"
        f"Дата мероприятия: {d['date_event']}\n"
        f"Время: {d['time_event']}\n"
        f"Длительность: {d.get('duration','—')}\n"
        f"Адрес: {d['address']}\n"
        f"Стоимость: {d['price']} руб\n"
        f"ФИО заказчика: {d.get('fiz_fio','')}\n"
        f"Паспорт: {d.get('fiz_passport','')}\n"
        f"Выдан: {d.get('fiz_issued','')}\n"
        f"Код: {d.get('fiz_code','')}"
    )
    msg = await update.message.reply_text(summary, reply_markup=kb_doc_confirm())
    track(ctx, msg.message_id)
    return DOC_CONFIRM

async def _show_doc_summary(message_or_update, ctx, via_message=True):
    """Итоговая сводка перед созданием документов."""
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
        f"Заказчик: {card.get('name', '?')}\n"
        f"ИНН: {card.get('inn', '?')}\n"
        f"Директор: {ctx.user_data.get('director', '?')}"
    )
    if via_message:
        msg = await message_or_update.message.reply_text(summary, reply_markup=kb_doc_confirm())
        track(ctx, msg.message_id)
    else:
        msg = await message_or_update.reply_text(summary, reply_markup=kb_doc_confirm())
        track(ctx, msg.message_id)
    return DOC_CONFIRM

async def doc_director(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['director'] = update.message.text.strip()
    await safe_delete(ctx.bot, update.message.chat_id, update.message.message_id)
    return await _show_doc_summary(update, ctx, via_message=True)

async def doc_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    chat_id = q.message.chat_id
    if q.data in ('doc_confirm_no', 'cancel'):
        await delete_tracked(ctx, chat_id)
        try:
            await q.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        msg = await q.message.reply_text("Главное меню:", reply_markup=kb_main())
        track(ctx, msg.message_id)
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
        tb = traceback.format_exc()
        print(f"BUILD_DOCS ERROR: {tb}")
        # Показываем последние 3 строки traceback прямо в боте для быстрой диагностики
        tb_short = '\n'.join(tb.strip().split('\n')[-3:])
        await q.message.reply_text(f"Ошибка: {e}\n\n{tb_short}")
    return ConversationHandler.END

async def cancel_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await safe_delete(ctx.bot, update.effective_chat.id, update.message.message_id)
    await delete_tracked(ctx, update.effective_chat.id)
    await update.message.reply_text("Отменено. Напиши /start")
    return ConversationHandler.END


async def global_doc_confirm_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data in ("doc_confirm_no", "cancel"):
        await delete_tracked(ctx, q.message.chat_id)
        try:
            await q.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        msg = await q.message.reply_text("Главное меню:", reply_markup=kb_main())
        track(ctx, msg.message_id)
    elif q.data == "doc_confirm_yes":
        await q.answer("Начни заново через меню.", show_alert=True)

async def post_init(app):
    await app.bot.set_my_commands([
        BotCommand("start",  "Главное меню"),
        BotCommand("kp",     "Создать КП"),
        BotCommand("docs",   "Документы"),
        BotCommand("cancel", "Отменить"),
    ])


def main():
    from telegram.request import HTTPXRequest
    request = HTTPXRequest(
        connection_pool_size=8,
        read_timeout=60,
        write_timeout=60,
        connect_timeout=30,
        pool_timeout=30,
    )
    app = Application.builder().token(TOKEN).request(request).post_init(post_init).build()

    calc_conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(calc_start, pattern=r'^menu_calc$'),
        ],
        states={
            CALC_FORMAT:  [CallbackQueryHandler(calc_got_format, pattern=r'^calc_')],
            CALC_HOURS:   [CallbackQueryHandler(calc_got_hours,  pattern=r'^ch_')],
            CALC_PEOPLE:  [MessageHandler(filters.TEXT & ~filters.COMMAND, calc_got_people)],
            CALC_RS:      [CallbackQueryHandler(calc_got_rs,     pattern=r'^rs_')],
        },
        fallbacks=[
            CallbackQueryHandler(menu_cb, pattern=r'^(cancel|menu_back|menu_calc)$'),
            CommandHandler("start", cmd_start),
        ],
        per_message=False,
    )

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
            DOC_TYPE:        [CallbackQueryHandler(doc_type, pattern=r'^(type_|cancel)')],
            DOC_FREE_INPUT:  [
                MessageHandler(filters.TEXT & ~filters.COMMAND, doc_free_input),
                CallbackQueryHandler(doc_form_cb),
            ],
            DOC_PAY_DATE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_pay_date)],
            DOC_CARD:        [
                MessageHandler(filters.TEXT & ~filters.COMMAND, doc_card_text),
                MessageHandler(filters.Document.ALL, doc_card_file),
            ],
            # НОВЫЕ состояния: анкета карточки
            DOC_CARD_REVIEW: [
                CallbackQueryHandler(doc_card_review_cb),
            ],
            DOC_CARD_EDIT:   [
                MessageHandler(filters.TEXT & ~filters.COMMAND, doc_card_edit_text),
            ],
            DOC_DIRECTOR:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_director)],
            DOC_FIZ_FIO:     [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_fiz_fio)],
            DOC_FIZ_PASS:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_fiz_pass)],
            DOC_FIZ_ISSUED:  [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_fiz_issued)],
            DOC_FIZ_CODE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, doc_fiz_code)],
            DOC_CONFIRM:     [CallbackQueryHandler(doc_confirm, pattern=r'^(doc_confirm_|cancel)')],
        },
        fallbacks=[CommandHandler("cancel", cancel_handler)],
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(calc_conv)
    app.add_handler(kp_conv)
    app.add_handler(doc_conv)
    app.add_handler(CallbackQueryHandler(menu_cb, pattern=r'^menu_back$'))
    app.add_handler(CallbackQueryHandler(global_doc_confirm_cb, pattern=r'^(doc_confirm_no|doc_confirm_yes|cancel)$'))

    print("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()
