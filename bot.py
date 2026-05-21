import os
import re
import zipfile
import shutil
import tempfile
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes, ConversationHandler
)

TOKEN = os.environ.get("BOT_TOKEN", "")

TEMPLATE_BIG   = "template_big.pptx"
TEMPLATE_SMALL = "template_small.pptx"
TEMPLATE_VYEZD = "template_vyezd.pptx"

NAME, LOCATION, ADDRESS, DATE, TIME_, FORMAT_TYPE, PROGRAM, PRICE, CONFIRM = range(9)

# Фиксированные блоки — длительность не меняется
FIXED_DUR = {
    "velkom": "20 мин",
    "break_": "10 мин",
}

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

# Базовые длительности по формату
BASE_DUR = {
    "game":   {"default": "1.5 часа", "arenda": "1.5 часа"},
    "packet": {"default": "1 час",    "arenda": "30 мин"},
    "free":   {"default": "1 час",    "arenda": "1 час"},
}


def get_base_dur(bid, fmt):
    if bid in FIXED_DUR:
        return FIXED_DUR[bid]
    base = BASE_DUR.get(fmt, BASE_DUR["free"])
    return base.get(bid, base["default"])


def kb_location():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Большая студия", callback_data="loc_big")],
        [InlineKeyboardButton("Малая студия",   callback_data="loc_small")],
        [InlineKeyboardButton("Выезд",          callback_data="loc_vyezd")],
    ])


def kb_format_type():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Игра (база: 1.5 ч на блок)",  callback_data="fmt_game")],
        [InlineKeyboardButton("Пакет (база: 1 ч на блок)",   callback_data="fmt_packet")],
        [InlineKeyboardButton("Свободный выбор",              callback_data="fmt_free")],
    ])


def kb_program(selected, fmt, dur_mode=None):
    rows = []
    n = 1
    for bid, bname in PROGRAM_BLOCKS:
        is_on = bid in selected
        is_fixed = bid in FIXED_DUR
        icon = "[x]" if is_on else "[ ]"
        num = str(n) if is_on else " "
        dur = selected.get(bid, get_base_dur(bid, fmt)) if is_on else ""
        label = f"{icon} {num}. {bname}" + (f" - {dur}" if dur else "")
        if is_on:
            n += 1

        # Показываем выбор длительности если нажата кнопка "время"
        if not is_fixed and is_on and dur_mode == bid:
            rows.append([InlineKeyboardButton(
                (">" if d == selected.get(bid) else "") + d,
                callback_data=f"dur_{bid}_{d}"
            ) for d in DURATIONS])

        row = [InlineKeyboardButton(label, callback_data=f"tog_{bid}")]
        # Кнопка смены времени для всех включённых нефиксированных блоков
        if not is_fixed and is_on and dur_mode != bid:
            row.append(InlineKeyboardButton("⏱", callback_data=f"editdur_{bid}"))
        rows.append(row)

    rows.append([InlineKeyboardButton("Готово", callback_data="prog_done")])
    return InlineKeyboardMarkup(rows)


def replace_shape_text(xml_str, shape_name, new_text, sz="1600", bold="1"):
    def replacer(m):
        sp = m.group(0)
        para = (
            f'<a:p><a:r>'
            f'<a:rPr lang="ru-RU" sz="{sz}" b="{bold}" dirty="0">'
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


def make_genitive(name):
    n = name.strip()
    low = n.lower()
    if low.endswith('ия'): return n[:-2] + 'ии'
    if low.endswith('ья'): return n[:-2] + 'ьи'
    if low.endswith('я'):  return n[:-1] + 'и'
    if low.endswith('а'):
        return n[:-1] + ('и' if low[-2] in 'гкхжшщч' else 'ы')
    if low.endswith('ь'):  return n[:-1] + 'и'
    return name


def build_pptx(data):
    loc = data['location']
    name = data['name']
    name_gen = make_genitive(name)
    date = data['date']
    time_ = data['time']
    program_lines = data['program_lines']
    price = data['price']
    address = data.get('address', '')
    is_vyezd = (loc == 'vyezd')

    template = {'big': TEMPLATE_BIG, 'small': TEMPLATE_SMALL, 'vyezd': TEMPLATE_VYEZD}[loc]
    loc_label = {'big': 'Большая студия', 'small': 'Малая студия', 'vyezd': 'Выезд'}[loc]

    fname = f"KP_{name}_{date}_{loc_label}.pptx"
    tmp_dir = tempfile.mkdtemp()
    dest_pptx = os.path.join(tmp_dir, fname)
    work_dir = os.path.join(tmp_dir, 'work')
    os.makedirs(work_dir)

    with zipfile.ZipFile(template, 'r') as z:
        z.extractall(work_dir)

    # Слайд 1
    s1_path = os.path.join(work_dir, 'ppt/slides/slide1.xml')
    with open(s1_path, 'r', encoding='utf-8') as f:
        s1 = f.read()
    s1 = re.sub(r'Программа для [А-Яа-яёЁ]+', f'Программа для {name_gen}', s1)
    s1 = s1.replace('Программа для имя', f'Программа для {name_gen}')
    with open(s1_path, 'w', encoding='utf-8') as f:
        f.write(s1)

    # Слайд 3
    s3_path = os.path.join(work_dir, 'ppt/slides/slide3.xml')
    with open(s3_path, 'r', encoding='utf-8') as f:
        s3 = f.read()

    s3 = replace_shape_text(s3, 'TextBox_new_51', f'Дата: {date}  |  Время: {time_}', sz="1600")
    s3 = replace_shape_text(s3, 'TextBox_new_54', program_lines, sz="1600")
    price_label = f'{price} руб (общая стоимость)' if is_vyezd else f'{price} руб/чел'
    s3 = replace_shape_text(s3, 'TextBox_new_55', price_label, sz="1600")
    addr = address if is_vyezd else 'Денисовский переулок 30, стр. 1'
    s3 = replace_shape_text(s3, 'TextBox 13', addr, sz="1600")

    with open(s3_path, 'w', encoding='utf-8') as f:
        f.write(s3)

    with zipfile.ZipFile(dest_pptx, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(work_dir):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, work_dir))

    return dest_pptx


# ─── Handlers ────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Напиши /kp чтобы создать КП.")


async def cmd_kp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("Как зовут клиента?")
    return NAME


async def got_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip()
    ctx.user_data['name'] = raw[0].upper() + raw[1:] if raw else raw
    await update.message.reply_text("Выбери локацию:", reply_markup=kb_location())
    return LOCATION


async def got_location(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    loc = query.data.replace('loc_', '')
    ctx.user_data['location'] = loc
    if loc == 'vyezd':
        await query.edit_message_text("Укажи адрес выезда:")
        return ADDRESS
    await query.edit_message_text("Дата мероприятия? Например: 15.06")
    return DATE


async def got_address(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['address'] = update.message.text.strip()
    await update.message.reply_text("Дата мероприятия? Например: 15.06")
    return DATE


async def got_date(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['date'] = update.message.text.strip()
    await update.message.reply_text("Время начала? Например: 19:00")
    return TIME_


async def got_time(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['time'] = update.message.text.strip()
    loc = ctx.user_data['location']
    if loc == 'vyezd':
        ctx.user_data['fmt'] = 'free'
        ctx.user_data['selected'] = {}
        ctx.user_data['dur_mode'] = None
        await update.message.reply_text(
            "Выбери блоки программы.\n"
            "Кнопка ⏱ меняет длительность блока:",
            reply_markup=kb_program({}, 'free')
        )
        return PROGRAM
    await update.message.reply_text(
        "Выбери формат:",
        reply_markup=kb_format_type()
    )
    return FORMAT_TYPE


async def got_format_type(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    fmt = query.data.replace('fmt_', '')
    ctx.user_data['fmt'] = fmt
    ctx.user_data['selected'] = {}
    ctx.user_data['dur_mode'] = None

    hints = {
        'game':   "Игра — база 1.5 ч на блок.\nКнопка ⏱ меняет длительность:",
        'packet': "Пакет — база 1 ч, аренда 30 мин.\nКнопка ⏱ меняет длительность:",
        'free':   "Свободный выбор.\nКнопка ⏱ меняет длительность:",
    }
    await query.edit_message_text(
        hints[fmt],
        reply_markup=kb_program({}, fmt)
    )
    return PROGRAM


async def program_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    sel = ctx.user_data.get('selected', {})
    fmt = ctx.user_data.get('fmt', 'free')

    if data.startswith('tog_'):
        bid = data[4:]
        if bid in sel:
            del sel[bid]
            if ctx.user_data.get('dur_mode') == bid:
                ctx.user_data['dur_mode'] = None
        else:
            sel[bid] = get_base_dur(bid, fmt)
        if bid not in FIXED_DUR:
            ctx.user_data['dur_mode'] = None

    elif data.startswith('editdur_'):
        ctx.user_data['dur_mode'] = data[8:]

    elif data.startswith('dur_'):
        parts = data.split('_', 2)
        sel[parts[1]] = parts[2]
        ctx.user_data['dur_mode'] = None

    elif data == 'prog_done':
        if not sel:
            await query.edit_message_text(
                "Выбери хотя бы один блок!",
                reply_markup=kb_program(sel, fmt)
            )
            return PROGRAM
        ctx.user_data['selected'] = sel
        await query.edit_message_text(
            "Стоимость?\nДля студии: рублей с человека\nДля выезда: общая сумма"
        )
        return PRICE

    await query.edit_message_reply_markup(
        reply_markup=kb_program(sel, fmt, ctx.user_data.get('dur_mode'))
    )
    return PROGRAM


async def got_price(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().replace(' ', '')
    try:
        price_int = int(''.join(filter(str.isdigit, raw)))
        price = f"{price_int:,}".replace(',', ' ')
    except Exception:
        price = raw
    ctx.user_data['price'] = price

    sel = ctx.user_data['selected']
    fmt = ctx.user_data.get('fmt', 'free')
    n = 1
    lines = []
    for bid, bname in PROGRAM_BLOCKS:
        if bid in sel:
            dur = sel[bid]
            lines.append(f"{n}) {bname} - {dur}")
            n += 1
    ctx.user_data['program_lines'] = '\n'.join(lines)

    loc_labels = {'big': 'Большая студия', 'small': 'Малая студия', 'vyezd': 'Выезд'}
    loc = loc_labels[ctx.user_data['location']]
    addr = ctx.user_data.get('address', 'Денисовский переулок 30, стр. 1')
    is_vyezd = ctx.user_data['location'] == 'vyezd'
    fmt_labels = {'game': 'Игра', 'packet': 'Пакет', 'free': 'Выезд / свободный'}
    fmt_label = fmt_labels.get(fmt, '')

    summary = (
        f"Проверь данные:\n\n"
        f"Клиент: {ctx.user_data['name']}\n"
        f"Локация: {loc}\n"
        f"Формат: {fmt_label}\n"
        f"Адрес: {addr}\n"
        f"Дата: {ctx.user_data['date']}\n"
        f"Время: {ctx.user_data['time']}\n"
        f"Программа:\n{ctx.user_data['program_lines']}\n"
        f"Стоимость: {price} {'руб (общая)' if is_vyezd else 'руб/чел'}"
    )
    await update.message.reply_text(
        summary,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("Верно - создать КП", callback_data="confirm_yes")],
            [InlineKeyboardButton("Начать заново",      callback_data="confirm_no")],
        ])
    )
    return CONFIRM


async def confirm_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == 'confirm_no':
        await query.edit_message_text("Начинаем заново. Напиши /kp")
        return ConversationHandler.END
    await query.edit_message_text("Готовлю КП...")
    try:
        path = build_pptx(ctx.user_data)
        fname = os.path.basename(path)
        with open(path, 'rb') as f:
            await query.message.reply_document(document=f, filename=fname)
        shutil.rmtree(os.path.dirname(path), ignore_errors=True)
    except Exception as e:
        await query.message.reply_text(f"Ошибка: {e}")
    return ConversationHandler.END


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отменено. Напиши /kp заново.")
    return ConversationHandler.END


def main():
    app = Application.builder().token(TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("kp", cmd_kp)],
        states={
            NAME:        [MessageHandler(filters.TEXT & ~filters.COMMAND, got_name)],
            LOCATION:    [CallbackQueryHandler(got_location,    pattern=r'^loc_')],
            ADDRESS:     [MessageHandler(filters.TEXT & ~filters.COMMAND, got_address)],
            DATE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, got_date)],
            TIME_:       [MessageHandler(filters.TEXT & ~filters.COMMAND, got_time)],
            FORMAT_TYPE: [CallbackQueryHandler(got_format_type, pattern=r'^fmt_')],
            PROGRAM:     [CallbackQueryHandler(program_cb)],
            PRICE:       [MessageHandler(filters.TEXT & ~filters.COMMAND, got_price)],
            CONFIRM:     [CallbackQueryHandler(confirm_cb,      pattern=r'^confirm_')],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(conv)
    print("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()
