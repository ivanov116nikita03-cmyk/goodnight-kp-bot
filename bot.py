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
SLIDE3_REF     = "slide3_ref.pptx"

NAME, LOCATION, ADDRESS, DATE, TIME_, PROGRAM, PRICE, CONFIRM = range(8)

PROGRAM_BLOCKS = [
    ("velkom",  "Велком",              "20 мин",   False),
    ("gn",      "Good Night",          "1 час",    True),
    ("break_",  "Перерыв",             "10 мин",   False),
    ("kk",      "Karaoke Star",        "1 час",    True),
    ("bad",     "Bad Night 21+",       "1 час",    True),
    ("ktokogo", "Кто Кого",            "1.5 часа", True),
    ("arenda",  "Аренда студии",       "1.5 часа", True),
    ("disco",   "Дискотека с диджеем", "1 час",    True),
    ("mafia",   "Мафия",               "1 час",    True),
]
DURATIONS = ["30 мин", "1 час", "1.5 часа", "2 часа"]
ORDER = [b[0] for b in PROGRAM_BLOCKS]


def kb_location():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Большая студия", callback_data="loc_big")],
        [InlineKeyboardButton("Малая студия",   callback_data="loc_small")],
        [InlineKeyboardButton("Выезд",          callback_data="loc_vyezd")],
    ])


def kb_program(selected: dict, dur_mode: str = None):
    rows = []
    n = 1
    for bid, bname, bdefault, has_dur in PROGRAM_BLOCKS:
        is_on = bid in selected
        icon = "[x]" if is_on else "[ ]"
        dur = selected.get(bid, bdefault)
        num = str(n) if is_on else " "
        label = f"{icon} {num}. {bname}" + (f" - {dur}" if is_on else "")
        if is_on:
            n += 1

        if has_dur and is_on and dur_mode == bid:
            dur_row = [InlineKeyboardButton(
                (">" if d == dur else "") + d,
                callback_data=f"dur_{bid}_{d}"
            ) for d in DURATIONS]
            rows.append(dur_row)

        row = [InlineKeyboardButton(label, callback_data=f"tog_{bid}")]
        if has_dur and is_on and dur_mode != bid:
            row.append(InlineKeyboardButton("время", callback_data=f"editdur_{bid}"))
        rows.append(row)

    rows.append([InlineKeyboardButton("Готово", callback_data="prog_done")])
    return InlineKeyboardMarkup(rows)


def fix_textbox13(xml_str, new_address):
    def replacer(m):
        sp = m.group(0)
        new_para = (
            f'<a:p><a:r><a:rPr lang="ru-RU" dirty="0"/>'
            f'<a:t xml:space="preserve">{new_address}</a:t></a:r>'
            f'<a:endParaRPr lang="ru-RU" dirty="0"/></a:p>'
        )
        return re.sub(r'<a:p>.*?</a:p>', new_para, sp, flags=re.DOTALL)
    return re.sub(
        r'<p:sp>(?:(?!<p:sp>).)*?name="TextBox 13".*?</p:sp>',
        replacer, xml_str, flags=re.DOTALL
    )


def build_slide3(ref_s3, date, time_, program_lines, price_str, is_vyezd=False, address=""):
    s3 = ref_s3
    s3 = re.sub(
        r'Дата: [^<|]+\|[^<]+Время: [^<]+',
        f'Дата: {date}  |  Время: {time_}', s3
    )
    old_prog = re.search(r'\d\) .+?(?=</a:t>)', s3, re.DOTALL)
    if old_prog:
        s3 = s3.replace(old_prog.group(0), program_lines)
    if is_vyezd:
        s3 = re.sub(r'[\d\s]+ руб/чел', f'{price_str} руб (общая стоимость)', s3)
    else:
        s3 = re.sub(r'[\d\s]+ руб/чел', f'{price_str} руб/чел', s3)
    addr = address if is_vyezd else 'Денисовский переулок 30, стр. 1'
    s3 = fix_textbox13(s3, addr)
    return s3


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

    with zipfile.ZipFile(SLIDE3_REF, 'r') as z:
        ref_s3 = z.read('ppt/slides/slide1.xml').decode('utf-8')
        ref_s3_rels = z.read('ppt/slides/_rels/slide1.xml.rels').decode('utf-8')

    s1_path = os.path.join(work_dir, 'ppt/slides/slide1.xml')
    with open(s1_path, 'r', encoding='utf-8') as f:
        s1 = f.read()
    s1 = re.sub(r'Программа для [А-Яа-яёЁ]+', f'Программа для {name_gen}', s1)
    if 'Программа для имя' in s1:
        s1 = s1.replace('Программа для имя', f'Программа для {name_gen}')
    with open(s1_path, 'w', encoding='utf-8') as f:
        f.write(s1)

    with zipfile.ZipFile(SLIDE3_REF, 'r') as z:
        ref_s3_data = z.read('ppt/slides/slide1.xml').decode('utf-8')
        ref_s3_rels_data = z.read('ppt/slides/_rels/slide1.xml.rels').decode('utf-8')

    s3_new = build_slide3(ref_s3_data, date, time_, program_lines, price, is_vyezd, address)
    s3_path = os.path.join(work_dir, 'ppt/slides/slide3.xml')
    s3_rels_path = os.path.join(work_dir, 'ppt/slides/_rels/slide3.xml.rels')
    with open(s3_path, 'w', encoding='utf-8') as f:
        f.write(s3_new)
    with open(s3_rels_path, 'w', encoding='utf-8') as f:
        f.write(ref_s3_rels_data)

    with zipfile.ZipFile(dest_pptx, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(work_dir):
            for file in files:
                fp = os.path.join(root, file)
                z.write(fp, os.path.relpath(fp, work_dir))

    return dest_pptx


async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Напиши /kp чтобы создать коммерческое предложение."
    )


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
    else:
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
    ctx.user_data['selected'] = {}
    ctx.user_data['dur_mode'] = None
    await update.message.reply_text(
        "Выбери блоки программы. Кнопка 'время' меняет длительность блока.",
        reply_markup=kb_program(ctx.user_data['selected'])
    )
    return PROGRAM


async def program_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    sel = ctx.user_data.get('selected', {})

    if data.startswith('tog_'):
        bid = data[4:]
        block = next((b for b in PROGRAM_BLOCKS if b[0] == bid), None)
        if bid in sel:
            del sel[bid]
        else:
            sel[bid] = block[2] if block else '1 час'
        ctx.user_data['dur_mode'] = None

    elif data.startswith('editdur_'):
        ctx.user_data['dur_mode'] = data[8:]

    elif data.startswith('dur_'):
        parts = data.split('_', 2)
        bid, chosen = parts[1], parts[2]
        sel[bid] = chosen
        ctx.user_data['dur_mode'] = None

    elif data == 'prog_done':
        if not sel:
            await query.edit_message_text(
                "Выбери хотя бы один блок!",
                reply_markup=kb_program(sel)
            )
            return PROGRAM
        ctx.user_data['selected'] = sel
        await query.edit_message_text(
            "Стоимость?\nДля студии: рублей с человека\nДля выезда: общая сумма"
        )
        return PRICE

    await query.edit_message_reply_markup(
        reply_markup=kb_program(sel, ctx.user_data.get('dur_mode'))
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
    n = 1
    lines = []
    for bid, bname, bdefault, has_dur in PROGRAM_BLOCKS:
        if bid in sel:
            dur = sel[bid]
            lines.append(f"{n}) {bname} - {dur}")
            n += 1
    ctx.user_data['program_lines'] = '\n'.join(lines)

    loc_labels = {'big': 'Большая студия', 'small': 'Малая студия', 'vyezd': 'Выезд'}
    loc = loc_labels[ctx.user_data['location']]
    addr = ctx.user_data.get('address', 'Денисовский переулок 30, стр. 1')
    is_vyezd = ctx.user_data['location'] == 'vyezd'

    summary = (
        f"Проверь данные:\n\n"
        f"Клиент: {ctx.user_data['name']}\n"
        f"Локация: {loc}\n"
        f"Адрес: {addr}\n"
        f"Дата: {ctx.user_data['date']}\n"
        f"Время: {ctx.user_data['time']}\n"
        f"Программа:\n{ctx.user_data['program_lines']}\n"
        f"Стоимость: {price} {'руб (общая)' if is_vyezd else 'руб/чел'}"
    )

    await update.message.reply_text(
        summary,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("Верно — создать КП", callback_data="confirm_yes")],
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
        await query.message.reply_text(f"Ошибка при генерации: {e}")

    return ConversationHandler.END


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отменено. Напиши /kp чтобы начать заново.")
    return ConversationHandler.END


def main():
    app = Application.builder().token(TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("kp", cmd_kp),
        ],
        states={
            NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, got_name)],
            LOCATION: [CallbackQueryHandler(got_location, pattern=r'^loc_')],
            ADDRESS:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_address)],
            DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, got_date)],
            TIME_:    [MessageHandler(filters.TEXT & ~filters.COMMAND, got_time)],
            PROGRAM:  [CallbackQueryHandler(program_cb)],
            PRICE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, got_price)],
            CONFIRM:  [CallbackQueryHandler(confirm_cb, pattern=r'^confirm_')],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(conv)

    print("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()
