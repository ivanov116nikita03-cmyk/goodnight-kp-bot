"""
generate_kp.py — модуль генерации PDF для Good Night Show КП
Использует reportlab + фоновые изображения
"""
import os
import tempfile
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Регистрируем шрифт (один раз при импорте)
_FONT_PATH = '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf'
try:
    pdfmetrics.registerFont(TTFont('SerifBold', _FONT_PATH))
except Exception:
    pass

# Базовая папка где лежат изображения и этот файл
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _img(name):
    return os.path.join(BASE_DIR, name)

W, H = 1280, 720


def make_kp_pdf(loc: str, name: str, date: str, time_: str,
                program_lines: str, price: str, address: str = "") -> tuple:
    """
    Генерирует PDF-файл КП и возвращает (путь_к_файлу, имя_файла, tmp_dir).

    loc: 'big' | 'small' | 'vyezd'
    name: имя клиента в родительном падеже (например 'Анны')
    date: дата мероприятия (например '15.06')
    time_: время начала (например '19:00')
    program_lines: строки программы через \n
    price: стоимость строкой (например '2 200')
    address: адрес выезда (только для loc='vyezd')
    """
    safe_date = date.replace('/', '-').replace('.', '-')
    loc_label = {'big': 'Большая студия', 'small': 'Малая студия', 'vyezd': 'Выезд'}[loc]
    # Убираем падежное окончание из имени для названия файла
    fname = f"KP_{name}_{safe_date}_{loc_label}.pdf"

    tmp_dir = tempfile.mkdtemp()
    out_path = os.path.join(tmp_dir, fname)

    c = canvas.Canvas(out_path, pagesize=(W, H))

    # ── СЛАЙД 1: имя клиента ──────────────────────────────────────────────────
    c.drawImage(ImageReader(_img('slide_bg1.png')), 0, 0, W, H)
    c.setFillColorRGB(1, 1, 1)
    c.setFont('SerifBold', 60)
    c.drawCentredString(W/2, H/2 + 75, f'Программа для {name}')
    c.drawCentredString(W/2, H/2 + 5, 'от GOOD NIGHT')
    c.showPage()

    # ── СЛАЙД 2: фото команды ─────────────────────────────────────────────────
    c.drawImage(ImageReader(_img('slide_bg2.png')), 0, 0, W, H)
    c.showPage()

    # ── СЛАЙД 3: программа и детали ───────────────────────────────────────────
    c.drawImage(ImageReader(_img('slide_bg3.png')), 0, 0, W, H)
    c.setFillColorRGB(1, 1, 1)
    c.setFont('SerifBold', 26)

    # Программа: x=550, y=470, шаг 36px вниз
    y_prog = 470
    for line in program_lines.split('\n'):
        c.drawString(550, y_prog, line)
        y_prog -= 36

    # Стоимость: x=550, y=290
    price_label = f'{price} руб (общая стоимость)' if loc == 'vyezd' else f'{price} руб/чел'
    c.drawString(550, 290, price_label)

    # Дата и время: x=70, y=57
    c.setFont('SerifBold', 22)
    c.drawString(70, 57, f'Дата: {date}   Время: {time_}')

    # Адрес: правый край - 60, y=57
    addr = address if loc == 'vyezd' else 'Денисовский переулок 30, стр. 1'
    addr_w = c.stringWidth(addr, 'SerifBold', 22)
    c.drawString(W - addr_w - 60, 57, addr)
    c.showPage()

    # ── СЛАЙД 4: фото студии (нет у выезда) ──────────────────────────────────
    if loc == 'big':
        c.drawImage(ImageReader(_img('slide_bg4_big.png')), 0, 0, W, H)
        c.showPage()
    elif loc == 'small':
        c.drawImage(ImageReader(_img('slide_bg4_small.png')), 0, 0, W, H)
        c.showPage()

    # ── СЛАЙД 5: партнёры ─────────────────────────────────────────────────────
    if loc in ('big', 'small'):
        c.drawImage(ImageReader(_img('slide_bg5_studio.png')), 0, 0, W, H)
    else:
        c.drawImage(ImageReader(_img('slide_bg5_vyezd.png')), 0, 0, W, H)
    c.showPage()

    # ── СЛАЙД 6: что ещё есть ─────────────────────────────────────────────────
    c.drawImage(ImageReader(_img('slide_bg6.png')), 0, 0, W, H)
    c.showPage()

    c.save()
    return out_path, fname, tmp_dir
