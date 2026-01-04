import sys, os
import time
from pathlib import Path
from tkinter import messagebox

from reportlab.graphics.barcode.eanbc import words
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


if sys.platform == "darwin":
    asset_dir = Path.home() / "Personal" / "PythonProject"/ 'assets'
elif sys.platform == "win32":
    asset_dir = Path(os.environ["APPDATA"]) / "MyApp"
else:
    asset_dir = Path.cwd() / 'assets'

font = asset_dir / 'DejaVuSans.ttf'


pdfmetrics.registerFont(TTFont("DejaVu", font))

def draw_service_card(c, w, h, data, y_offset=0):
    def Y(y):
        return y - y_offset

    c.setFont("DejaVu", 10)

    # ─── TOP SECTION ─────────────────────────────
    c.drawString(40, Y(h-50), "Дата на приемане:")
    c.line(150, Y(h-52), 300, Y(h-52))
    c.drawString(155, Y(h-48), data['Дата'])

    c.drawString(350, Y(h-50), "Подпис на клиент:")
    c.rect(470, Y(h-80), 100, 40)

    c.drawString(40, Y(h-90), "Дата на получаване:")
    c.line(170, Y(h-92), 300, Y(h-92))
    c.drawString(175, Y(h-86), time.strftime("%d %b %Y", time.localtime()))

    c.setDash(3, 3)
    c.line(40, Y(h-115), w-40, Y(h-115))
    c.setDash()

    # ─── TITLE ──────────────────────────────────
    c.setFont("DejaVu", 20)
    c.drawString(40, Y(h-160), "СЕРВИЗНА КАРТА")

    c.setFont("DejaVu", 12)
    c.drawString(40, Y(h-195), "№")
    c.line(60, Y(h-197), 140, Y(h-197))
    c.drawString(65, Y(h-190), data['Стикер'])

    # ─── RIGHT DETAILS ──────────────────────────
    c.setFont("DejaVu", 10)

    c.drawString(300, Y(h-160), "Марка:")
    c.line(350, Y(h-162), 550, Y(h-162))
    c.drawString(355, Y(h-158),data['Марка'])

    c.drawString(300, Y(h-190), "Модел:")
    c.line(350, Y(h-192), 550, Y(h-192))
    c.drawString(355, Y(h-188), data['Модел'])

    c.drawString(300, Y(h-220), "IMEI / Serial No:")
    c.line(420, Y(h-222), 550, Y(h-222))
    c.drawString(425, Y(h-218), data['Сериен Номер'] or data['IMEI'])

    c.drawString(300, Y(h-255), "Работи / не работи:")
    c.line(430, Y(h-257), 550, Y(h-257))
    c.drawString(435, Y(h-253), data['Работи'])

    c.drawString(300, Y(h-285), "Външно състояние:")
    c.line(430, Y(h-287), 550, Y(h-287))
    c.drawString(435, Y(h-283), data['Външно състояние'])

    c.drawString(300, Y(h-315), "Описание на дефекта:")
    c.drawString(300, Y(h-345), data['Описание на дефекта'])

    # ─── LOGO ───────────────────────────────────
    c.drawImage(
        asset_dir / "q.png",
        50,
        Y(h-330),
        width=220,
        height=100,
        preserveAspectRatio=True,
        mask="auto"
    )

    c.rect(40, Y(h-340), 230, 110)
    c.drawCentredString(155, Y(h-325), "0883 48 27 26 / 0883 48 27 28")

    # ─── CUT LINE ───────────────────────────────
    c.setDash(8, 0)
    c.line(10, Y(h-365), w-10, Y(h-365))
    c.setDash()


def create_service_card(work_directory, data:dict):
    name = data['Сериен Номер'] or data['IMEI']
    name = work_directory / name
    c = canvas.Canvas(f"{name}.pdf", pagesize=A4)
    w, h = A4

    # First card (top)
    draw_service_card(c, w, h, data, y_offset=0)

    # Second card (bottom)
    draw_service_card(c, w, h, data, y_offset=420)

    c.save()
    messagebox.showinfo("Готово", "PDF успешно создан")
