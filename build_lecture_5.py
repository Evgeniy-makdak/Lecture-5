#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Лекция 5: Геометрическая оптика — как линзы собирают луч в пятно.
Источник: «Лекция_5_Геометрическая оптика_как линзы собирают луч..docx»
"""

import os
import re
from datetime import datetime

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, KeepTogether
from reportlab.lib.colors import darkgray, gray
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PIL import Image, ImageDraw, ImageFont


BASE = os.path.dirname(os.path.abspath(__file__))
PPTX_PATH = os.path.join(BASE, "Лекция_5_Геометрическая_оптика.pptx")
PDF_PATH = os.path.join(BASE, "Лекция_5_Раскадровка_для_лектора.pdf")
DOCX_PATH = os.path.join(BASE, "Лекция_5_Геометрическая оптика.docx")
SOURCE_DOCX = os.path.join(BASE, "Лекция_5_Геометрическая оптика_как линзы собирают луч..docx")


def _detect_font():
    for p in [
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf"),
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "segoeui.ttf"),
    ]:
        if os.path.exists(p):
            bold = p.replace("arial.ttf", "arialbd.ttf").replace("segoeui.ttf", "segoeuib.ttf")
            return p, bold if os.path.exists(bold) else p
    return None, None


def _detect_math_font():
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for p in [
        os.path.join(windir, "Fonts", "timesi.ttf"),
        os.path.join(windir, "Fonts", "cambriai.ttf"),
        os.path.join(windir, "Fonts", "times.ttf"),
    ]:
        if os.path.exists(p):
            return p
    return None


FONT_PATH, FONT_BOLD_PATH = _detect_font()
MATH_FONT_PATH = _detect_math_font()
if FONT_PATH:
    pdfmetrics.registerFont(TTFont("CyrFont", FONT_PATH))
    pdfmetrics.registerFont(TTFont("CyrFontBold", FONT_BOLD_PATH or FONT_PATH))
    PDF_FONT, PDF_FONT_BOLD = "CyrFont", "CyrFontBold"
else:
    PDF_FONT, PDF_FONT_BOLD = "Helvetica", "Helvetica-Bold"

if MATH_FONT_PATH:
    pdfmetrics.registerFont(TTFont("MathFont", MATH_FONT_PATH))
    PDF_MATH = "MathFont"
else:
    PDF_MATH = PDF_FONT


def _pil_font(size=14, bold=False):
    cands = []
    if bold:
        cands += [
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arialbd.ttf"),
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "segoeuib.ttf"),
        ]
    cands += [FONT_PATH, os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf")]
    for p in cands:
        if p and os.path.exists(p):
            try:
                return ImageFont.truetype(p, size=size)
            except OSError:
                pass
    return ImageFont.load_default()


def _tw(d, text, font):
    text = _pil_safe(text)
    try:
        return int(d.textlength(text, font=font))
    except Exception:
        return len(text) * 7


def _draw_text(d, xy, text, fill, font):
    """Единая точка вывода текста на PIL-схемах (безопасная кодировка)."""
    d.text(xy, _pil_safe(text), fill=fill, font=font)


def _draw_sup(d, x, y, base, sup, font, color="#333"):
    fs = _pil_font(max(6, getattr(font, "size", 12) - 5))
    d.text((x, y), base, fill=color, font=font)
    ox = x + _tw(d, base, font)
    d.text((ox, y - 4), sup, fill=color, font=fs)
    return ox + _tw(d, sup, fs)


def _pil_greek_font(size=12):
    windir = os.environ.get("WINDIR", r"C:\Windows")
    for name in ("seguisym.ttf", "segoeui.ttf", "arial.ttf"):
        path = os.path.join(windir, "Fonts", name)
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except OSError:
                continue
    return _pil_font(size)


def _draw_greek(d, x, y, ch, font, color):
    """λ и др. — Segoe Symbol; ν рисуем отдельно (иначе путается с v)."""
    if ch in ("\u03bd", "ν"):
        return _draw_greek_nu(d, x, y, int(getattr(font, "size", 12) or 12), color)
    g = _pil_greek_font(int(getattr(font, "size", 12) or 12))
    d.text((x, y), ch, fill=color, font=g)
    return x + _tw(d, ch, g)


def _draw_greek_nu(d, x, y, size, color):
    """Греческая ν через supersample Segoe UI Symbol."""
    s = max(12, int(size) + 2)
    scale = 6
    path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "seguisym.ttf")
    if not os.path.exists(path):
        path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf")
    big = ImageFont.truetype(path, s * scale)
    tmp = Image.new("RGBA", (s * scale * 2, s * scale * 2), (0, 0, 0, 0))
    td = ImageDraw.Draw(tmp)
    if isinstance(color, str) and color.startswith("#"):
        hx = color[1:]
        if len(hx) == 3:
            hx = "".join(c * 2 for c in hx)
        fill = tuple(int(hx[i:i + 2], 16) for i in (0, 2, 4)) + (255,)
    else:
        fill = (0, 0, 0, 255)
    td.text((scale, scale), "\U0001D708", fill=fill, font=big)
    bbox = tmp.getbbox()
    if not bbox:
        return x + s
    glyph = tmp.crop(bbox)
    tw = max(8, glyph.width // scale)
    th = max(8, glyph.height // scale)
    glyph = glyph.resize((tw, th), Image.Resampling.LANCZOS)
    host = getattr(d, "_image", None)
    if host is not None:
        host.paste(glyph, (int(x), int(y)), glyph)
    else:
        d.bitmap((x, y), glyph.convert("L"), fill=color)
    return x + tw + 1


def _to_unicode_sup(text: str) -> str:
    mp = str.maketrans("0123456789+-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻")
    return text.translate(mp)


def pretty_math(text):
    """PPTX: греческие буквы и верхние индексы (Unicode)."""
    if not text:
        return text
    t = text
    t = re.sub(r"\blambda\b", "λ", t, flags=re.I)
    t = re.sub(r"\bpi\b", "π", t, flags=re.I)
    t = re.sub(r"\bDelta\s+n\b", "Δn", t)
    t = re.sub(r"\bDelta\b", "Δ", t)
    t = re.sub(r"\btheta(\d)\b", lambda m: "θ" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\btheta\b", "θ", t, flags=re.I)
    t = re.sub(r"см2\b", "см²", t)
    t = re.sub(r"мм2\b", "мм²", t)
    t = re.sub(r"/см2\b", "/см²", t)
    t = re.sub(r"/мм2\b", "/мм²", t)
    t = re.sub(r"(\d+(?:,\d+)?)·10\^(-?\d+)", lambda m: m.group(1) + "·10" + _to_unicode_sup(m.group(2)), t)
    t = re.sub(r"10\^(-?\d+)", lambda m: "10" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\^(-?\d+)", lambda m: _to_unicode_sup(m.group(1)), t)
    t = t.replace("NA^2", "NA²").replace("NA2", "NA²")
    t = re.sub(r"\bM2\b", "M²", t)
    return t


def _pil_safe(text):
    """Текст для PIL-схем: сохраняем верхние индексы и греческие буквы."""
    if not text:
        return text
    t = str(text)
    t = re.sub(r"\blambda\b", "λ", t, flags=re.I)
    t = re.sub(r"\bpi\b", "π", t, flags=re.I)
    t = re.sub(r"\btheta(\d)\b", lambda m: "θ" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\btheta\b", "θ", t, flags=re.I)
    t = re.sub(r"см2\b", "см²", t)
    t = re.sub(r"/см2\b", "/см²", t)
    t = re.sub(r"10\^(-?\d+)", lambda m: "10" + _to_unicode_sup(m.group(1)), t)
    t = re.sub(r"\^(-?\d+)", lambda m: _to_unicode_sup(m.group(1)), t)
    t = t.replace("NA^2", "NA²")
    return t


def pdf_sub(text):
    """Кириллица + индексы/степени для ReportLab; греческие — MathFont."""
    if not text:
        return text

    def math_ch(ch):
        return f'<font name="{PDF_MATH}"><i>{ch}</i></font>'

    t = text
    t = t.replace("CO₂", "CO<sub>2</sub>").replace("CO2", "CO<sub>2</sub>")
    t = re.sub(r"\bM2\b", "M²", t)
    t = t.replace("M²", "M<sup>2</sup>").replace("M^2", "M<sup>2</sup>")
    t = t.replace("−", "-")
    t = t.replace("·", "&middot;")
    t = re.sub(r"\blambda\b", "λ", t, flags=re.I)
    t = re.sub(r"\bpi\b", "π", t, flags=re.I)
    t = re.sub(r"\bDelta\s+n\b", "Δn", t)
    t = re.sub(r"\bDelta\b", "Δ", t)
    t = re.sub(r"\btheta(\d)\b", r"θ<sub>\1</sub>", t)
    t = re.sub(r"\btheta\b", "θ", t, flags=re.I)
    t = re.sub(r"см2\b", "см<sup>2</sup>", t)
    t = re.sub(r"мм2\b", "мм<sup>2</sup>", t)
    t = re.sub(r"/см2\b", "/см<sup>2</sup>", t)
    t = re.sub(r"/мм2\b", "/мм<sup>2</sup>", t)
    t = re.sub(r"NA\^2", "NA<sup>2</sup>", t)
    t = re.sub(r"NA2\b", "NA<sup>2</sup>", t)
    t = re.sub(r"\(([^()]+)\)\^(-?\d+)", r"(\1)<sup>\2</sup>", t)
    t = re.sub(r"(\d),(\d+)\^(-?\d+)", r"\1,\2<sup>\3</sup>", t)
    t = re.sub(r"10\^(-?\d+)", r"10<sup>\1</sup>", t)
    t = re.sub(r"10·10\^(-?\d+)", r"10&middot;10<sup>\1</sup>", t)
    t = re.sub(r"2·10\^(-?\d+)", r"2&middot;10<sup>\1</sup>", t)
    t = re.sub(r"(\d+(?:,\d+)?)·10\^(-?\d+)", r"\1&middot;10<sup>\2</sup>", t)
    t = re.sub(r"\^(-?\d+)", r"<sup>\1</sup>", t)
    uni_sup = str.maketrans({
        "⁰": "<sup>0</sup>", "¹": "<sup>1</sup>", "²": "<sup>2</sup>", "³": "<sup>3</sup>",
        "⁴": "<sup>4</sup>", "⁵": "<sup>5</sup>", "⁶": "<sup>6</sup>", "⁷": "<sup>7</sup>",
        "⁸": "<sup>8</sup>", "⁹": "<sup>9</sup>", "⁺": "<sup>+</sup>", "⁻": "<sup>-</sup>",
    })
    t = t.translate(uni_sup)
    for ch in ("ν", "ω", "π", "λ", "Λ", "ħ", "Δ", "θ"):
        t = t.replace(ch, math_ch(ch))
    return t


def save_pil_image(draw_fn, path):
    img = draw_fn()
    img.save(path, "PNG")
    return path


# pasted into build_lecture_5.py — diagrams through main()

# ── Диаграммы ─────────────────────────────────────────────────────────────

def draw_intensity_comparison_pil():
    """Сценарии 1 и 2: I = P/S при d = 1 мм и d = 50 мкм."""
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(12), _pil_font(10)

    _draw_text(d, (12, 8), "Плотность мощности I = P / S (P = 200 Вт)", fill="#333", font=fh)

    # Сценарий 1 — большое пятно
    cx1, cy1, r1 = 150, 210, 55
    d.ellipse([(cx1 - r1, cy1 - r1), (cx1 + r1, cy1 + r1)], outline="#1565c0", width=2)
    _draw_text(d, (70, 55), "Без фокусировки", fill="#1565c0", font=f)
    _draw_text(d, (70, 80), "d = 1 мм", fill="#333", font=fs)
    _draw_text(d, (70, 100), "S ~ 0,008 см²", fill="#666", font=fs)
    _draw_text(d, (70, 125), "I ~ 25 кВт/см²", fill="#e64a19", font=f)

    # Сценарий 2 — малое пятно (масштаб увеличен)
    cx2, cy2, r2 = 430, 210, 18
    d.ellipse([(cx2 - r2, cy2 - r2), (cx2 + r2, cy2 + r2)], outline="#e64a19", width=3)
    _draw_text(d, (340, 55), "С фокусировкой", fill="#e64a19", font=f)
    _draw_text(d, (340, 80), "d = 50 мкм", fill="#333", font=fs)
    _draw_text(d, (340, 100), "S ~ 2·10⁻⁵ см²", fill="#666", font=fs)
    _draw_text(d, (340, 125), "I ~ 10 МВт/см²", fill="#c62828", font=f)

    # Стрелка ×400
    d.line([(260, 210), (310, 210)], fill="#333", width=2)
    d.polygon([(310, 205), (325, 210), (310, 215)], fill="#333")
    _draw_text(d, (248, 175), "x 400", fill="#c62828", font=fh)

    _draw_text(d, (12, 340), "Та же мощность — другая площадь: линза сжимает энергию в пространстве", fill="#555", font=fs)
    return img


def draw_snell_lens_pil():
    """Закон Снеллиуса и сборка лучей в фокус."""
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(11), _pil_font(9)

    _draw_text(d, (12, 8), "Преломление и фокусировка линзы", fill="#333", font=fh)

    # Линза (выпуклая)
    d.arc([(240, 80), (380, 320)], 270, 90, fill="#009688", width=3)
    d.arc([(240, 80), (380, 320)], 90, 270, fill="#009688", width=3)

    # Ось
    d.line([(50, 200), (570, 200)], fill="#ccc", width=1)
    _draw_text(d, (300, 330), "оптическая ось", fill="#999", font=fs)

    # Лучи: параллельные -> сходятся
    for y_off, col in [(160, "#1565c0"), (200, "#e64a19"), (240, "#1565c0")]:
        d.line([(50, y_off), (290, y_off)], fill=col, width=2)
        d.line([(290, y_off), (350, 200)], fill=col, width=2)
        d.line([(350, 200), (520, y_off + (200 - y_off) // 2)], fill=col, width=2)

    d.ellipse([(345, 195), (355, 205)], fill="#c62828")
    _draw_text(d, (358, 188), "фокус", fill="#c62828", font=fs)

    x = 12
    _draw_text(d, (x, 355), "n1 sin ", fill="#333", font=fs)
    x += _tw(d, "n1 sin ", fs)
    x = _draw_greek(d, x, 355, "θ", fs, "#333")
    x = _draw_sup(d, x, 355, "", "1", fs, "#333")
    _draw_text(d, (x, 355), " = n2 sin ", fill="#333", font=fs)
    x += _tw(d, " = n2 sin ", fs)
    x = _draw_greek(d, x, 355, "θ", fs, "#333")
    x = _draw_sup(d, x, 355, "", "2", fs, "#333")
    _draw_text(d, (x + 8, 355), "  |  n = c / v", fill="#666", font=fs)
    return img


def draw_spot_size_pil():
    """d ~ λ / NA."""
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(11), _pil_font(9)

    x = 12
    _draw_text(d, (x, 8), "Дифракционный предел: d ~ ", fill="#333", font=fh)
    x += _tw(d, "Дифракционный предел: d ~ ", fh)
    x = _draw_greek(d, x, 8, "\u03bb", fh, "#333")
    _draw_text(d, (x, 8), " / NA", fill="#333", font=fh)

    # Конус сходящихся лучей
    apex = (310, 280)
    d.polygon([(180, 80), (440, 80), apex[0], apex[1]], outline="#1565c0", fill="#e3f2fd")
    d.line([(180, 80), (440, 80)], fill="#1565c0", width=2)

    # Пятно в фокусе
    d.ellipse([(295, 268), (325, 288)], fill="#e64a19", outline="#c62828", width=2)
    _draw_text(d, (330, 265), "d", fill="#c62828", font=f)

    x = 12
    _draw_text(d, (x, 310), "d ~ ", fill="#1565c0", font=f)
    x += _tw(d, "d ~ ", f)
    x = _draw_greek(d, x, 310, "\u03bb", f, "#1565c0")
    _draw_text(d, (x, 310), " / NA", fill="#1565c0", font=f)

    _draw_text(d, (12, 340), "NA = 0,1 -> d ~ 10,7 мкм  |  NA = 0,4 -> d ~ 2,7 мкм", fill="#333", font=fs)
    _draw_text(d, (12, 360), "Больше NA -> меньше пятно (круче конус)", fill="#666", font=fs)
    return img


def draw_dof_chart_pil():
    """d и DOF для трёх NA."""
    img = Image.new("RGB", (620, 420), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs, fxs = _pil_font(13, True), _pil_font(10), _pil_font(9), _pil_font(8)

    _draw_text(d, (12, 6), "Компромисс NA: пятно d и глубина резкости DOF", fill="#333", font=fh)
    x = 12
    _draw_text(d, (x, 28), "DOF ~ ", fill="#666", font=fs)
    x += _tw(d, "DOF ~ ", fs)
    x = _draw_greek(d, x, 28, "\u03bb", fs, "#666")
    _draw_text(d, (x, 28), " / NA", fill="#666", font=fs)
    x += _tw(d, " / NA", fs)
    x = _draw_sup(d, x, 28, "", "2", fs, "#666")
    _draw_text(d, (x, 28), "  (падает быстрее, чем d)", fill="#666", font=fs)

    cases = [
        ("NA=0,1  пром.", 10.7, 107, "#009688"),
        ("NA=0,4  лаб.", 2.7, 6.7, "#e64a19"),
        ("NA=0,5  экстр.", 2.1, 4.3, "#c62828"),
    ]
    x0, y0, h = 70, 58, 230
    max_v = 120
    bar_w = 38
    gap = 95

    for i, (label, d_val, dof_val, col) in enumerate(cases):
        bx = x0 + i * gap
        _draw_text(d, (bx - 10, y0 - 20), label, fill=col, font=f)
        h_dof = int(dof_val / max_v * h)
        d.rectangle([(bx, y0 + h - h_dof), (bx + bar_w, y0 + h)], fill=col)
        _draw_text(d, (bx - 8, y0 + h + 6), f"DOF {dof_val} мкм", fill=col, font=fs)
        h_d = max(8, int(d_val / max_v * h))
        d.rectangle([(bx + bar_w + 10, y0 + h - h_d), (bx + bar_w + 24, y0 + h)], fill="#333")
        _draw_text(d, (bx + bar_w + 28, y0 + h - h_d - 2), f"d={d_val}", fill="#333", font=fs)

    _draw_text(d, (12, 395), "Малый NA: терпим к рельефу порошка", fill="#555", font=fs)
    _draw_text(d, (320, 395), "NA=0,5: DOF << неровности слоя", fill="#555", font=fs)
    return img


def draw_m2_comparison_pil():
    """d_реал = M2 * d_идеал для волоконного и диодного лазера."""
    img = Image.new("RGB", (620, 400), "white")
    d = ImageDraw.Draw(img)
    fh, f, fs = _pil_font(14, True), _pil_font(11), _pil_font(9)

    _draw_text(d, (12, 8), "Качество пучка M²: d_реал = M² · d_идеал", fill="#333", font=fh)

    # Волоконный
    cx1, cy1, r1 = 160, 200, 13
    d.ellipse([(cx1 - r1, cy1 - r1), (cx1 + r1, cy1 + r1)], fill="#009688", outline="#00695c", width=2)
    _draw_text(d, (80, 60), "Волоконный лазер", fill="#009688", font=f)
    _draw_text(d, (80, 85), "M² = 1,3; d_ид = 10 мкм", fill="#333", font=fs)
    _draw_text(d, (80, 110), "d_реал ~ 13 мкм", fill="#009688", font=f)

    # Диодный
    cx2, cy2, r2 = 430, 200, 80
    d.ellipse([(cx2 - r2, cy2 - r2), (cx2 + r2, cy2 + r2)], outline="#e64a19", width=2, fill="#ffebee")
    d.ellipse([(cx2 - 13, cy2 - 13), (cx2 + 13, cy2 + 13)], fill="#009688", outline="#00695c", width=1)
    _draw_text(d, (350, 60), "Диодный лазер", fill="#e64a19", font=f)
    _draw_text(d, (350, 85), "M² = 20; d_ид = 10 мкм", fill="#333", font=fs)
    _draw_text(d, (350, 110), "d_реал ~ 200 мкм", fill="#e64a19", font=f)

    _draw_text(d, (12, 320), "I ~ 1/d²: плохой пучок → пятно ×20 → I в 400 раз меньше", fill="#c62828", font=fs)
    _draw_text(d, (12, 345), "В СЛП — волоконные лазеры с M² близким к 1", fill="#555", font=fs)
    return img


NA_TABLE = {
    "headers": ["Параметр", "A: промышл.", "B: лаб.", "C: экстр."],
    "rows": [
        ["Числ. апертура (NA)", "0,1", "0,4", "0,5"],
        ["Пятно d, мкм", "~10,7", "~2,7", "~2,1"],
        ["DOF, мкм", "~107", "~6,7", "~4,3"],
        ["СЛП", "Да, стандарт", "Только лаб.", "Нереализуемо"],
        ["Почему?", "DOF > рельефа", "DOF < рельефа", "DOF << рельефа"],
    ],
}

# ── Раскадровка ───────────────────────────────────────────────────────────

READING_WPM = 130

SECTION_STAGE = {
    "1. ТИТУЛЬНЫЙ": (0.15, "Пауза на вход: слайд титула."),
    "2. ПЛАН ЛЕКЦИИ": (0.15, "Показать пять пунктов плана на слайде."),
    "3. ЗАЧЕМ НАМ ОПТИКА? (С РАСЧЁТАМИ НА ДОСКЕ)": (
        4.5,
        "Доска: I = P/S; сценарий 1 (d = 1 мм) и сценарий 2 (d = 50 мкм); сравнение 25 кВт/см2 и 10 МВт/см2.",
    ),
    "4. ЗАКОН СНЕЛЛИУСА — КАК РАБОТАЕТ ЛИНЗА": (
        3.0,
        "Доска: n1 sin theta1 = n2 sin theta2; расшифровка n, theta; схема фокуса.",
    ),
    "5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)": (
        5.0,
        "Доска: d ~ lambda/NA; примеры NA = 0,1 и NA = 0,4; NA = n sin theta.",
    ),
    "6. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2) И ТРИ СЛУЧАЯ": (
        12.0,
        "Доска: DOF ~ lambda/NA2; таблица трёх случаев; три причины, почему NA = 0,5 не для СЛП.",
    ),
    "7. КАЧЕСТВО ПУЧКА M² И ПОДВЕДЕНИЕ ИТОГОВ": (
        4.0,
        "Доска: d_реал = M^2 · d_идеал; итоговые тезисы.",
    ),
}


def _count_words(text):
    return len(re.findall(r"[A-Za-zА-Яа-яЁё0-9]+(?:[-'][A-Za-zА-Яа-яЁё0-9]+)*", text))


def _fmt_clock(minutes):
    sec = max(0, int(round(minutes * 60)))
    return f"{sec // 60:02d}:{sec % 60:02d}"


def _fmt_duration(minutes):
    sec = max(0, int(round(minutes * 60)))
    if sec < 60:
        return f"{sec} сек"
    half = round(minutes * 2) / 2
    if abs(half - minutes) > 0.05:
        return f"{minutes:.1f}".replace(".", ",") + " мин"
    if half == int(half):
        return f"{int(half)} мин"
    return f"{str(half).replace('.', ',')} мин"


def _time_script(blocks):
    out = []
    cum = 0.0
    total_words = 0
    total_read = 0.0
    total_pause = 0.0
    for title, timing, paragraphs in blocks:
        words = sum(_count_words(p) for p in paragraphs)
        total_words += words
        pause_min, cue = SECTION_STAGE.get(title, (0.0, ""))
        if timing != "auto":
            m = re.search(r"(\d\d:\d\d)\s*[–-]\s*(\d\d:\d\d)", timing)
            wall_min = None
            if m:
                def _to_min(x):
                    mm, ss = x.split(":")
                    return int(mm) + int(ss) / 60.0
                wall_min = max(0.0, float(_to_min(m.group(2)) - _to_min(m.group(1))))
            if wall_min is None:
                wall_min = words / READING_WPM + float(pause_min)
            read_min = words / READING_WPM
            total_read += float(read_min)
            total_pause += max(0.0, float(wall_min) - float(read_min))
            out.append((title, timing, paragraphs, cue))
            cum += float(wall_min)
        else:
            read_min = words / READING_WPM
            wall_min = read_min + pause_min
            total_read += read_min
            total_pause += pause_min
            start, end = cum, cum + wall_min
            label = (
                f"{_fmt_clock(start)} – {_fmt_clock(end)} "
                f"(стенка {_fmt_duration(wall_min)} = чтение {_fmt_duration(read_min)}"
                f" + доска/паузы {_fmt_duration(pause_min)})"
            )
            out.append((title, label, paragraphs, cue))
            cum = end
    return out, total_words, cum, total_read, total_pause


# Заключительный абзац — после итоговой схемы M² (PDF/DOCX), не на слайде
EPILOGUE = (
    "Связь со следующей лекцией: мы говорили о линзах, как будто свет — идеальные прямые лучи. "
    "Но в d (диаметр) ≈ lambda (длина волны) / NA (числовая апертура) спрятана волновая природа — дифракция. "
    "На следующем занятии («Волновая оптика») разберём, почему пятно не может быть бесконечно малым "
    "и что такое интерференция."
)

SCRIPT = [
    ("1. ТИТУЛЬНЫЙ", "auto", [
        "Добрый день. Мы продолжаем курс. В Главе 2 мы выяснили главное: сталь плавится не от мощности как таковой, а от плотности мощности — сколько ватт пришлось на каждый квадратный сантиметр пятна.",
        "Сегодня мы разберём инструмент, которым мы управляем этой плотностью: оптическую систему.",
        "Мы ответим на три вопроса: как линза собирает свет в пятно; почему пятно не может быть бесконечно малым; что такое глубина резкости (DOF) и почему она критична для 3D-печати.",
    ]),
    ("2. ПЛАН ЛЕКЦИИ", "auto", [
        "План будет таким.",
        "Первое. Связь с Главой 2: ещё раз про плотность мощности I (интенсивность) — на примерах с расчётами на доске.",
        "Второе. Как работает линза: закон преломления Снеллиуса — базовая формула с расшифровкой.",
        "Третье. Формула №1: размер пятна d (диаметр пятна). От чего он зависит и почему вы не можете сделать его 1 мкм на вашем станке. Введём понятие числовой апертуры (NA).",
        "Четвёртое. Формула №2: глубина резкости (DOF). Почему уменьшение пятна требует идеально ровного слоя порошка. Разберём три случая — NA (числовая апертура) = 0,1, 0,4 и 0,5.",
        "Пятое. Формула №3: качество пучка M^2 (параметр «эм-квадрат»). Почему плохой пучок сводит на нет все преимущества хорошей оптики.",
    ]),
    ("3. ЗАЧЕМ НАМ ОПТИКА? (С РАСЧЁТАМИ НА ДОСКЕ)", "auto", [
        "Давайте вспомним ключевое понятие из Главы 2 — плотность мощности, или интенсивность I (интенсивность, плотность мощности). Это отношение мощности лазера P (мощность) к площади пятна S (площадь): I = P / S.",
        "Размерность I (интенсивности): ватт на квадратный сантиметр (Вт/см2) или ватт на квадратный миллиметр (Вт/мм2).",
        "Теперь — два сценария. Я буду писать расчёты на доске.",
        "Сценарий 1. Без фокусировки. Пусть диаметр пятна d (диаметр) = 1 мм. Переводим в сантиметры: 1 мм = 0,1 см. Площадь круга: S (площадь) = pi (d/2)^2 = pi (0,05)^2 ≈ 3,14 · 0,0025 = 0,00785 см2. Округлим до 0,008 см2.",
        "Мощность лазера P (мощность) = 200 Вт. Плотность мощности: I (интенсивность) = 200 / 0,008 = 25 000 Вт/см2 = 25 кВт/см2.",
        "Что произойдёт с металлом при I (интенсивности) = 25 кВт/см2? Он нагреется, может даже покраснеть. Но стабильного расплава не будет — слишком мало энергии на единицу площади.",
        "Сценарий 2. С фокусировкой. Ставим линзу и сжимаем луч в пятно d (диаметр) = 50 мкм. Переводим: 50 мкм = 0,005 см. Радиус r = 0,0025 см.",
        "Площадь: S (площадь) = pi r^2 = 3,14 · (0,0025)^2 = 3,14 · 6,25·10^-6 ≈ 1,96·10^-5 см2. Округлим до 2·10^-5 см2.",
        "I (интенсивность) = 200 / (2·10^-5) = 10 000 000 Вт/см2 = 10 МВт/см2. Это в 400 раз выше, чем в первом сценарии!",
        "Вывод: одна и та же мощность P (мощность), но результат кардинально разный. Вся магия — в линзе. Она не создаёт энергию, она её сжимает в пространстве. Поэтому оптика — главный усилитель лазерного станка.",
        "Вопрос: почему мы не можем сжать луч в бесконечно маленькое пятно d (диаметр)? Потому что есть физические ограничения — дифракция и глубина резкости (DOF). Об этом дальше.",
    ]),
    ("4. ЗАКОН СНЕЛЛИУСА — КАК РАБОТАЕТ ЛИНЗА", "auto", [
        "Теперь — как линза вообще работает. В основе лежит закон преломления Снеллиуса. Это базовая формула, которую полезно знать, даже если вы не будете её использовать для расчётов.",
        "n1 · sin(theta1) = n2 · sin(theta2).",
        "Расшифровка символов. n1 (показатель преломления первой среды) — например, воздуха. У воздуха n ≈ 1,0003, обычно принимают n = 1.",
        "n2 (показатель преломления второй среды) — например, стекла линзы. У кварцевого стекла n ≈ 1,45–1,5.",
        "theta1 — угол падения луча, отсчитанный от перпендикуляра (нормали) к поверхности. theta2 — угол преломления, тоже от нормали.",
        "Показатель преломления n — это отношение скорости света в вакууме c к скорости света в среде v: n = c / v.",
        "Физический смысл: когда свет переходит из среды с меньшим n (показателем преломления) в среду с большим n, он притягивается к нормали: theta2 < theta1. Когда выходит из стекла в воздух — наоборот, отклоняется от нормали.",
        "Как это создаёт линзу? Линза — кусок стекла с изогнутыми поверхностями. Лучи через края встречают поверхность под большим углом и преломляются сильнее. Лучи через центр — слабее. В результате все лучи собираются в одну точку — фокус.",
        "Вывод: разница показателей преломления Delta n (разность показателей преломления) на границе линзы определяет, насколько сильно она может преломить луч. Чем больше Delta n, тем сильнее фокусировка и тем короче фокусное расстояние.",
    ]),
    ("5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)", "auto", [
        "Теперь переходим к главному ограничению — дифракционному пределу. Даже идеальная линза не может собрать свет в математическую точку. Минимальный диаметр пятна определяется формулой: d (диаметр) ≈ lambda (длина волны) / NA (числовая апертура).",
        "Расшифровка символов. d — диаметр пятна в фокусе (в микрометрах). Это минимальный размер, который может дать данная оптика.",
        "lambda (лямбда, λ) — длина волны лазера. Для волоконного лазера lambda = 1,07 мкм.",
        "NA — числовая апертура (безразмерная величина, от 0 до 1). Мера того, насколько широкий конус света собирает линза.",
        "Что такое числовая апертура? NA = n · sin(theta), где n — показатель преломления среды (в воздухе n = 1), theta — половина угла схождения конуса лучей.",
        "Простыми словами: чем больше NA (числовая апертура), тем круче лучи сходятся в фокус и тем меньше пятно d (диаметр).",
        "Пример 1. Типичная линза для СЛП. NA (числовая апертура) = 0,1 — характерно для многих F-theta линз (телецентрических объективов). По-русски — «эф-тета линза». Пятно d остаётся круглым по всей рабочей области (например, 250×250 мм).",
        "Подставляем: d (диаметр) ≈ 1,07 / 0,1 = 10,7 мкм. Это идеальный расчёт. Реальный размер будет больше из-за качества пучка M^2 — об этом в конце лекции.",
        "Пример 2. Более сильная линза: NA (числовая апертура) = 0,4. Тогда d (диаметр) ≈ 1,07 / 0,4 = 2,7 мкм. Казалось бы — берём такую линзу и печатаем с разрешением 3 мкм! Но не всё так просто. Есть вторая формула, которая всё портит.",
    ]),
    ("6. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2) И ТРИ СЛУЧАЯ", "auto", [
        "Глубина резкости DOF (Depth of Focus) — расстояние вдоль оси луча, в пределах которого диаметр пятна d (диаметр) остаётся примерно таким же, как в фокусе. Если выйдете за этот диапазон, пятно увеличится, I (интенсивность, плотность мощности) упадёт, процесс плавления сорвётся.",
        "Формула: DOF (глубина резкости) ≈ lambda (длина волны) / NA^2 (числовая апертура в квадрате).",
        "Расшифровка. DOF — глубина резкости (в микрометрах). lambda — длина волны (1,07 мкм). NA — числовая апертура (в знаменателе в квадрате!).",
        "Обратите внимание: NA (числовая апертура) в квадрате в знаменателе. DOF (глубина резкости) падает быстрее, чем размер пятна d (диаметр). Это главный компромисс оптики.",
        "Рассмотрим три случая — от промышленного стандарта до экстремальной оптики.",
        "Случай А. NA (числовая апертура) = 0,1 (типичная промышленная линза): d (диаметр) ≈ 10,7 мкм; DOF (глубина резкости) ≈ 1,07 / 0,01 = 107 мкм ≈ 0,1 мм.",
        "На практике: при NA (числовой апертуре) = 0,1 DOF (глубина резкости) около 0,1 мм. Если порошковый слой имеет неровности ±0,05 мм (типично для фракции 20–40 мкм), пятно d остаётся стабильным. Процесс устойчив.",
        "Случай Б. NA (числовая апертура) = 0,4 (короткофокусная лабораторная оптика): d (диаметр) ≈ 2,7 мкм; DOF (глубина резкости) ≈ 1,07 / 0,16 ≈ 6,7 мкм.",
        "DOF (глубина резкости) всего 6,7 мкм. Как только порошок поднимется или опустится на 3–5 мкм, пятно d (диаметр) начнёт заметно увеличиваться. При смещении на 7 мкм I (интенсивность, плотность мощности) упадёт в 4 раза. Пришлось бы выравнивать слой с точностью до пары микрон — с порошком 20–40 мкм это невозможно.",
        "Закономерный вопрос: «А что, если взять линзу ещё круче — NA (числовая апертура) = 0,5? Пятно d ~2 мкм и фантастическое разрешение?»",
        "Случай В. NA (числовая апертура) = 0,5 (экстремальная оптика): d (диаметр) ≈ 2,14 мкм; DOF (глубина резкости) ≈ 1,07 / 0,25 ≈ 4,3 мкм. Фокус острый только в диапазоне ±2 мкм.",
        "Почему это невозможно в реальном СЛП-станке? Три причины.",
        "Причина №1: рельеф порошка. Фракция 20–40 мкм даёт рельеф слоя 10–20 мкм. Расстояние от фокуса до соседней частицы — 15 мкм, что в 3–4 раза больше DOF (глубины резкости). Пятно d (диаметр) на «горбах» огромное, на впадинах — нулевая I (интенсивность).",
        "Причина №2: физический предел малой DOF (глубины резкости). Даже при идеально выровненном слое луч «проваливается» в поры между частицами при сканировании.",
        "Причина №3: экономика и надёжность. NA (числовая апертура) = 0,5 требует порошка 1–5 мкм (дорого, взрывоопасно), активной автофокусировки и короткофокусной оптики у порошка — риск загрязнения линзы.",
        "Итог: линза с NA (числовой апертурой) = 0,5 — как микроскоп: идеальная картинка только на идеально плоском объекте. В СЛП — слой шероховатого порошка с перепадами в десятки микрон. Поэтому индустрия выбирает NA (числовую апертуру) ≈ 0,08–0,15.",
        "Таблица трёх случаев — на слайде и в конспекте.",
    ]),
    ("7. КАЧЕСТВО ПУЧКА M² И ПОДВЕДЕНИЕ ИТОГОВ", "auto", [
        "Третий важный фактор. Даже при идеальной линзе с большой NA (числовой апертурой) реальное пятно d (диаметр) больше, чем предсказывает d ≈ lambda (длина волны) / NA. Потому что реальный лазерный пучок — не идеально гауссов.",
        "Качество пучка описывается параметром M^2 (читается «эм-квадрат», параметр качества пучка). У идеального лазера M^2 = 1. У волоконных M^2 ≈ 1,1–1,5. У диодных M^2 ≈ 10–30.",
        "Формула реального пятна: d_реал (реальный диаметр) = M^2 (параметр качества пучка) · d_идеал (идеальный диаметр).",
        "Пример. Волоконный: M^2 = 1,3, d_идеал = 10 мкм → d_реал ≈ 13 мкм. Диодный: M^2 = 20, d_идеал = 10 мкм → d_реал ≈ 200 мкм.",
        "Вывод: диодный лазер при той же мощности P (мощность) 200 Вт даст пятно d (диаметр) в 20 раз больше. I (интенсивность, плотность мощности) будет в 400 раз меньше! Именно поэтому в СЛП используют волоконные лазеры с M^2 (параметром качества пучка) близким к 1.",
        "Итоговые тезисы.",
        "Первое. Плотность мощности I (интенсивность) = P (мощность) / S (площадь) — главный параметр. Линза — усилитель, сжимающий энергию в пространстве.",
        "Второе. Размер пятна d (диаметр) ≈ lambda (длина волны) / NA (числовая апертура) — обратно пропорционален числовой апертуре.",
        "Третье. Глубина резкости DOF ≈ lambda (длина волны) / NA^2 (числовая апертура в квадрате) — падает быстрее, чем d (диаметр). Малое пятно требует идеально ровного слоя.",
        "Четвёртое. Реальное пятно d_реал = M^2 (параметр качества пучка) · d_идеал — зависит от качества пучка. Большое M^2 сводит на нет преимущества оптики.",
        "Пятое. Почему не NA (числовая апертура) = 0,5 в СЛП? DOF (глубина резкости) ~4 мкм, а порошок имеет неровности 20–40 мкм. Промышленность выбирает NA ≈ 0,08–0,15: пятно d 30–50 мкм, DOF 100–300 мкм.",
    ]),
]


def _full_script():
    timed, *_rest = _time_script(SCRIPT)
    return timed

# ── PPTX ──────────────────────────────────────────────────────────────────

def _add_na_table(slide, ACCENT, LT):
    headers = NA_TABLE["headers"]
    rows = NA_TABLE["rows"]
    tbl = slide.shapes.add_table(
        1 + len(rows), len(headers),
        Inches(0.35), Inches(1.35), Inches(12.65), Inches(5.7),
    ).table
    col_w = [Inches(2.8), Inches(2.95), Inches(2.95), Inches(2.95)]
    for ci, w in enumerate(col_w[: len(headers)]):
        tbl.columns[ci].width = w
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = pretty_math(h)
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(13)
            p.font.color.rgb = LT
            p.font.name = "Calibri"
            p.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = ACCENT
    for ri, row in enumerate(rows, start=1):
        bg = RGBColor(0x24, 0x24, 0x3E) if ri % 2 else RGBColor(0x1E, 0x1E, 0x32)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = pretty_math(val)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12 if ci == 0 else 11)
                p.font.color.rgb = LT
                p.font.name = "Calibri"
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                p.word_wrap = True
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg


def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    ACCENT = RGBColor(0x00, 0x96, 0x88)
    ACCENT2 = RGBColor(0xE6, 0x4A, 0x19)
    LT = RGBColor(0xEE, 0xEE, 0xEE)
    GT = RGBColor(0x75, 0x75, 0x75)

    def abg(s):
        f = s.background.fill
        f.solid()
        f.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

    def _font_for(tx):
        if any(ch in tx for ch in "νλτωθħΔπ"):
            return "Cambria Math"
        return "Calibri"

    def at(s, l, t, w, h, tx, sz=18, b=False, c=LT, a=PP_ALIGN.LEFT):
        tx = pretty_math(tx)
        tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = tx
        p.font.size = Pt(sz)
        p.font.bold = b
        p.font.color.rgb = c
        p.font.name = _font_for(tx)
        p.alignment = a

    def bullets(s, items, left=1.5, top=1.55, w=6.2, sz=17):
        tb = s.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(5.5))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            it = pretty_math(it)
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = it
            p.font.size = Pt(sz)
            p.font.color.rgb = LT
            p.font.name = _font_for(it)
            p.space_after = Pt(12 if w >= 10 else 6)

    def hdr(s, title, accent=ACCENT, sz=26):
        abg(s)
        at(s, 1.5, 0.30, 10.3, 0.9, pretty_math(title), sz, True)
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.22), Inches(10.3), Pt(3))
        ln.fill.solid()
        ln.fill.fore_color.rgb = accent
        ln.line.fill.background()

    tmpdir = os.path.join(BASE, "_diagrams_tmp_lecture5")
    os.makedirs(tmpdir, exist_ok=True)

    slide_defs = [
        (None, [
            ("hdr", ("1. ГЕОМЕТРИЧЕСКАЯ ОПТИКА", 28)),
            ("at", (1.5, 2.0, 10.3, 1.0, "Лекция 5: как линзы собирают луч в пятно", 22)),
            ("at", (1.5, 3.0, 10.3, 1.2, "Плотность мощности, NA, DOF и M^2", 18, False, GT)),
            ("at", (1.5, 4.2, 10.3, 1.0, "Оптика как главный усилитель лазерного станка", 17)),
        ]),
        (None, [
            ("hdr", ("2. ПЛАН ЛЕКЦИИ", 30)),
            ("bullets", ([
                "1. Плотность мощности I = P/S — расчёты",
                "2. Закон Снеллиуса и работа линзы",
                "3. Формула №1: d ~ lambda/NA",
                "4. Формула №2: DOF ~ lambda/NA2",
                "5. Формула №3: M^2 и итоги",
            ], 1.5, 1.55, 10.3, 18)),
        ]),
        (draw_intensity_comparison_pil, [
            ("hdr", ("3. ЗАЧЕМ НАМ ОПТИКА?", 26), ACCENT2),
            ("bullets", ([
                "I = P / S — плотность мощности",
                "d = 1 мм -> I ~ 25 кВт/см2",
                "d = 50 мкм -> I ~ 10 МВт/см2",
                "x400 при той же мощности 200 Вт",
                "Линза сжимает энергию в пространстве",
            ], 1.5, 1.45, 5.4, 14)),
        ]),
        (draw_snell_lens_pil, [
            ("hdr", ("4. ЗАКОН СНЕЛЛИУСА", 26), ACCENT2),
            ("bullets", ([
                "n1 sin(theta1) = n2 sin(theta2)",
                "n = c / v — показатель преломления",
                "Края линзы преломляют сильнее",
                "Лучи сходятся в фокус",
            ], 1.5, 1.45, 5.4, 14)),
        ]),
        (draw_spot_size_pil, [
            ("hdr", ("5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)", 22), ACCENT2),
            ("bullets", ([
                "d ~ lambda / NA",
                "NA = n sin(theta)",
                "NA=0,1 -> d~10,7 мкм",
                "NA=0,4 -> d~2,7 мкм",
                "Идеал; реальность — M^2",
            ], 1.5, 1.45, 5.4, 13)),
        ]),
        ("table", [
            ("hdr", ("6. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2)", 22), ACCENT2),
        ]),
        (draw_m2_comparison_pil, [
            ("hdr", ("7. M^2 И ИТОГИ", 28), ACCENT2),
            ("bullets", ([
                "d_реал = M^2 * d_идеал",
                "Волокно M^2~1,3; диод M^2~20",
                "I ~ 1/d^2 — качество пучка критично",
                "NA~0,1 — компромисс для СЛП",
            ], 1.5, 1.45, 5.4, 13)),
        ]),
    ]

    for i, entry in enumerate(slide_defs):
        draw_fn = entry[0]
        actions = entry[1]
        pic = {"left": 6.8, "top": 1.45, "width": 5.9}
        s = prs.slides.add_slide(prs.slide_layouts[6])
        for act in actions:
            kind = act[0]
            if kind == "hdr":
                args = act[1]
                accent = act[2] if len(act) > 2 else ACCENT
                hdr(s, args[0], accent, args[1] if len(args) > 1 else 26)
            elif kind == "at":
                at(s, *act[1])
            elif kind == "bullets":
                a = act[1]
                bullets(s, a[0], a[1], a[2], a[3], a[4])
        if draw_fn == "table":
            _add_na_table(s, ACCENT, LT)
        elif draw_fn:
            png = save_pil_image(draw_fn, os.path.join(tmpdir, f"diag_{i}.png"))
            s.shapes.add_picture(
                png, Inches(pic["left"]), Inches(pic["top"]), width=Inches(pic["width"]),
            )

    prs.save(PPTX_PATH)
    print(f"OK: {PPTX_PATH}")


def build_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.3 * cm, bottomMargin=1.3 * cm,
    )
    styles = getSampleStyleSheet()
    sh2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, fontSize=13, leading=17)
    sh3 = ParagraphStyle("H3", parent=styles["Heading3"], fontName=PDF_FONT_BOLD, fontSize=11, leading=14)
    body = ParagraphStyle("B", parent=styles["Normal"], fontName=PDF_FONT, fontSize=10, leading=13, alignment=TA_JUSTIFY)
    hint = ParagraphStyle("H", parent=styles["Normal"], fontName=PDF_FONT, fontSize=9, leading=12, textColor=darkgray, alignment=TA_JUSTIFY)
    stage = ParagraphStyle(
        "Stage", parent=styles["Normal"], fontName=PDF_FONT, fontSize=9,
        leading=12, textColor=darkgray, alignment=TA_JUSTIFY, leftIndent=6,
    )

    pdf_tmp = os.path.join(BASE, "_diagrams_tmp_lecture5_pdf")
    os.makedirs(pdf_tmp, exist_ok=True)

    diagram_for_slide = {
        "3. ЗАЧЕМ НАМ ОПТИКА? (С РАСЧЁТАМИ НА ДОСКЕ)": draw_intensity_comparison_pil,
        "4. ЗАКОН СНЕЛЛИУСА — КАК РАБОТАЕТ ЛИНЗА": draw_snell_lens_pil,
        "5. РАЗМЕР ПЯТНА (ФОРМУЛА №1)": draw_spot_size_pil,
        "6. ГЛУБИНА РЕЗКОСТИ (ФОРМУЛА №2) И ТРИ СЛУЧАЯ": draw_dof_chart_pil,
        "7. КАЧЕСТВО ПУЧКА M² И ПОДВЕДЕНИЕ ИТОГОВ": draw_m2_comparison_pil,
    }

    script = _full_script()
    total_words = sum(sum(_count_words(p) for p in paras) for _, _, paras, _ in script)

    elements = []
    elements.append(Paragraph(pdf_sub("ЛЕКЦИЯ 5: ГЕОМЕТРИЧЕСКАЯ ОПТИКА"), sh2))
    elements.append(Paragraph(
        f"Раскадровка • (7 слайдов) • ~{total_words} слов • {datetime.now().strftime('%d.%m.%Y')}",
        hint,
    ))
    elements.append(Paragraph(
        "Тайминг «стенка» = произнесение текста (~130 слов/мин) + работа у доски/со схемой. "
        "Строки «Доска / пауза:» — ремарки лектору, не читаются вслух. Целевой хронометраж лекции: 40–45 минут.",
        hint,
    ))
    elements.append(Spacer(1, 8))

    for si, (title, timing, paragraphs, cue) in enumerate(script):
        block = []
        block.append(Paragraph(pdf_sub(f"СЛАЙД: {title} | {timing}"), sh3))
        if cue:
            block.append(Paragraph(pdf_sub(f"Доска / пауза: {cue}"), stage))
        for para in paragraphs:
            block.append(Paragraph(pdf_sub(para), body))
        block.append(Spacer(1, 5))

        draw_fn = diagram_for_slide.get(title)
        if draw_fn:
            img = draw_fn()
            safe = re.sub(r"[^\w\-]+", "_", title)[:40]
            png = os.path.join(pdf_tmp, f"pdf_{safe}.png")
            img.save(png, "PNG")
            w_px, h_px = img.size
            img_w = 16.5 * cm if w_px >= 900 else 13 * cm
            block.append(RLImage(png, width=img_w, height=img_w * (h_px / w_px)))
            block.append(Spacer(1, 6))

        if si == len(script) - 1:
            block.append(Paragraph(pdf_sub(EPILOGUE), body))
            elements.append(KeepTogether(block))
        else:
            elements.extend(block)

    def _page_num(canvas, _doc):
        canvas.saveState()
        canvas.setFont(PDF_FONT, 10)
        canvas.setFillColor(gray)
        canvas.drawCentredString(A4[0] / 2, 0.85 * cm, str(canvas.getPageNumber()))
        canvas.restoreState()

    doc.build(elements, onFirstPage=_page_num, onLaterPages=_page_num)
    print(f"OK: {PDF_PATH} (~{total_words} слов)")


def build_docx():
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn

    script = _full_script()
    total_words = sum(sum(_count_words(p) for p in paras) for _, _, paras, _ in script)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.5)
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    style.font.size = Pt(11)

    def add(text, *, bold=False, size=11, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"
        return p

    add("ЛЕКЦИЯ 5: ГЕОМЕТРИЧЕСКАЯ ОПТИКА: КАК ЛИНЗЫ СОБИРАЮТ ЛУЧ", bold=True, size=14, space_after=8)
    add(
        f"Раскадровка для лектора • ~{total_words} слов • "
        f"{datetime.now().strftime('%d.%m.%Y')}",
        size=10, space_after=10,
    )
    add(
        "Тайминг «стенка» = произнесение текста (~130 слов/мин) + работа у доски/со схемой. "
        "Строки «Доска / пауза:» — ремарки лектору, не читаются вслух. Целевой хронометраж: 40–45 минут.",
        size=9, space_after=12,
    )

    for title, timing, paragraphs, cue in script:
        add(f"СЛАЙД: {pretty_math(title)}", bold=True, size=12, space_after=2)
        add(timing, size=10, space_after=4)
        if cue:
            add(f"Доска / пауза: {pretty_math(cue)}", size=10, space_after=6)
        for para in paragraphs:
            add(pretty_math(para), size=11, space_after=6)

    add(pretty_math(EPILOGUE), size=11, space_after=6)

    doc.save(DOCX_PATH)
    print(f"OK: {DOCX_PATH}")


if __name__ == "__main__":
    print("=" * 70)
    print("СБОРКА ЛЕКЦИИ 5")
    print("=" * 70)
    build_pptx()
    build_pdf()
    build_docx()
    print("=" * 70)
    print("ГОТОВО")
    print("=" * 70)
