#!/usr/bin/env python3
"""
Yahoo Fantasy Baseball Analyzer v2 — Deck Generator
Design system: fetch_template_with_backgrounds.yaml (背景圖導向)
Slide count: 8 (Cover + Agenda + 5 content + Closing)
"""

import os
import json
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─────────────────────────────────────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR    = r"c:\Users\00905847\antropic skills\skills_test"
SKILL_ASSETS = os.path.join(BASE_DIR, "skills", "pptx", "assets", "backgrounds")
COVER_BG_SRC   = os.path.join(SKILL_ASSETS, "cover-background.png")
CONTENT_BG_SRC = os.path.join(SKILL_ASSETS, "content-background.png")

DECK_NAME   = "Yahoo-Fantasy-Baseball-Analyzer-Deck"
BUNDLE_DIR  = os.path.join(BASE_DIR, DECK_NAME)
ASSETS_DIR  = os.path.join(BUNDLE_DIR, "assets")
BG_DIR      = os.path.join(ASSETS_DIR, "backgrounds")
OUTPUT_PPTX = os.path.join(BUNDLE_DIR, f"{DECK_NAME}.pptx")
COVER_BG    = os.path.join(BG_DIR, "cover-background.png")
CONTENT_BG  = os.path.join(BG_DIR, "content-background.png")

# ─────────────────────────────────────────────────────────────────────────────
# Design tokens  (from fetch_template_with_backgrounds.yaml)
# ─────────────────────────────────────────────────────────────────────────────
FONT = "Microsoft JhengHei"

def rgb(r, g, b):
    return RGBColor(r, g, b)

C = {
    "accent":    rgb(0x38, 0xB0, 0x6A),
    "softGreen": rgb(0xD6, 0xF0, 0xE3),
    "green2":    rgb(0x2E, 0x94, 0x59),
    "green3":    rgb(0x21, 0x7A, 0x4A),
    "text":      rgb(0x1A, 0x1A, 0x1A),
    "text2":     rgb(0x4A, 0x4A, 0x4A),
    "text3":     rgb(0x7A, 0x7A, 0x7A),
    "surface":   rgb(0xF5, 0xF5, 0xF5),
    "border":    rgb(0xE8, 0xE8, 0xE8),
    "panel":     rgb(0xEF, 0xEF, 0xEF),
    "white":     rgb(0xFF, 0xFF, 0xFF),
    "yellow":    rgb(0xF5, 0xC5, 0x18),
    "orange":    rgb(0xF7, 0x58, 0x01),
}

# Slide dimensions
SW = Inches(10)
SH = Inches(5.625)

# Content safe area (from template)
SA_X = Inches(0.90)
SA_Y = Inches(0.45)
SA_W = Inches(8.60)
SA_H = Inches(4.60)

# ─────────────────────────────────────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────────────────────────────────────

def set_bg(slide, img_path):
    """Place image as full-slide background (lowest z-order)."""
    pic = slide.shapes.add_picture(img_path, 0, 0, SW, SH)
    sp = pic._element
    sp.getparent().remove(sp)
    slide.shapes._spTree.insert(2, sp)


def add_rect(slide, x, y, w, h, fill_color, line_color=None, line_width=Pt(1)):
    """Add a filled rectangle; line_color=None removes the border."""
    shape = slide.shapes.add_shape(1, x, y, w, h)  # 1 = MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, x, y, w, h,
             font_size, bold=True, color=None,
             align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txBox


def card_box(slide, x, y, w, h, accent_color=None):
    """Draw surface card with optional top accent bar. Returns inner-y after bar."""
    add_rect(slide, x, y, w, h, C["surface"], C["border"])
    bar_h = Pt(7)
    if accent_color:
        add_rect(slide, x, y, w, bar_h, accent_color)
    return y + (bar_h if accent_color else 0)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 1 – Cover
# ─────────────────────────────────────────────────────────────────────────────
def slide_cover(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, COVER_BG)

    # White content panel — center of slide
    px, py, pw, ph = Inches(1.3), Inches(1.2), Inches(7.4), Inches(3.3)
    add_rect(sl, px, py, pw, ph, C["white"])
    # Green accent bar at top of white panel
    add_rect(sl, px, py, pw, Pt(8), C["accent"])

    # Slide badge top-left of panel
    add_text(sl, "⚾  Yahoo Fantasy Baseball", px + Inches(0.22), py + Inches(0.18),
             Inches(6.5), Inches(0.36), 14, bold=True, color=C["accent"])

    # Main title
    add_text(sl, "Analyzer v2", px + Inches(0.22), py + Inches(0.60),
             Inches(6.8), Inches(0.85), 40, bold=True, color=C["text"])

    # Divider line
    add_rect(sl, px + Inches(0.22), py + Inches(1.5), Inches(2.2), Pt(3), C["accent"])

    # Subtitle
    add_text(sl, "智能棒球夢幻聯盟球員評估分析工具",
             px + Inches(0.22), py + Inches(1.65), Inches(6.8), Inches(0.42),
             16, bold=True, color=C["text2"])

    # Tagline
    add_text(sl, "支援 H2H Categories / H2H Points 雙模式  ·  CSV / JSON 資料匯入  ·  雷達圖比較  ·  交易評估",
             px + Inches(0.22), py + Inches(2.12), Inches(6.8), Inches(0.32),
             13, bold=True, color=C["text3"])

    # Date
    add_text(sl, "2026.05", px + Inches(0.22), py + Inches(2.6),
             Inches(2.0), Inches(0.28), 13, bold=True, color=C["text3"])


# ─────────────────────────────────────────────────────────────────────────────
# Slide 2 – Agenda (4 feature cards)
# ─────────────────────────────────────────────────────────────────────────────
def slide_agenda(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    add_text(sl, "功能模組總覽", SA_X, SA_Y, SA_W, Inches(0.42),
             24, bold=True, color=C["text"])
    add_text(sl, "四大核心分析模組，覆蓋完整夢幻棒球決策鏈",
             SA_X, SA_Y + Inches(0.42), SA_W, Inches(0.28),
             14, bold=True, color=C["accent"])

    modules = [
        ("01", "打者分析",   "valueScore 綜合評分 / 位置篩選 / KPI 過濾 / Elite‥Risk 分層", C["accent"]),
        ("02", "投手分析",   "ERA / WHIP / K-BB% / K/9 / 勝投存活 / 先發‧救援篩選",           C["green2"]),
        ("03", "球員比較",   "雙球員指標對比 / 雷達圖 / 跨打投共同指標比較",                   C["yellow"]),
        ("04", "交易評估",   "多球員換算模擬 / 勝負判定 / 差距橫條視覺化",                     C["orange"]),
    ]

    cw = Inches(4.1)
    ch = Inches(1.75)
    cx0 = SA_X
    cy0 = SA_Y + Inches(0.85)
    gap_x = Inches(0.25)
    gap_y = Inches(0.18)

    for i, (num, title, body, acc) in enumerate(modules):
        cx = cx0 + (i % 2) * (cw + gap_x)
        cy = cy0 + (i // 2) * (ch + gap_y)
        inner_y = card_box(sl, cx, cy, cw, ch, acc)

        # Large number
        add_text(sl, num, cx + Inches(0.12), inner_y + Inches(0.08),
                 Inches(0.55), Inches(0.55), 36, bold=True, color=acc)
        # Title
        add_text(sl, title, cx + Inches(0.70), inner_y + Inches(0.12),
                 cw - Inches(0.82), Inches(0.36), 18, bold=True, color=C["text"])
        # Body
        add_text(sl, body, cx + Inches(0.12), inner_y + Inches(0.62),
                 cw - Inches(0.24), Inches(0.65), 13, bold=True, color=C["text2"])


# ─────────────────────────────────────────────────────────────────────────────
# Slide 3 – 球員分析 (Left Panel + Right 3-section)
# ─────────────────────────────────────────────────────────────────────────────
def slide_player_analysis(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    # Left panel
    lw = Inches(2.5)
    add_rect(sl, SA_X, SA_Y, lw, SA_H, C["panel"], C["border"])
    add_rect(sl, SA_X, SA_Y, Pt(7), SA_H, C["accent"])

    add_text(sl, "02", SA_X + Inches(0.14), SA_Y + Inches(0.12),
             Inches(0.6), Inches(0.55), 36, bold=True, color=C["accent"])
    add_text(sl, "球員\n分析", SA_X + Inches(0.14), SA_Y + Inches(0.75),
             lw - Inches(0.25), Inches(0.72), 22, bold=True, color=C["text"])
    add_text(sl, "打者 & 投手\n雙軌評分系統", SA_X + Inches(0.14), SA_Y + Inches(1.55),
             lw - Inches(0.25), Inches(0.55), 14, bold=True, color=C["accent"])
    add_text(sl, "依玩法模式自動\n切換評分公式，\n精準反映聯盟\n規則差異。",
             SA_X + Inches(0.14), SA_Y + Inches(2.2), lw - Inches(0.25), Inches(1.2),
             13, bold=True, color=C["text2"])

    # Right 3 sections
    rx = SA_X + lw + Inches(0.22)
    rw = SA_W - lw - Inches(0.22)
    section_h = (SA_H - Inches(0.18)) / 3

    sections = [
        ("⚾  打者模組（Batters）",
         "G · AB · R · HR · RBI · SB · AVG · OBP · SLG\n"
         "H2H Cat 模式下各類別獨立比較，Points 模式下依積分換算總分。",
         C["accent"]),
        ("🎯  投手模組（Pitchers）",
         "W · SV · K · ERA · WHIP · K/9 · K-BB%\n"
         "先發 / 救援可個別篩選；支援「防禦率類別」的懲罰係數調整。",
         C["green2"]),
        ("⚡  評等分層（Tier System）",
         "Elite（頂尖）  Solid（穩定）  Watch（觀察）  Risk（風險）\n"
         "依全聯盟球員百分位自動分層，協助快速識別 FA 撿人時機。",
         C["yellow"]),
    ]

    for i, (title, body, acc) in enumerate(sections):
        sy = SA_Y + i * (section_h + Inches(0.06))
        inner_y = card_box(sl, rx, sy, rw, section_h - Inches(0.06), acc)
        add_text(sl, title, rx + Inches(0.14), inner_y + Inches(0.1),
                 rw - Inches(0.28), Inches(0.34), 15, bold=True, color=C["text"])
        add_text(sl, body, rx + Inches(0.14), inner_y + Inches(0.48),
                 rw - Inches(0.28), Inches(0.8), 13, bold=True, color=C["text2"])


# ─────────────────────────────────────────────────────────────────────────────
# Slide 4 – 智能輔助功能 (Horizontal 3-row)
# ─────────────────────────────────────────────────────────────────────────────
def slide_smart_features(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    add_text(sl, "03  智能輔助功能", SA_X, SA_Y, SA_W, Inches(0.38),
             20, bold=True, color=C["text"])
    add_text(sl, "讓資料說話，讓每個決策都有數據支撐",
             SA_X, SA_Y + Inches(0.38), SA_W, Inches(0.28),
             14, bold=True, color=C["accent"])

    rows = [
        ("📊  valueScore 綜合評分",
         "依 H2H Categories 或 Points 模式，對每位球員計算標準化分數。"
         "整合多項統計指標為單一排名依據，支援攻守類別加權 + 風險調整滑桿。",
         C["accent"]),
        ("🎛️  KPI 過濾 + 欄位客製",
         "點擊 KPI 卡片可快速篩選該指標強球員；欄位開關讓你只看需要的統計欄。"
         "提供三種一鍵預設：精簡 / 標準 / 完整。",
         C["green2"]),
        ("⭐  觀察名單 + 靈活資料匯入",
         "星號標記目標球員，清單自動儲存（localStorage）。"
         "支援 CSV / JSON 拖拽上傳、文字貼上，亦可自動讀取本地 Yahoo 匯出 CSV。",
         C["yellow"]),
    ]

    row_h = Inches(1.2)
    gap   = Inches(0.1)
    for i, (title, body, acc) in enumerate(rows):
        ry = SA_Y + Inches(0.78) + i * (row_h + gap)
        inner_y = card_box(sl, SA_X, ry, SA_W, row_h, acc)
        add_text(sl, title, SA_X + Inches(0.14), inner_y + Inches(0.1),
                 SA_W - Inches(0.28), Inches(0.32), 15, bold=True, color=C["text"])
        add_text(sl, body, SA_X + Inches(0.14), inner_y + Inches(0.46),
                 SA_W - Inches(0.28), Inches(0.65), 13, bold=True, color=C["text2"])


# ─────────────────────────────────────────────────────────────────────────────
# Slide 5 – 玩法模式比較 (Comparison Table)
# ─────────────────────────────────────────────────────────────────────────────
def slide_mode_comparison(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    # Header bar
    hdr_h = Inches(0.65)
    add_rect(sl, SA_X, SA_Y, SA_W, hdr_h, C["surface"], C["border"])
    add_rect(sl, SA_X, SA_Y, SA_W, Pt(7), C["accent"])
    add_text(sl, "04  玩法模式比較", SA_X + Inches(0.14), SA_Y + Inches(0.12),
             Inches(3.5), Inches(0.38), 18, bold=True, color=C["text"])
    add_text(sl, "H2H Categories vs H2H Points  —  選對模式，制定最佳陣容策略",
             SA_X + Inches(3.7), SA_Y + Inches(0.18), Inches(5.1), Inches(0.28),
             13, bold=True, color=C["text3"])

    # Column title row
    col_gap  = Inches(0.12)
    label_w  = Inches(1.35)
    col_w    = (SA_W - label_w - col_gap * 3) / 2
    cy = SA_Y + hdr_h + Inches(0.1)
    col_hdr_h = Inches(0.40)

    # Label placeholder
    add_rect(sl, SA_X, cy, label_w, col_hdr_h, C["panel"], C["border"])

    for j, (lbl, acc) in enumerate([("H2H Categories（類別制）", C["accent"]),
                                     ("H2H Points（積分制）",      C["orange"])]):
        cx = SA_X + label_w + col_gap + j * (col_w + col_gap)
        add_rect(sl, cx, cy, col_w, col_hdr_h, acc)
        add_text(sl, lbl, cx + Inches(0.1), cy + Inches(0.06),
                 col_w - Inches(0.2), Inches(0.28), 13, bold=True, color=C["white"])

    # Data rows
    row_data = [
        ("勝負判定",
         "每類別各自比較勝負，統計總勝負場次",
         "每週積分總和決勝，無類別概念"),
        ("策略重點",
         "均衡各指標，避免嚴重缺漏類別",
         "集中培養高積分球員，單指標豪賭可行"),
        ("球員價值",
         "平衡型球員更有價值，雙刀流選手稀缺",
         "極端表現者（大 K 先發）更值錢"),
        ("應避免",
         "極端高 K 但 ERA 差的投手（拖累類別）",
         "ERA 指標被積分稀釋，只看總積分量"),
    ]

    row_h = Inches(0.68)
    gap_r = Inches(0.05)
    for k, (cat, cat_txt, pts_txt) in enumerate(row_data):
        ry = cy + col_hdr_h + Inches(0.06) + k * (row_h + gap_r)
        # Category label
        add_rect(sl, SA_X, ry, label_w, row_h, C["surface"], C["border"])
        add_text(sl, cat, SA_X + Inches(0.08), ry + Inches(0.18),
                 label_w - Inches(0.16), Inches(0.32), 12, bold=True, color=C["text3"])
        # H2H Cat cell
        cx1 = SA_X + label_w + col_gap
        add_rect(sl, cx1, ry, col_w, row_h, C["surface"], C["border"])
        add_text(sl, cat_txt, cx1 + Inches(0.1), ry + Inches(0.1),
                 col_w - Inches(0.2), row_h - Inches(0.12), 12, bold=True, color=C["text2"])
        # H2H Points cell
        cx2 = cx1 + col_w + col_gap
        add_rect(sl, cx2, ry, col_w, row_h, C["surface"], C["border"])
        add_text(sl, pts_txt, cx2 + Inches(0.1), ry + Inches(0.1),
                 col_w - Inches(0.2), row_h - Inches(0.12), 12, bold=True, color=C["text2"])


# ─────────────────────────────────────────────────────────────────────────────
# Slide 6 – 使用流程 (Flow 5-step)
# ─────────────────────────────────────────────────────────────────────────────
def slide_workflow(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    add_text(sl, "05  使用流程", SA_X, SA_Y, SA_W, Inches(0.38),
             20, bold=True, color=C["text"])
    add_text(sl, "五步驟完成夢幻聯盟全方位分析",
             SA_X, SA_Y + Inches(0.38), SA_W, Inches(0.28),
             14, bold=True, color=C["accent"])

    steps = [
        ("01", "匯入資料",  "CSV / JSON\n拖拽或貼上\n自動讀取本地"),
        ("02", "選擇模式",  "H2H Categories\n或\nH2H Points"),
        ("03", "查看排行",  "打者頁 / 投手頁\nKPI 過濾\n欄位客製"),
        ("04", "比較球員",  "雙球員對比\n雷達圖\n差距視覺化"),
        ("05", "評估交易",  "輸入換算名單\n自動計算差距\n勝負判定輸出"),
    ]

    n    = len(steps)
    sw   = (SA_W - Inches(0.2)) / n
    fy   = SA_Y + Inches(0.85)
    fh   = Inches(3.6)
    circ = Inches(0.56)

    for i, (num, title, body) in enumerate(steps):
        fx = SA_X + i * (sw + Inches(0.05))

        # Step circle
        cy_c = fy
        add_rect(sl, fx + (sw - circ) / 2, cy_c, circ, circ, C["accent"])
        add_text(sl, num,
                 fx + (sw - circ) / 2, cy_c + Inches(0.06),
                 circ, Inches(0.44), 15, bold=True, color=C["white"],
                 align=PP_ALIGN.CENTER)

        # Arrow (except last)
        if i < n - 1:
            ax = fx + sw + Inches(0.025)
            add_text(sl, "▶", ax, cy_c + Inches(0.1), Inches(0.15), Inches(0.36),
                     14, bold=True, color=C["border"])

        # Content card
        card_y = fy + Inches(0.65)
        card_h = fh - Inches(0.65)
        inner  = card_box(sl, fx, card_y, sw - Inches(0.05), card_h, C["accent"])
        add_text(sl, title,
                 fx + Inches(0.08), inner + Inches(0.1),
                 sw - Inches(0.21), Inches(0.36), 14, bold=True, color=C["text"],
                 align=PP_ALIGN.CENTER)
        add_text(sl, body,
                 fx + Inches(0.08), inner + Inches(0.52),
                 sw - Inches(0.21), Inches(1.3), 12, bold=True, color=C["text2"],
                 align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
# Slide 7 – 進階工具 (Three-Card)
# ─────────────────────────────────────────────────────────────────────────────
def slide_advanced_tools(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    add_text(sl, "06  進階分析工具", SA_X, SA_Y, SA_W, Inches(0.38),
             20, bold=True, color=C["text"])
    add_text(sl, "深度決策支援，讓每個動作都有數據依據",
             SA_X, SA_Y + Inches(0.38), SA_W, Inches(0.28),
             14, bold=True, color=C["accent"])

    cards = [
        ("Rule 01", "🔍  球員比較", C["accent"],
         "同時比較兩位球員所有指標，自動計算各統計面向差距，"
         "以長條圖與雷達圖呈現強弱對比。支援跨打投比較（顯示共同指標）。",
         "💡 最佳應用：FA 撿人決策 / 交易人選評估"),
        ("Rule 02", "🔄  交易評估", C["yellow"],
         "設定「給出」與「接收」名單，系統自動加總雙方 valueScore 及各項統計均值，"
         "輸出交易勝負判定（Win / Lose / Even）與差距橫條圖。",
         "💡 最佳應用：多對多複雜交易談判"),
        ("Rule 03", "⭐  觀察名單", C["orange"],
         "點擊球員旁的星號即可加入觀察名單，資料自動儲存於瀏覽器（localStorage），"
         "重新整理頁面也不會遺失。可快速切換至觀察名單視圖。",
         "💡 最佳應用：標記目標球員，等待最佳時機出手"),
    ]

    cw = (SA_W - Inches(0.3)) / 3
    ch = SA_H - Inches(0.85)

    for i, (rule, title, acc, body, tip) in enumerate(cards):
        cx = SA_X + i * (cw + Inches(0.15))
        cy = SA_Y + Inches(0.85)
        inner_y = card_box(sl, cx, cy, cw, ch, acc)

        add_text(sl, rule, cx + Inches(0.12), inner_y + Inches(0.1),
                 cw - Inches(0.24), Inches(0.28), 13, bold=True, color=acc)
        add_text(sl, title, cx + Inches(0.12), inner_y + Inches(0.38),
                 cw - Inches(0.24), Inches(0.36), 15, bold=True, color=C["text"])

        # Divider
        add_rect(sl, cx + Inches(0.12), inner_y + Inches(0.80),
                 cw - Inches(0.24), Pt(1), C["border"])

        add_text(sl, body, cx + Inches(0.12), inner_y + Inches(0.92),
                 cw - Inches(0.24), Inches(1.65), 12, bold=True, color=C["text2"])

        # Tip box
        tip_y = inner_y + ch - Pt(7) - Inches(0.58)
        add_rect(sl, cx + Inches(0.12), tip_y,
                 cw - Inches(0.24), Inches(0.52), C["softGreen"])
        add_text(sl, tip, cx + Inches(0.17), tip_y + Inches(0.08),
                 cw - Inches(0.34), Inches(0.40), 11, bold=True, color=C["text3"])


# ─────────────────────────────────────────────────────────────────────────────
# Slide 8 – Closing
# ─────────────────────────────────────────────────────────────────────────────
def slide_closing(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, CONTENT_BG)

    # Left accent line
    add_rect(sl, SA_X, SA_Y, Pt(7), SA_H, C["accent"])

    lx = SA_X + Inches(0.2)

    add_text(sl, "Yahoo Fantasy Baseball Analyzer v2",
             lx, SA_Y + Inches(0.25), Inches(7.5), Inches(0.55),
             26, bold=True, color=C["text"])
    add_rect(sl, lx, SA_Y + Inches(0.88), Inches(2.6), Pt(3), C["accent"])
    add_text(sl, "你的夢幻聯盟決策，從此更有依據。",
             lx, SA_Y + Inches(1.0), Inches(7.5), Inches(0.38),
             16, bold=True, color=C["accent"])

    rows = [
        ("⚾", "完整打者 + 投手分析",   "涵蓋主流 H2H 統計類別，雙模式一鍵切換"),
        ("📊", "智能評分與分層",         "自動計算 valueScore，Elite / Solid / Watch / Risk"),
        ("🔄", "交易 + 比較工具",         "數據支援的決策，不再靠直覺"),
        ("📁", "靈活資料匯入",           "CSV / JSON / 拖拽 / 本地自動讀取"),
    ]

    for i, (icon, title, body) in enumerate(rows):
        ry = SA_Y + Inches(1.58) + i * Inches(0.74)
        add_rect(sl, lx, ry, SA_W - Inches(0.2), Inches(0.65), C["surface"], C["border"])
        # Icon bg
        add_rect(sl, lx, ry, Inches(0.55), Inches(0.65), C["softGreen"])
        add_text(sl, icon, lx + Inches(0.04), ry + Inches(0.1),
                 Inches(0.47), Inches(0.45), 18, bold=True,
                 align=PP_ALIGN.CENTER)
        add_text(sl, title, lx + Inches(0.65), ry + Inches(0.06),
                 Inches(3.0), Inches(0.30), 14, bold=True, color=C["text"])
        add_text(sl, body, lx + Inches(0.65), ry + Inches(0.36),
                 Inches(5.5), Inches(0.26), 12, bold=True, color=C["text3"])


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # 1. Build bundle folder
    os.makedirs(BG_DIR, exist_ok=True)
    shutil.copy2(COVER_BG_SRC, COVER_BG)
    shutil.copy2(CONTENT_BG_SRC, CONTENT_BG)
    print(f"[+] Bundle folder: {BUNDLE_DIR}")

    # 2. Build presentation
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH

    print("[+] Generating slides...")
    slide_cover(prs)
    print("  1/8 Cover")
    slide_agenda(prs)
    print("  2/8 Agenda")
    slide_player_analysis(prs)
    print("  3/8 Player Analysis")
    slide_smart_features(prs)
    print("  4/8 Smart Features")
    slide_mode_comparison(prs)
    print("  5/8 Mode Comparison")
    slide_workflow(prs)
    print("  6/8 Workflow")
    slide_advanced_tools(prs)
    print("  7/8 Advanced Tools")
    slide_closing(prs)
    print("  8/8 Closing")

    # 3. Save
    prs.save(OUTPUT_PPTX)
    print(f"[+] Saved: {OUTPUT_PPTX}")

    # 4. Resource manifest
    manifest = {
        "deck": DECK_NAME,
        "template": "fetch_template_with_backgrounds.yaml",
        "generated": "2026-05-25",
        "slides": 8,
        "resources": [
            {"type": "background", "role": "cover",
             "source": COVER_BG_SRC, "bundled_as": "assets/backgrounds/cover-background.png"},
            {"type": "background", "role": "content",
             "source": CONTENT_BG_SRC, "bundled_as": "assets/backgrounds/content-background.png"},
        ]
    }
    with open(os.path.join(ASSETS_DIR, "resource-manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print("[+] resource-manifest.json written")
    print("[✓] Bundle complete.")


if __name__ == "__main__":
    main()
