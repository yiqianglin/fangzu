# -*- coding: utf-8 -*-
"""
【固化版】把收据 xlsx 渲染成 jpg —— 输入 = 收据 xlsx 目录，无 hardcode。

用法：
    python 每月租金生成/通用/导出收据图片.py --year 2026 --month 6
    python 每月租金生成/通用/导出收据图片.py --year 2026 --month 6 --village 变电站

输入：
    {ROOT}/{village}/{year}年{month}月/收据/*.xlsx     # 第 3 步生成收据.py 产出

输出：
    {ROOT}/{village}/{year}年{month}月/收据图片/*.jpg

铁律（写进 SKILL.md）：
- K7（房租行备注）红色字体（#D32F2F），其它列保持黑色
- K5 / K6 / K8 / K9 黑色（K8/K9 普通备注，无下月预告）
- 自动换行（最多 2 行，溢出末尾省略号）
- jpg 尺寸：720 × 314（动态高度由行高合计算出）
"""
import argparse
import os
import sys
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont


# ==================== 配置 ====================
FONT_PATH = "/System/Library/Fonts/STHeiti Medium.ttc"
IMG_WIDTH = 720
MARGIN_LEFT = 12
MARGIN_TOP = 8

BLACK = (0, 0, 0)
RED = (211, 47, 47)        # K7 下月预告专用
LIGHT_BG = (255, 255, 248)
LINE_COLOR = (60, 60, 60)
HEADER_BG = (245, 245, 235)


def load_fonts():
    return {
        "title": ImageFont.truetype(FONT_PATH, 20),
        "normal": ImageFont.truetype(FONT_PATH, 14),
        "small": ImageFont.truetype(FONT_PATH, 13),
        "bold": ImageFont.truetype(FONT_PATH, 15),
        "amount": ImageFont.truetype(FONT_PATH, 13),
    }


def draw_receipt(ws, output_path, fonts):
    MARGIN_BOTTOM = 8
    actual_height = MARGIN_TOP + 30 + 24 + (24 + 20 + 30 + 30 + 30 + 26 + 26 + 32 + 26) + MARGIN_BOTTOM
    img = Image.new("RGB", (IMG_WIDTH, actual_height), LIGHT_BG)
    draw = ImageDraw.Draw(img)
    x0 = MARGIN_LEFT
    y0 = MARGIN_TOP

    title = str(ws['A1'].value or "房租、水、电费（专用）收据")
    title_w = draw.textlength(title, font=fonts["title"])
    draw.text(((IMG_WIDTH - title_w) / 2, y0), title, fill=BLACK, font=fonts["title"])
    y0 += 30

    user_name = str(ws['A2'].value or "")
    date_str = str(ws['H2'].value or "")
    draw.text((x0, y0), user_name, fill=BLACK, font=fonts["normal"])
    date_w = draw.textlength(date_str, font=fonts["normal"])
    draw.text((IMG_WIDTH - MARGIN_LEFT - date_w, y0), date_str, fill=BLACK, font=fonts["normal"])
    y0 += 24

    table_width = IMG_WIDTH - 2 * MARGIN_LEFT
    col_widths = [80, 72, 72, 56, 56, 48, 48, 48, 48, 48, 0]
    col_widths[10] = table_width - sum(col_widths[:10])
    col_x = [x0]
    for w in col_widths:
        col_x.append(col_x[-1] + w)
    table_right = col_x[-1]

    row_heights = [24, 20, 30, 30, 30, 26, 26, 32, 26]

    def draw_row_bg(y, h, color):
        draw.rectangle([x0, y, table_right, y + h], fill=color)

    def draw_hline(y):
        draw.line([(x0, y), (table_right, y)], fill=LINE_COLOR, width=1)

    def draw_vlines(y, h):
        for i in range(len(col_x)):
            draw.line([(col_x[i], y), (col_x[i], y + h)], fill=LINE_COLOR, width=1)

    def text_center(text, col_idx, y, h, font_key="normal", color=BLACK):
        text = str(text) if text is not None else ""
        if not text:
            return
        f = fonts[font_key]
        tw = draw.textlength(text, font=f)
        tx = col_x[col_idx] + (col_widths[col_idx] - tw) / 2
        ty = y + (h - f.size) / 2 - 1
        draw.text((tx, ty), text, fill=color, font=f)

    def text_center_span(text, col_start, col_end, y, h, font_key="normal", color=BLACK):
        text = str(text) if text is not None else ""
        if not text:
            return
        f = fonts[font_key]
        tw = draw.textlength(text, font=f)
        span_left = col_x[col_start]
        span_right = col_x[col_end] + col_widths[col_end]
        tx = span_left + (span_right - span_left - tw) / 2
        ty = y + (h - f.size) / 2 - 1
        draw.text((tx, ty), text, fill=color, font=f)

    def text_left(text, x, y, h, font_key="normal"):
        f = fonts[font_key]
        ty = y + (h - f.size) / 2 - 1
        draw.text((x + 4, ty), str(text), fill=BLACK, font=f)

    def text_right(text, x_right, y, h, font_key="normal"):
        f = fonts[font_key]
        tw = draw.textlength(str(text), font=f)
        ty = y + (h - f.size) / 2 - 1
        draw.text((x_right - tw - 4, ty), str(text), fill=BLACK, font=f)

    def draw_k_cell(text, y, h, font_key="small", color=BLACK):
        """K 列（备注列）自动换行（最多 2 行）"""
        if not text:
            return
        f = fonts[font_key]
        cell_w = col_widths[10] - 4
        words = list(str(text))
        lines, cur = [], ""
        for ch in words:
            if draw.textlength(cur + ch, font=f) > cell_w:
                lines.append(cur)
                cur = ch
            else:
                cur += ch
        if cur:
            lines.append(cur)
        if len(lines) > 2:
            lines = lines[:2]
            while draw.textlength(lines[1] + "…", font=f) > cell_w and len(lines[1]) > 1:
                lines[1] = lines[1][:-1]
            lines[1] += "…"
        total_h = len(lines) * (f.size + 2)
        ty = y + (h - total_h) / 2
        for line in lines:
            tw = draw.textlength(line, font=f)
            tx = col_x[10] + (col_widths[10] - tw) / 2
            draw.text((tx, ty), line, fill=color, font=f)
            ty += f.size + 2

    # 表头
    y = y0
    h = row_heights[0]
    h2 = row_heights[1]
    total_header_h = h + h2
    draw_row_bg(y, total_header_h, HEADER_BG)
    draw_hline(y)

    text_center("项 目", 0, y, total_header_h, "bold")
    text_center("上 月", 1, y, total_header_h, "bold")
    text_center("本 月", 2, y, total_header_h, "bold")
    text_center("实 用", 3, y, total_header_h, "bold")
    text_center("单 价", 4, y, total_header_h, "bold")
    text_center_span("金    额", 5, 9, y, h, "bold")
    text_center("备注", 10, y, total_header_h, "bold")

    for i in [0, 1, 2, 3, 4, 5, 10, 11]:
        draw.line([(col_x[i], y), (col_x[i], y + total_header_h)], fill=LINE_COLOR, width=1)
    draw.line([(col_x[5], y + h), (col_x[10], y + h)], fill=LINE_COLOR, width=1)

    y += h
    h = h2
    for i, label in enumerate(['万', '千', '百', '十', '元']):
        text_center(label, 5 + i, y, h, "small")
    for i in range(6, 10):
        draw.line([(col_x[i], y), (col_x[i], y + h)], fill=LINE_COLOR, width=1)
    draw_hline(y + h)

    def draw_data_row(y, h, label, ws_row):
        draw_hline(y)
        text_center(label, 0, y, h, "normal")
        for ci, col_letter in enumerate(['B', 'C', 'D', 'E'], 1):
            val = ws[f'{col_letter}{ws_row}'].value
            if val is not None:
                text_center(str(val), ci, y, h, "normal")
        for ci, col_letter in enumerate(['F', 'G', 'H', 'I', 'J'], 5):
            val = ws[f'{col_letter}{ws_row}'].value
            if val is not None:
                text_center(str(val), ci, y, h, "amount")
        kval = ws[f'K{ws_row}'].value
        if kval is not None:
            draw_k_cell(str(kval), y, h, "small", BLACK)
        draw_vlines(y, h)

    y += h
    h = row_heights[2]
    draw_data_row(y, h, "水费(吨)", 5)

    y += h
    h = row_heights[3]
    draw_data_row(y, h, "电费(度)", 6)

    # 房租（K7 红色）
    y += h
    h = row_heights[4]
    draw_hline(y)
    text_center("房 租", 0, y, h, "normal")
    rent_period = ws['B7'].value
    if rent_period:
        text_center_span(str(rent_period), 1, 4, y, h, "normal")
    for ci, col_letter in enumerate(['F', 'G', 'H', 'I', 'J'], 5):
        val = ws[f'{col_letter}7'].value
        if val is not None:
            text_center(str(val), ci, y, h, "amount")
    k7val = ws['K7'].value
    if k7val is not None:
        draw_k_cell(str(k7val), y, h, "small", RED)
    merge_skip = {2, 3, 4}
    for i in range(len(col_x)):
        if i not in merge_skip:
            draw.line([(col_x[i], y), (col_x[i], y + h)], fill=LINE_COLOR, width=1)

    # 卫生费/网费
    y += h
    h = row_heights[5]
    draw_hline(y)
    text_center("卫生费", 0, y, h, "small")
    text_center("网 费", 2, y, h, "small")
    draw_vlines(y, h)

    # 管理费 + 其它
    y += h
    h = row_heights[6]
    draw_hline(y)
    text_center("管理费", 0, y, h, "small")
    val_b9 = ws['B9'].value
    if val_b9:
        text_center(str(val_b9), 1, y, h, "normal")
    text_center("其 它", 2, y, h, "small")
    val_d9 = ws['D9'].value
    if val_d9:
        text_center(str(val_d9), 3, y, h, "normal")
    for ci, col_letter in enumerate(['F', 'G', 'H', 'I', 'J'], 5):
        val = ws[f'{col_letter}9'].value
        if val is not None:
            text_center(str(val), ci, y, h, "amount")
    draw_vlines(y, h)

    # 合计
    y += h
    h = row_heights[7]
    draw_hline(y)
    total_text = str(ws['A10'].value or "")
    text_left(total_text, x0, y, h, "normal")
    amount_text = str(ws['I10'].value or "")
    text_right(amount_text, table_right - 4, y, h, "bold")
    draw.line([(x0, y), (x0, y + h)], fill=LINE_COLOR, width=1)
    draw.line([(table_right, y), (table_right, y + h)], fill=LINE_COLOR, width=1)
    draw_hline(y + h)

    # 开票人
    y += h
    h = row_heights[8]
    issuer = str(ws['A11'].value or "开票人：")
    text_left(issuer, x0, y, h, "normal")
    receiver = str(ws['H11'].value or "收款人：张")
    text_right(receiver, table_right, y, h, "normal")

    img.save(output_path, "JPEG", quality=95)


def main():
    ap = argparse.ArgumentParser(description="把收据 xlsx 批量渲染成 jpg")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--village", default="官田村")
    ap.add_argument("--root", default=None)
    args = ap.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 脚本在 {root}/通用/ 下，root 默认 = 父目录
    root = args.root or os.path.dirname(script_dir)

    month_dir = os.path.join(root, args.village, f"{args.year}年{args.month}月")
    dir_xlsx = os.path.join(month_dir, "收据")
    dir_img = os.path.join(month_dir, "收据图片")

    if not os.path.isdir(dir_xlsx):
        print(f"❌ 找不到收据 xlsx 目录：{dir_xlsx}（请先跑 通用/生成收据.py）", file=sys.stderr)
        sys.exit(1)

    files = sorted([f for f in os.listdir(dir_xlsx) if f.endswith('.xlsx') and not f.startswith('~$')])
    if not files:
        print(f"❌ {dir_xlsx} 下没有收据 xlsx", file=sys.stderr)
        sys.exit(1)

    os.makedirs(dir_img, exist_ok=True)
    fonts = load_fonts()

    print(f"📥 收据 xlsx：{dir_xlsx}（{len(files)} 份）")
    print(f"📤 输出 jpg：{dir_img}")
    print()

    for fname in files:
        xlsx_path = os.path.join(dir_xlsx, fname)
        img_path = os.path.join(dir_img, fname.replace('.xlsx', '.jpg'))
        wb = load_workbook(xlsx_path)
        ws = wb.active
        draw_receipt(ws, img_path, fonts)
        wb.close()
        print(f"  ✅ {fname} → {os.path.basename(img_path)}")

    print(f"\n🎉 共渲染 {len(files)} 张")
    print(f"  → 下一步：python 每月租金生成/通用/生成打印图.py --year {args.year} --month {args.month} --village {args.village}")


if __name__ == "__main__":
    main()
