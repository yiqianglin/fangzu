"""
2026年6月1日收据 - 本地模式（仅 Step 4 + Step 5）

前置：
- 应收租金已存：每月租金生成/官田村/2026年6月/2026年6月应收租金.xlsx（含合计行）
- 房屋说明：每月租金生成/通用/房屋说明-官田村.xlsx
- 收据模板：每月租金生成/通用/收据模板.xlsx

⚠️ 历史遗留脚本（v1 时代一次性产物）—— 新流程一律走 `每月租金生成/通用/生成收据.py`，本脚本仅留作存档参考。

按 skill 铁律：
- A2 用户名称 = `{房号}房  {租客姓名}`（两个空格 + 姓名）
- K7 下月预告（红色）= `下月（7月）房租改为{下月房租}元`
  ※ 仅当「房屋说明.下月房租」列**有值**时才写；为空则 K7 留空
- K9 / K5 / K6 / K10 一律不写下月预告
"""
import os
import shutil
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

# ==================== 路径 ====================
ROOT = "/Users/linyiqiang/workspace/tencent/专利相关/每月租金生成/官田村"
COMMON = f"{ROOT}/通用"
TEMPLATE_HOUSE = f"{COMMON}/房屋说明-官田村.xlsx"
TEMPLATE_RECEIPT = f"{COMMON}/收据模板.xlsx"

DIR_M6 = f"{ROOT}/2026年6月"
RECEIVABLE_XLSX = f"{DIR_M6}/2026年6月应收租金.xlsx"
DIR_RECEIPT = f"{DIR_M6}/收据"
DIR_RECEIPT_IMG = f"{DIR_M6}/收据图片"

for d in (DIR_RECEIPT, DIR_RECEIPT_IMG):
    os.makedirs(d, exist_ok=True)

WATER_PRICE_DEFAULT = 6
ELEC_PRICE = 1
TARGET_MONTH = 6
NEXT_MONTH = TARGET_MONTH + 1 if TARGET_MONTH < 12 else 1


# ==================== 工具函数 ====================
def num_to_chinese(num):
    digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    units = ['', '拾', '佰', '仟', '万']
    if num == 0:
        return '零元整'
    result = ''
    s = str(int(num))
    L = len(s)
    for i, d in enumerate(s):
        digit = int(d)
        unit_idx = L - 1 - i
        if digit != 0:
            result += digits[digit] + units[unit_idx]
        else:
            if result and not result.endswith('零'):
                result += '零'
    result = result.rstrip('零')
    return result + '元整'


def split_amount_digits(amount):
    amt = int(amount)
    res = {'F': None, 'G': None, 'H': None, 'I': None, 'J': None}
    if amt >= 10000:
        res['F'] = (amt // 10000) % 10
    if amt >= 1000:
        res['G'] = (amt // 1000) % 10
    if amt >= 100:
        res['H'] = (amt // 100) % 10
    if amt >= 10:
        res['I'] = (amt // 10) % 10
    res['J'] = amt % 10
    return res


# ==================== 加载房屋说明（取姓名 + 下月房租 + 水单价） ====================
hwb = load_workbook(TEMPLATE_HOUSE, data_only=True)
hws = hwb.active
HOUSE = {}
hdr = [c.value for c in hws[1]]
idx = {name: hdr.index(name) for name in hdr if name}
for row in hws.iter_rows(min_row=2, values_only=True):
    if not row or not row[0]:
        continue
    rno = str(row[0]).strip()
    HOUSE[rno] = {
        "name": (row[idx["租客姓名"]] or "").strip() if idx.get("租客姓名") is not None else "",
        "rent": row[idx["房租"]] or 0,
        "next_rent": row[idx["下月房租"]] if idx.get("下月房租") is not None else None,
        "water_price": row[idx["水单价"]] if idx.get("水单价") is not None else None,
    }


# ==================== 读 6 月应收租金 ====================
rwb = load_workbook(RECEIVABLE_XLSX, data_only=True)
rws = rwb.active
rcv_hdr = [c.value for c in rws[1]]
ridx = {name: rcv_hdr.index(name) for name in rcv_hdr if name}

receivable_rows = []
for row in rws.iter_rows(min_row=2, values_only=True):
    if not row or not row[0]:
        continue
    if str(row[0]).strip() == "合计":
        continue
    receivable_rows.append(row)


# ==================== Step 4：批量生成收据 xlsx ====================
def gen_receipt(data):
    rno = str(data[ridx["房号"]]).strip()
    period = data[ridx["本期租期文案"]]
    pay_date = data[ridx["本期收款日"]]
    rent = data[ridx["房租"]] or 0
    mgmt = data[ridx["管理费"]] or 0
    gas = data[ridx["燃气费"]] or 0
    pw = data[ridx["上月水表"]]
    cw = data[ridx["本月水表"]]
    dw = data[ridx["水表差"]]
    fw = data[ridx["水费"]] or 0
    pe = data[ridx["上月电表"]]
    ce = data[ridx["本月电表"]]
    de = data[ridx["电表差"]]
    fe = data[ridx["电费"]] or 0
    total = data[ridx["应收合计"]] or 0

    h = HOUSE.get(rno, {})
    name = h.get("name", "")
    next_rent = h.get("next_rent")  # 仅当房屋说明里显式有「下月房租」才不为空
    water_price = h.get("water_price") or WATER_PRICE_DEFAULT

    out_xlsx = f"{DIR_RECEIPT}/{rno}房_6月收据.xlsx"
    shutil.copy2(TEMPLATE_RECEIPT, out_xlsx)
    wb = load_workbook(out_xlsx)
    ws = wb.active
    ws.title = f"{rno}房6月收据"

    # 用户名称：房号 + 两个空格 + 姓名（铁律）
    if name:
        ws['A2'] = f'用户名称：    {rno}房  {name}'
    else:
        ws['A2'] = f'用户名称：    {rno}房'
        print(f"  ⚠️  {rno}房：房屋说明未填租客姓名，请补全")
    ws['H2'] = pay_date

    # 水费
    ws['B5'] = pw
    ws['C5'] = cw
    ws['D5'] = dw
    ws['E5'] = water_price
    for col, v in split_amount_digits(fw).items():
        ws[f'{col}5'] = v

    # 电费
    ws['B6'] = pe
    ws['C6'] = ce
    ws['D6'] = de
    ws['E6'] = ELEC_PRICE
    for col, v in split_amount_digits(fe).items():
        ws[f'{col}6'] = v

    # 房租
    ws['B7'] = f"{period}房租"
    for col, v in split_amount_digits(rent).items():
        ws[f'{col}7'] = v

    # K7 下月预告（仅当房屋说明.下月房租有值时才写，红色，固定模板）
    if next_rent is not None and str(next_rent).strip() != "":
        ws['K7'] = f"下月（{NEXT_MONTH}月）房租改为{int(next_rent)}元"
    else:
        ws['K7'] = None

    # 卫生费/网费 留空
    for col in 'FGHIJ':
        ws[f'{col}8'] = None

    # 管理费 + 其它（燃气）
    ws['B9'] = mgmt if mgmt > 0 else None
    ws['D9'] = gas if gas > 0 else None
    row9 = (mgmt or 0) + (gas or 0)
    if row9 > 0:
        for col, v in split_amount_digits(row9).items():
            ws[f'{col}9'] = v
    else:
        for col in 'FGHIJ':
            ws[f'{col}9'] = None
    # K9 不写下月预告（铁律）

    # 合计
    ws['A10'] = f'合计   人民币（大写）{num_to_chinese(total)}'
    ws['I10'] = f'(¥：{total}元)'

    wb.save(out_xlsx)
    return out_xlsx, rno, fw, fe, rent, mgmt, gas, total


print("📝 生成收据 xlsx：")
results = []
for r in receivable_rows:
    info = gen_receipt(r)
    results.append(info)
    p, rno, fw, fe, rent, mgmt, gas, total = info
    print(f"  ✅ {rno}房 - 水{fw} 电{fe} 租{rent} 管{mgmt} 燃{gas} 合计 ¥{total}")


# ==================== Step 5：渲染收据图片 ====================
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
        """K 列（备注列），col_idx=10，自动换行（最多 2 行），居中。"""
        if not text:
            return
        f = fonts[font_key]
        cell_w = col_widths[10] - 4
        # 简单按字符宽度估算分行
        words = list(str(text))
        lines, cur = [], ""
        for ch in words:
            if draw.textlength(cur + ch, font=f) > cell_w:
                lines.append(cur)
                cur = ch
                if len(lines) >= 1:
                    # 第二行收尾
                    pass
            else:
                cur += ch
        if cur:
            lines.append(cur)
        if len(lines) > 2:
            lines = lines[:2]
            # 末尾省略
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
        # K 列普通行黑色
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
    # K7：红色下月预告
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
    # K9 不写下月预告（铁律）
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


print("\n🖼️  渲染收据图片：")
fonts = load_fonts()
for info in results:
    xlsx_path = info[0]
    fname = os.path.basename(xlsx_path)
    img_name = fname.replace('.xlsx', '.jpg')
    img_path = f"{DIR_RECEIPT_IMG}/{img_name}"
    wb = load_workbook(xlsx_path)
    ws = wb.active
    draw_receipt(ws, img_path, fonts)
    wb.close()
    print(f"  ✅ {img_name}")

# ==================== 汇总 ====================
sum_total = sum(x[7] for x in results)
sum_rent = sum(x[4] for x in results)
sum_mgmt = sum(x[5] for x in results)
sum_gas = sum(x[6] for x in results)
sum_w = sum(x[2] for x in results)
sum_e = sum(x[3] for x in results)
print(f"\n🎉 全部完成！")
print(f"  - 收据 xlsx：{DIR_RECEIPT}（{len(results)} 份）")
print(f"  - 收据图片：{DIR_RECEIPT_IMG}（{len(results)} 张）")
print(f"  - 总应收：¥{sum_total}（房租 ¥{sum_rent} / 管理 ¥{sum_mgmt} / 燃气 ¥{sum_gas} / 水费 ¥{sum_w} / 电费 ¥{sum_e}）")
