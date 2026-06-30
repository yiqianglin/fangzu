"""
变电站 2026年6月1日收据批量生成
- 数据来源：2026年6月/2026年6月应收租金.xlsx
- 收据模板：通用/收据模板.xlsx
- 输出：
  · 收据 xlsx → 2026年6月/收据/
  · 收据图片 jpg → 2026年6月/收据图片/

特殊说明：
- 408 单价：水 7 元/吨、电 1.5 元/度（其他都是水6/电1）
- 105 行实际是 104+105 合算（按 105 出收据）
- 房租文案前不再追加"房租"二字（应收里已写好）
"""
import os
import shutil
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

# ==================== 路径 ====================
ROOT = "/Users/linyiqiang/workspace/tencent/专利相关/每月租金生成/变电站"
COMMON = f"{ROOT}/通用"
TEMPLATE_RECEIPT = f"{COMMON}/收据模板.xlsx"
DIR_M6 = f"{ROOT}/2026年6月"
RECEIVABLE_XLSX = f"{DIR_M6}/2026年6月应收租金.xlsx"
DIR_RECEIPT = f"{DIR_M6}/收据"
DIR_RECEIPT_IMG = f"{DIR_M6}/收据图片"
for d in (DIR_RECEIPT, DIR_RECEIPT_IMG):
    os.makedirs(d, exist_ok=True)


# ==================== 工具函数 ====================
def num_to_chinese(num):
    digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    units = ['', '拾', '佰', '仟', '万']
    if num == 0:
        return '零元整'
    result = ''
    s = str(int(num))
    length = len(s)
    for i, d in enumerate(s):
        digit = int(d)
        unit_idx = length - 1 - i
        if digit != 0:
            result += digits[digit] + units[unit_idx]
        else:
            if result and not result.endswith('零'):
                result += '零'
    result = result.rstrip('零')
    result += '元整'
    return result


def split_amount_digits(amount):
    """将整数金额拆到 万/千/百/十/元 五列（F~J），高位空填 None。"""
    amt = int(amount)
    result = {'F': None, 'G': None, 'H': None, 'I': None, 'J': None}
    if amt >= 10000:
        result['F'] = (amt // 10000) % 10
    if amt >= 1000:
        result['G'] = (amt // 1000) % 10
    if amt >= 100:
        result['H'] = (amt // 100) % 10
    if amt >= 10:
        result['I'] = (amt // 10) % 10
    result['J'] = amt % 10
    return result


# ==================== Step 1：从应收 xlsx 读数据 ====================
wb = load_workbook(RECEIVABLE_XLSX, data_only=True)
ws = wb.active
# 字段顺序：房号|本期租期文案|本期收款日|房租|管理费|燃气费|上月水|本月水|水差|水费
#         |上月电|本月电|电差|电费|应收合计|异常|已收|备注
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[0] or str(r[0]) == '合计':
        continue
    rows.append({
        'rno': str(r[0]),
        'period': r[1] or '',
        'pay_date': r[2] or '',
        'rent': r[3] or 0,
        'mgmt': r[4] or 0,
        'gas': r[5] or 0,
        'wp': r[6], 'wc': r[7], 'wd': r[8] if r[8] is not None else 0,
        'wf': r[9] or 0,
        'ep': r[10], 'ec': r[11], 'ed': r[12] if r[12] is not None else 0,
        'ef': r[13] or 0,
        'total': r[14] or 0,
        'note': r[17] or '',
    })
wb.close()
print(f"📊 从应收 xlsx 读取 {len(rows)} 户")

# 读"房屋说明-变电站.xlsx" 拿租客姓名（最后一列）
TENANT = {}
_hwb = load_workbook('房屋说明-变电站.xlsx', data_only=True)
_hws = _hwb.active
_headers = [c.value for c in _hws[1]]
_name_idx = _headers.index('租客姓名') if '租客姓名' in _headers else len(_headers) - 1
for r in _hws.iter_rows(min_row=2, values_only=True):
    if r and r[0] is not None:
        TENANT[str(r[0])] = (r[_name_idx] or '').strip() if isinstance(r[_name_idx], str) else (r[_name_idx] or '')
_hwb.close()


# ==================== Step 2：生成收据 xlsx ====================
def get_water_price(rno):
    return 7 if rno == '408' else 6

def get_elec_price(rno):
    return 1.5 if rno == '408' else 1

def fmt_price(p):
    """单价显示：整数就整数，小数就保留一位"""
    return int(p) if float(p).is_integer() else p

xlsx_paths = []
for d in rows:
    rno = d['rno']
    out_xlsx = f"{DIR_RECEIPT}/{rno}房_6月收据.xlsx"
    shutil.copy2(TEMPLATE_RECEIPT, out_xlsx)
    rwb = load_workbook(out_xlsx)
    rws = rwb.active
    rws.title = f"{rno}房6月收据"

    _name = TENANT.get(rno, '')
    # 105 行实际涵盖 104+105，按 105 出收据，姓名取 104 的（104 才是真正的租客）
    if rno == '105' and not _name:
        _name = TENANT.get('104', '')
    rws['A2'] = f'用户名称：    {rno}房 {_name}' if _name else f'用户名称：    {rno}房'
    rws['H2'] = d['pay_date']

    # 水费行
    rws['B5'] = d['wp']
    rws['C5'] = d['wc']
    rws['D5'] = d['wd']
    rws['E5'] = fmt_price(get_water_price(rno))
    for col, v in split_amount_digits(d['wf']).items():
        rws[f'{col}5'] = v
    # 408 把"特殊单价"备注写进 K5
    if rno == '408':
        rws['K5'] = '7元/吨'

    # 电费行
    rws['B6'] = d['ep']
    rws['C6'] = d['ec']
    rws['D6'] = d['ed']
    rws['E6'] = fmt_price(get_elec_price(rno))
    for col, v in split_amount_digits(int(d['ef'])).items():
        rws[f'{col}6'] = v
    if rno == '408':
        rws['K6'] = '1.5元/度'

    # 房租行
    rws['B7'] = f"{d['period']}房租" if d['period'] else "房租"
    for col, v in split_amount_digits(d['rent']).items():
        rws[f'{col}7'] = v

    # 卫生费/网费 留空
    for col in 'FGHIJ':
        rws[f'{col}8'] = None

    # 管理费 + 其它（燃气）
    rws['B9'] = d['mgmt'] if d['mgmt'] > 0 else None
    rws['D9'] = d['gas'] if d['gas'] > 0 else None
    row9 = (d['mgmt'] or 0) + (d['gas'] or 0)
    if row9 > 0:
        for col, v in split_amount_digits(row9).items():
            rws[f'{col}9'] = v
    else:
        for col in 'FGHIJ':
            rws[f'{col}9'] = None

    # 合计
    rws['A10'] = f'合计   人民币（大写）{num_to_chinese(d["total"])}'
    rws['I10'] = f'(¥：{int(d["total"])}元)'

    rwb.save(out_xlsx)
    xlsx_paths.append((rno, out_xlsx, d))
    print(f"  ✅ {rno}房 - 水{d['wf']:.0f} 电{d['ef']:.0f} 租{d['rent']} 管{d['mgmt']} 合计 ¥{d['total']:.0f}")


# ==================== Step 3：渲染收据图片 ====================
FONT_PATH = "/System/Library/Fonts/STHeiti Medium.ttc"
IMG_WIDTH = 720
MARGIN_LEFT = 12
MARGIN_TOP = 8
BLACK = (0, 0, 0)
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

    def text_center(text, col_idx, y, h, font_key="normal"):
        text = str(text) if text is not None else ""
        if not text:
            return
        f = fonts[font_key]
        tw = draw.textlength(text, font=f)
        tx = col_x[col_idx] + (col_widths[col_idx] - tw) / 2
        ty = y + (h - f.size) / 2 - 1
        draw.text((tx, ty), text, fill=BLACK, font=f)

    def text_center_span(text, col_start, col_end, y, h, font_key="normal"):
        text = str(text) if text is not None else ""
        if not text:
            return
        f = fonts[font_key]
        tw = draw.textlength(text, font=f)
        span_left = col_x[col_start]
        span_right = col_x[col_end] + col_widths[col_end]
        tx = span_left + (span_right - span_left - tw) / 2
        ty = y + (h - f.size) / 2 - 1
        draw.text((tx, ty), text, fill=BLACK, font=f)

    def text_left(text, x, y, h, font_key="normal"):
        f = fonts[font_key]
        ty = y + (h - f.size) / 2 - 1
        draw.text((x + 4, ty), str(text), fill=BLACK, font=f)

    def text_right(text, x_right, y, h, font_key="normal"):
        f = fonts[font_key]
        tw = draw.textlength(str(text), font=f)
        ty = y + (h - f.size) / 2 - 1
        draw.text((x_right - tw - 4, ty), str(text), fill=BLACK, font=f)

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
        val = ws[f'K{ws_row}'].value
        if val is not None:
            text_center(str(val), 10, y, h, "small")
        draw_vlines(y, h)

    y += h
    h = row_heights[2]
    draw_data_row(y, h, "水费(吨)", 5)

    y += h
    h = row_heights[3]
    draw_data_row(y, h, "电费(度)", 6)

    # 房租
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


print("\n🖼️  渲染收据图片：")
fonts = load_fonts()
for rno, xlsx_path, _ in xlsx_paths:
    img_name = f"{rno}房_6月收据.jpg"
    img_path = f"{DIR_RECEIPT_IMG}/{img_name}"
    rwb = load_workbook(xlsx_path)
    rws = rwb.active
    draw_receipt(rws, img_path, fonts)
    rwb.close()
    print(f"  ✅ {img_name}")

print(f"\n🎉 全部完成！")
print(f"  - 收据 xlsx：{DIR_RECEIPT}（{len(xlsx_paths)} 份）")
print(f"  - 收据图片：{DIR_RECEIPT_IMG}（{len(xlsx_paths)} 张）")
