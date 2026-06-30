"""变电站 6月1日应收租金生成"""
import os, re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)
os.makedirs('通用', exist_ok=True)
os.makedirs('2026年6月', exist_ok=True)

# ============= 1. 模板 =============
HEADERS = ['房号','本期租期文案','本期收款日',
           '房租','管理费','燃气费',
           '上月水表','本月水表','水表差','水费',
           '上月电表','本月电表','电表差','电费',
           '应收合计','异常','已收','备注']
WIDTHS = [10, 22, 14, 7, 8, 8, 9, 9, 7, 7, 9, 9, 7, 7, 10, 10, 7, 36]
THIN = Side(style='thin', color='999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FONT = Font(name='Microsoft YaHei', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill('solid', fgColor='4472C4')
HEADER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
BODY_FONT = Font(name='Microsoft YaHei', size=11)
BODY_ALIGN = Alignment(horizontal='center', vertical='center')
NOTE_ALIGN = Alignment(horizontal='left', vertical='center', wrap_text=True)
TOTAL_FONT = Font(name='Microsoft YaHei', size=11, bold=True, color='C00000')
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')

tpl = '通用/应收租金-模板.xlsx'
wb = Workbook()
ws = wb.active
ws.title = '应收租金'
for i, w in enumerate(WIDTHS, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
for j, h in enumerate(HEADERS, 1):
    c = ws.cell(row=1, column=j, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = HEADER_ALIGN
    c.border = BORDER
ws.row_dimensions[1].height = 32
ws.freeze_panes = 'A2'
wb.save(tpl)
print(f'✅ 模板已建：{tpl}')

# ============= 2. 数据 =============
house_wb = load_workbook('房屋说明-变电站.xlsx', data_only=True)
hws = house_wb.active
houses = {}
for row in hws.iter_rows(min_row=2, values_only=True):
    if row[0]:
        rno = str(row[0])
        houses[rno] = {
            'rent': row[3] or 0,
            'mgmt': row[4] or 0,
            'gas':  row[5] or 0,
            'period_raw': row[6] or '',
            'note': row[8] or '',
        }

prev_readings = {
    '101': (929, 6766), '102': (506, 9997), '103': (643, 6134),
    '104': (169, None), '105': (2105, 8540),
    '201': (457, 6778), '202': (556, 3379), '203': (321, 4194),
    '205': (597, 6479), '206': (181, 7252), '207': (269, 5549), '208': (756, 3091),
    '301': (265, 4898), '302': (429, 4144), '303': (322, 7515),
    '305': (293, 7220), '306': (733, 8074), '307': (353, 3939), '308': (535, 4842),
    '401': (349, 3675), '402': (673, 4373), '405': (587, 5482),
    '406': (310, 3029), '407': (534, 8254), '408': (495, 922),
    '501': (255, 7751), '502': (379, 8101), '505': (399, 8275),
    '506': (379, 8234), '507': (215, 4809), '508': (319, 6318),
    '601': (441, 5954), '602': (397, 6181), '605': (311, 278),
    '606': (357, 6282), '607': (310, 7006), '608': (353, 5062),
}
curr_readings = {
    '101': (932, 6975), '102': (508, 28), '103': (646, 6173),
    '104': (169, None), '105': (2109, 8688),
    '201': (458, 6889), '202': (558, 3398), '203': (321, 4208),
    '205': (601, 6575), '206': (182, 7266), '207': (271, 5598), '208': (758, 3120),
    '301': (267, 4925), '302': (430, 4189), '303': (327, 7568),
    '305': (294, 7272), '306': (736, 8104), '307': (355, 3964), '308': (539, 4985),
    '401': (349, 3682), '402': (682, 4548), '405': (588, 5570),
    '406': (312, 3056), '407': (542, 8441), '408': (498, 960),
    '501': (256, 7839), '502': (379, 8169), '505': (403, 8326),
    '506': (381, 8334), '507': (217, 4844), '508': (322, 6377),
    '601': (443, 6113), '602': (398, 6267), '605': (313, 412),
    '606': (357, 6369), '607': (314, 7192), '608': (355, 5147),
}

def get_price(rno):
    if rno == '408':
        return 7, 1.5
    return 6, 1

def shift_period(text):
    if not text:
        return ''
    def repl(m):
        n = int(m.group(1))
        return f'{n+1}月' if n < 12 else '1月'
    return re.sub(r'(\d+)月', repl, text)

# ============= 3. 计算 =============
rows = []
all_rooms = sorted(houses.keys(), key=lambda x: int(x))

for rno in all_rooms:
    h = houses[rno]
    period = shift_period(h['period_raw'])
    pay_date = '2026年6月1日'
    notes = []

    if rno == '104':
        # 104+105 合并，按 105 出收据：用 104 档案的房租/管理/租期，水电合并算
        p104w, p104e = prev_readings.get('104', (None, None))
        c104w, c104e = curr_readings.get('104', (None, None))
        p105w, p105e = prev_readings.get('105', (None, None))
        c105w, c105e = curr_readings.get('105', (None, None))

        # 水：104 + 105 累计差
        water_prev = (p104w or 0) + (p105w or 0)
        water_curr = (c104w or 0) + (c105w or 0)
        wd_raw = water_curr - water_prev
        if wd_raw <= 0:
            water_diff = 1
            notes.append('水表无变化，按保底1吨')
        else:
            water_diff = wd_raw
        # 电：104 无表，只用 105
        if p105e is None or c105e is None:
            elec_prev, elec_curr, elec_diff = p105e, c105e, 1
            notes.append('无电表读数，按保底1度')
        else:
            elec_prev, elec_curr = p105e, c105e
            ed_raw = c105e - p105e
            if ed_raw <= 0:
                elec_diff = 1
                notes.append('电表无变化，按保底1度')
            else:
                elec_diff = ed_raw

        wp_unit, ep_unit = 6, 1
        rent = h['rent'] or 0          # 1900
        mgmt = h['mgmt'] or 0          # 40
        gas  = h['gas']  or 0
        water_fee = water_diff * wp_unit
        elec_fee  = elec_diff * ep_unit
        total = rent + mgmt + gas + water_fee + elec_fee
        notes.insert(0, '104+105合算，按105出收据（104无电表，水表合并）')
        rows.append({
            'rno': '105',                      # 收据房号 = 105
            'period': period,
            'pay_date': pay_date,
            'rent': rent, 'mgmt': mgmt, 'gas': gas,
            'wp': water_prev, 'wc': water_curr, 'wd': water_diff, 'wf': water_fee,
            'ep': elec_prev,  'ec': elec_curr,  'ed': elec_diff,  'ef': elec_fee,
            'total': total,
            'abnormal': '是' if any('⚠' in n for n in notes) else '',
            'paid': False,
            'note': '；'.join(notes),
        })
        continue

    p = prev_readings.get(rno)
    c = curr_readings.get(rno)
    if not p or not c:
        rent = h['rent'] or 0
        mgmt = h['mgmt'] or 0
        gas  = h['gas']  or 0
        rows.append({
            'rno': rno, 'period': period, 'pay_date': pay_date,
            'rent': rent, 'mgmt': mgmt, 'gas': gas,
            'wp': None, 'wc': None, 'wd': None, 'wf': None,
            'ep': None, 'ec': None, 'ed': None, 'ef': None,
            'total': rent + mgmt + gas,
            'abnormal': '是', 'paid': False,
            'note': '⚠无水电读数',
        })
        continue

    pw, pe = p
    cw, ce = c
    wp_unit, ep_unit = get_price(rno)

    # 水
    if pw is None or cw is None:
        water_prev, water_curr = pw, cw
        water_diff = 1
        water_fee = water_diff * wp_unit
        notes.append('无水表读数，按保底1吨')
    else:
        water_prev, water_curr = pw, cw
        d = cw - pw
        if d == 0:
            water_diff = 1
            notes.append('水表无变化，按保底1吨')
        elif d < 0:
            water_diff = d
            notes.append(f'⚠水表读数异常({pw}→{cw})')
        else:
            water_diff = d
        water_fee = water_diff * wp_unit if water_diff > 0 else 0

    # 电
    if pe is None or ce is None:
        elec_prev, elec_curr = pe, ce
        elec_diff = None
        elec_fee = 0
        notes.append('无电表')
    else:
        elec_prev, elec_curr = pe, ce
        if rno == '102' and ce < pe:
            elec_diff = (ce + 10000) - pe
            notes.append(f'电表翻表({pe}→{ce}+10000)')
        else:
            d = ce - pe
            if d == 0:
                elec_diff = 1
                notes.append('电表无变化，按保底1度')
            elif d < 0:
                elec_diff = d
                notes.append(f'⚠电表读数异常({pe}→{ce})')
            else:
                elec_diff = d
        elec_fee = elec_diff * ep_unit if (elec_diff is not None and elec_diff > 0) else 0

    rent = h['rent'] or 0
    mgmt = h['mgmt'] or 0
    gas  = h['gas']  or 0
    if rno == '408':
        notes.append('单价：水7元/吨，电1.5元/度')

    total = rent + mgmt + gas + (water_fee or 0) + (elec_fee or 0)
    rows.append({
        'rno': rno, 'period': period, 'pay_date': pay_date,
        'rent': rent, 'mgmt': mgmt, 'gas': gas,
        'wp': water_prev, 'wc': water_curr, 'wd': water_diff, 'wf': water_fee,
        'ep': elec_prev,  'ec': elec_curr,  'ed': elec_diff,  'ef': elec_fee,
        'total': total,
        'abnormal': '是' if any('⚠' in n for n in notes) else '',
        'paid': False,
        'note': '；'.join(notes),
    })

# ============= 4. 写入 =============
wb = load_workbook(tpl)
ws = wb.active

def write_body(row_idx, vals, *, is_total=False):
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=row_idx, column=j, value=v)
        c.border = BORDER
        if is_total:
            c.font = TOTAL_FONT
            c.fill = TOTAL_FILL
        else:
            c.font = BODY_FONT
        if j == 18:
            c.alignment = NOTE_ALIGN
        else:
            c.alignment = BODY_ALIGN

total_rent = total_mgmt = total_gas = 0
total_water_fee = total_elec_fee = 0
total_all = 0
for i, r in enumerate(rows, start=2):
    write_body(i, [
        r['rno'], r['period'], r['pay_date'],
        r['rent'], r['mgmt'], r['gas'] or '',
        r['wp'], r['wc'], r['wd'], r['wf'],
        r['ep'], r['ec'], r['ed'], r['ef'],
        r['total'], r['abnormal'], r['paid'], r['note'],
    ])
    total_rent += r['rent'] or 0
    total_mgmt += r['mgmt'] or 0
    total_gas  += r['gas']  or 0
    total_water_fee += r['wf'] or 0
    total_elec_fee  += r['ef'] or 0
    total_all += r['total'] or 0

total_row = 2 + len(rows)
write_body(total_row, [
    '合计', '', '',
    total_rent, total_mgmt, total_gas,
    '', '', '', total_water_fee,
    '', '', '', total_elec_fee,
    total_all, '', '', f'共 {len(rows)} 户',
], is_total=True)

out = '2026年6月/2026年6月应收租金.xlsx'
wb.save(out)
print(f'✅ 应收租金已生成：{out}')
print(f'   共 {len(rows)} 行，应收合计 ¥{total_all:,.2f}')
print()
print(f'{"房号":<10}{"租金":>6}{"水":>5}{"电":>5}{"水费":>7}{"电费":>7}{"合计":>8}  备注')
for r in rows:
    print(f'{r["rno"]:<10}{r["rent"]:>6}{(r["wd"] or 0):>5}{(r["ed"] or 0):>5}'
          f'{(r["wf"] or 0):>7.1f}{(r["ef"] or 0):>7.1f}{r["total"]:>8.1f}  {r["note"]}')
