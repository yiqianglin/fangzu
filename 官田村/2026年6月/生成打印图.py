"""
将 收据图片/ 里的 jpg 按 A4 纸张排版（一行一张），输出到 打印/ 目录。

布局：
- A4 纵向 @300dpi = 2480x3508 px
- 边距 100 px
- 每页 3 张收据，宽度铺满（2280 px），高度按原比例缩放
- 收据之间留间距
"""
import os
from PIL import Image

SRC = '/Users/linyiqiang/workspace/tencent/专利相关/每月租金生成/官田村/2026年6月/收据图片'
DST = '/Users/linyiqiang/workspace/tencent/专利相关/每月租金生成/官田村/2026年6月/打印'

# A4 @300dpi
A4_W, A4_H = 2480, 3508
MARGIN = 100
PER_PAGE = 3

os.makedirs(DST, exist_ok=True)

# 收据按房号顺序排序
files = sorted(
    [f for f in os.listdir(SRC) if f.lower().endswith('.jpg')],
    key=lambda x: int(x.split('房')[0])
)
print(f'共 {len(files)} 张收据')

# 缩放后单张尺寸
target_w = A4_W - 2 * MARGIN  # 2280
sample = Image.open(os.path.join(SRC, files[0]))
ratio = sample.size[1] / sample.size[0]
target_h = int(target_w * ratio)  # 约 995
print(f'每张缩放后: {target_w} x {target_h}')

# 间距：把 3 张垂直居中分布
total_h = target_h * PER_PAGE
free = (A4_H - 2 * MARGIN) - total_h
gap = free // (PER_PAGE - 1) if PER_PAGE > 1 else 0
print(f'每页 {PER_PAGE} 张, 间距 {gap} px')

pages = [files[i:i+PER_PAGE] for i in range(0, len(files), PER_PAGE)]
print(f'共 {len(pages)} 页')

for idx, page_files in enumerate(pages, 1):
    canvas = Image.new('RGB', (A4_W, A4_H), 'white')
    for i, fname in enumerate(page_files):
        im = Image.open(os.path.join(SRC, fname)).convert('RGB')
        im = im.resize((target_w, target_h), Image.LANCZOS)
        x = MARGIN
        y = MARGIN + i * (target_h + gap)
        canvas.paste(im, (x, y))
    out = os.path.join(DST, f'打印_第{idx}页.jpg')
    canvas.save(out, 'JPEG', quality=92, dpi=(300, 300))
    print(f'  -> {out}  收据: {[f.split("房")[0]+"房" for f in page_files]}')

print('完成。')
